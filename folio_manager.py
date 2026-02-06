"""Gestión persistente y atómica del contador de folios.

Proporciona funciones seguras para reservar folios (uno o en bloque),
consultar y fijar el último folio. Usa un archivo JSON en `data/folio_counter.json`
y un lock por archivo `data/folio_counter.lock` para evitar condiciones de carrera
entre procesos.
"""
from __future__ import annotations
import os
import sys
import json
import time
from typing import Tuple


def _get_paths() -> Tuple[str, str]:
    """Return (counter_path, lock_path).

    Lookup order:
    1. Directory from FOLIO_DATA_DIR env var (if set)
    2. When frozen, APP_DIR/data where APP_DIR = dirname(sys.executable)
    3. Ascend from current working directory looking for a `data` folder
    4. Fallback to package-local `data` next to this module
    """
    candidates = []
    # 1) env override
    env_dir = os.environ.get("FOLIO_DATA_DIR") or os.environ.get('IMAGENESVC_DATA_DIR')
    if env_dir:
        candidates.append(env_dir)

    # 2) when frozen, prefer folder next to the exe and also its _internal sibling
    try:
        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
            candidates.append(os.path.join(exe_dir, "data"))
            candidates.append(os.path.join(exe_dir, "_internal", "data"))
    except Exception:
        pass

    # 3) consider BASE_DIR (sys._MEIPASS) and its _internal/data (when frozen)
    try:
        base_dir = getattr(sys, '_MEIPASS', None)
        if base_dir:
            candidates.append(os.path.join(base_dir, 'data'))
            candidates.append(os.path.join(base_dir, '_internal', 'data'))
    except Exception:
        pass

    # 4) ascend from cwd and look for a data folder
    try:
        cwd = os.path.abspath(os.getcwd())
        parts = cwd.split(os.path.sep)
        for i in range(len(parts), 0, -1):
            base = os.path.sep.join(parts[:i])
            candidates.append(os.path.join(base, "data"))
    except Exception:
        pass

    # 5) package-local data folder
    candidates.append(os.path.join(os.path.dirname(__file__), "data"))

    # Normalize candidates: only keep unique, existing or creatable paths
    seen = set()
    norm_cands = []
    for c in candidates:
        if not c:
            continue
        c_abs = os.path.abspath(c)
        if c_abs in seen:
            continue
        seen.add(c_abs)
        norm_cands.append(c_abs)

    # Prefer candidate that contains expected data files (non-empty), score them
    want_files = ["folio_counter.json", "historial_visitas.json", "Clientes.json"]
    best = None
    best_score = -1
    # detect exe_dir to prefer the data folder next to the executable when frozen
    exe_dir = None
    try:
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
    except Exception:
        exe_dir = None

    for cand in norm_cands:
        try:
            if not os.path.exists(cand):
                # try to create folder so file checks below can run
                os.makedirs(cand, exist_ok=True)
        except Exception:
            continue
        score = 0
        for wf in want_files:
            p = os.path.join(cand, wf)
            try:
                if os.path.exists(p) and os.path.isfile(p) and os.path.getsize(p) > 0:
                    score += 1
            except Exception:
                pass
        # small bonus if the folder is writable
        try:
            if os.access(cand, os.W_OK):
                score += 0.1
        except Exception:
            pass
        # big preference if this candidate is the folder next to the exe (dist\...\data)
        try:
            if exe_dir:
                exe_data = os.path.abspath(os.path.join(exe_dir, 'data'))
                if os.path.abspath(cand) == exe_data:
                    score += 100
        except Exception:
            pass
        if score > best_score:
            best_score = score
            best = cand

    chosen = None
    if best is not None and best_score >= 0:
        chosen = best

    # If we couldn't pick by score, fallback to first writable candidate
    if not chosen:
        for cand in norm_cands:
            try:
                if not os.path.exists(cand):
                    os.makedirs(cand, exist_ok=True)
                if os.access(cand, os.W_OK):
                    chosen = cand
                    break
            except Exception:
                continue

    # final fallback: package-local
    if not chosen:
        base = os.path.join(os.path.dirname(__file__), "data")
        try:
            os.makedirs(base, exist_ok=True)
        except Exception:
            pass
        chosen = base

    # Write debug note about chosen candidate (best effort)
    try:
        dbg = os.path.join(chosen, 'startup_exe_debug.log')
        with open(dbg, 'a', encoding='utf-8') as _f:
            _f.write(f"folio_manager.chosen={chosen}\n")
    except Exception:
        pass

    counter = os.path.join(chosen, "folio_counter.json")
    lock = os.path.join(chosen, "folio_counter.lock")
    return counter, lock


def _acquire_lock(lock_path: str, timeout: float = 5.0, poll: float = 0.05) -> bool:
    start = time.time()
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return True
        except FileExistsError:
            if (time.time() - start) >= timeout:
                return False
            time.sleep(poll)


def _release_lock(lock_path: str) -> None:
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def _read_counter(counter_path: str) -> int:
    try:
        if os.path.exists(counter_path):
            with open(counter_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return int(data.get("last", 0))
    except Exception:
        pass
    return 0


def _write_counter(counter_path: str, value: int) -> None:
    tmp = counter_path + ".tmp"
    data = {"last": int(value)}
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f)
    try:
        os.replace(tmp, counter_path)
    except Exception:
        try:
            if os.path.exists(counter_path):
                os.remove(counter_path)
        except Exception:
            pass
        os.replace(tmp, counter_path)


def reserve_next(timeout: float = 5.0) -> int:
    """Reserva y devuelve el siguiente folio (entero).

    Adquiere lock, lee el último folio, incrementa en 1, lo persiste y devuelve
    el nuevo valor.
    """
    counter_path, lock_path = _get_paths()
    ok = _acquire_lock(lock_path, timeout=timeout)
    if not ok:
        raise TimeoutError("No se pudo adquirir el lock del contador de folios")
    try:
        last = _read_counter(counter_path)
        nuevo = int(last) + 1
        _write_counter(counter_path, nuevo)
        return nuevo
    finally:
        _release_lock(lock_path)


def reserve_block(count: int, timeout: float = 5.0) -> int:
    """Reserva un bloque de `count` folios y devuelve el primer folio del bloque."""
    if count <= 0:
        raise ValueError("count debe ser > 0")
    counter_path, lock_path = _get_paths()
    ok = _acquire_lock(lock_path, timeout=timeout)
    if not ok:
        raise TimeoutError("No se pudo adquirir el lock del contador de folios")
    try:
        last = _read_counter(counter_path)
        start = int(last) + 1
        nuevo = int(last) + int(count)
        _write_counter(counter_path, nuevo)
        return start
    finally:
        _release_lock(lock_path)


def get_last() -> int:
    """Devuelve el último folio persistido (0 si no existe)."""
    counter_path, _ = _get_paths()
    return _read_counter(counter_path)


def set_last(value: int, timeout: float = 5.0) -> None:
    """Fija el último folio a `value` (usa lock)."""
    counter_path, lock_path = _get_paths()
    ok = _acquire_lock(lock_path, timeout=timeout)
    if not ok:
        raise TimeoutError("No se pudo adquirir el lock del contador de folios")
    try:
        _write_counter(counter_path, int(value))
    finally:
        _release_lock(lock_path)


def format_folio(n: int, width: int = 6) -> str:
    return str(int(n)).zfill(width)


__all__ = [
    "reserve_next",
    "reserve_block",
    "get_last",
    "set_last",
    "format_folio",
]
