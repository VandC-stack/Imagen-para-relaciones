"""Gestión persistente y atómica del contador de folios.

Proporciona funciones seguras para reservar folios (uno o en bloque),
consultar y fijar el último folio. Usa un archivo JSON en `data/folio_counter.json`
y un lock por archivo `data/folio_counter.lock` para evitar condiciones de carrera
entre procesos.
"""
from __future__ import annotations
import os
import json
import time
from typing import Tuple


def _get_paths() -> Tuple[str, str]:
    base = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "folio_counter.json"), os.path.join(base, "folio_counter.lock")


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
