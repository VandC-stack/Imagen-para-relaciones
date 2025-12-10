import os
import json
import io
from typing import List, Dict

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DATA_DIR = os.path.join(PROJECT_ROOT, 'data')
EVIDENCE_CONFIG = os.path.join(DATA_DIR, 'evidence_paths.json')

os.makedirs(DATA_DIR, exist_ok=True)


def _load_config():
    try:
        with open(EVIDENCE_CONFIG, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(cfg: dict):
    try:
        with open(EVIDENCE_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def save_paths_for_group(paths: List[str]) -> bool:
    cfg = _load_config()
    abs_paths = [os.path.abspath(p) for p in paths]
    cfg['bosch'] = abs_paths
    return _save_config(cfg)


def get_saved_paths() -> List[str]:
    cfg = _load_config()
    return cfg.get('bosch', [])


def find_evidence_for_codes(codes: List[str], base_paths: List[str] = None) -> Dict[str, List[str]]:
    """
    Búsqueda para BOSCH. Similar a grupo_axo: no recursivo.
    """
    results = {str(c): [] for c in codes}
    if base_paths is None:
        base_paths = get_saved_paths()

    exts = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}

    for p in base_paths:
        if not os.path.isdir(p):
            continue
        try:
            files = os.listdir(p)
        except Exception:
            continue

        for code in codes:
            code_str = str(code)
            # exact match
            for fn in files:
                name, ext = os.path.splitext(fn)
                if ext.lower() not in exts:
                    continue
                if name == code_str:
                    results[code_str].append(os.path.join(p, fn))

            # substring
            if not results[code_str]:
                for fn in files:
                    if code_str in fn:
                        name, ext = os.path.splitext(fn)
                        if ext.lower() in exts:
                            results[code_str].append(os.path.join(p, fn))

    # unique
    for k, v in results.items():
        seen = set(); unique = []
        for item in v:
            if item not in seen:
                seen.add(item); unique.append(item)
        results[k] = unique

    return results


def load_image_bytes(path: str) -> io.BytesIO:
    buf = io.BytesIO()
    with open(path, 'rb') as f:
        buf.write(f.read())
    buf.seek(0)
    return buf
