import os
import json
import re
import io
from typing import List, Dict

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DATA_DIR = os.path.join(PROJECT_ROOT, 'data')
EVIDENCE_CONFIG = os.path.join(DATA_DIR, 'evidence_paths.json')

os.makedirs(DATA_DIR, exist_ok=True)


def _load_config():
    if not os.path.exists(EVIDENCE_CONFIG):
        return {}
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


def _normalize_variants(code: str) -> List[str]:
    """Return a set of normalized variants for robust matching."""
    if code is None:
        return []
    s = str(code)
    variants = {s}
    # remove dots
    variants.add(s.replace('.', ''))
    # remove underscores
    variants.add(s.replace('_', ''))
    # remove dots and underscores
    variants.add(s.replace('.', '').replace('_', ''))
    # uppercase and lowercase
    variants.update({v.upper() for v in list(variants)})
    variants.update({v.lower() for v in list(variants)})
    return list(variants)


def save_paths_for_group(paths: List[str]) -> bool:
    """Save base paths for GRUPO_AXO inside data/evidence_paths.json"""
    cfg = _load_config()
    cfg.setdefault('grupo_axo', [])
    # store unique absolute paths
    abs_paths = [os.path.abspath(p) for p in paths]
    cfg['grupo_axo'] = abs_paths
    return _save_config(cfg)


def get_saved_paths() -> List[str]:
    cfg = _load_config()
    return cfg.get('grupo_axo', [])


def find_evidence_for_codes(codes: List[str], base_paths: List[str] = None) -> Dict[str, List[str]]:
    """
    Buscar imágenes para cada código en `codes` dentro de `base_paths`.
    Matching strategy: exact filename (without extension) -> substring -> normalized variants.
    For GRUPO_AXO we search non-recursively in each base path.

    Returns a dict: code -> list of file paths found (may be empty).
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
            # exact match on filename without extension
            for fn in files:
                name, ext = os.path.splitext(fn)
                if ext.lower() not in exts:
                    continue
                if name == code_str:
                    results[code_str].append(os.path.join(p, fn))

            # substring match
            if not results[code_str]:
                for fn in files:
                    if any(ch.isalnum() for ch in fn) and code_str in fn:
                        name, ext = os.path.splitext(fn)
                        if ext.lower() in exts:
                            results[code_str].append(os.path.join(p, fn))

            # normalized variants
            if not results[code_str]:
                variants = _normalize_variants(code_str)
                for fn in files:
                    name, ext = os.path.splitext(fn)
                    if ext.lower() not in exts:
                        continue
                    nm = name.replace('.', '').replace('_', '')
                    for v in variants:
                        if v and v.replace('.', '').replace('_', '') == nm:
                            results[code_str].append(os.path.join(p, fn))
                            break

    # remove duplicates while preserving order
    for k, v in results.items():
        seen = set()
        unique = []
        for item in v:
            if item not in seen:
                seen.add(item)
                unique.append(item)
        results[k] = unique

    return results


def load_image_bytes(path: str) -> io.BytesIO:
    """Return image bytes as BytesIO for insertion in PDFs. Raises if file not found."""
    buf = io.BytesIO()
    with open(path, 'rb') as f:
        buf.write(f.read())
    buf.seek(0)
    return buf
