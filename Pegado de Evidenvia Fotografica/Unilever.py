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
    cfg['unilever'] = abs_paths
    return _save_config(cfg)


def get_saved_paths() -> List[str]:
    cfg = _load_config()
    return cfg.get('unilever', [])


def _normalize(code: str) -> str:
    # keep punctuation but trim spaces
    return str(code).strip()


def _matches(code: str, filename: str) -> bool:
    """Match preserving punctuation: try exact on filename (w/ punctuation), then substring.
    For Unilever we prefer to keep punctuation as user requested.
    """
    name, _ = os.path.splitext(filename)
    if name == code:
        return True
    if code in filename:
        return True
    # try remove spaces
    if code.replace(' ', '') == name.replace(' ', ''):
        return True
    return False


def find_evidence_for_codes(codes: List[str], base_paths: List[str] = None) -> Dict[str, List[str]]:
    """
    Unilever: search recursively through multiple base paths.
    Matching preserves punctuation (dots, underscores, etc.) and attempts exact filename and substring.
    """
    results = {str(c): [] for c in codes}
    if base_paths is None:
        base_paths = get_saved_paths()

    exts = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}

    for base in base_paths:
        if not os.path.isdir(base):
            continue
        for root, _, files in os.walk(base):
            for fn in files:
                _, ext = os.path.splitext(fn)
                if ext.lower() not in exts:
                    continue
                for code in codes:
                    code_str = _normalize(str(code))
                    try:
                        if _matches(code_str, fn):
                            results[code_str].append(os.path.join(root, fn))
                    except Exception:
                        continue

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
