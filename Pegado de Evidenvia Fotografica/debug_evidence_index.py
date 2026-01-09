import os, json, re, sys

def _normalizar(s):
    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

IMG_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}

def construir_indice(evidence_cfg_path):
    try:
        with open(evidence_cfg_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f) or {}
    except Exception as e:
        print("No pude leer", evidence_cfg_path, ":", e)
        return {}
    index = {}
    for grp, lst in (cfg or {}).items():
        for carpeta in lst:
            for root, _, files in os.walk(carpeta):
                for nombre in files:
                    base, ext = os.path.splitext(nombre)
                    if ext.lower() not in IMG_EXTS:
                        continue
                    core = re.sub(r"[\s\-_]*\(\s*\d+\s*\)$", "", base)
                    core = re.sub(r"[\s\-_]+\d+$", "", core)
                    key = _normalizar(core)
                    if not key:
                        continue
                    index.setdefault(key, []).append(os.path.join(root, nombre))
    return index

if __name__ == '__main__':
    code = (sys.argv[1] if len(sys.argv) > 1 else '').strip()
    cfg_path = os.path.join('data', 'evidence_paths.json')
    print("Usando config:", cfg_path)
    idx = construir_indice(cfg_path)
    print("Claves indexadas (tot):", len(idx))
    sample = list(idx.keys())[:50]
    print("Muestra claves:", sample)
    if code:
        k = _normalizar(code)
        print("Buscando clave normalizada:", k)
        found = idx.get(k)
        if found:
            print("Rutas encontradas para", code, ":")
            for p in found:
                print(" -", p)
        else:
            print("No hay entradas exactas para", code)
            # show near matches
            near = [kk for kk in idx.keys() if k in kk or kk in k]
            if near:
                print("Coincidencias parciales en índice:")
                for kk in near[:30]:
                    print(" *", kk, "->", idx.get(kk)[:5])
            else:
                print("No se encontraron coincidencias parciales tampoco.")