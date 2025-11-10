import os
import json
import sys
import re
import unicodedata
from tkinter import Tk, filedialog
from docx import Document
from docx.shared import Inches, Cm
from docx.text.paragraph import Paragraph
from PIL import Image

# ------------------------------------------------------------
# MÓDULO DE REGISTRO DE FALLOS
# ------------------------------------------------------------
try:
    from registro_fallos import registrar_fallo, limpiar_registro, mostrar_registro
except Exception as e:
    print(f"Error al importar registro_fallos: {e}")
    registrar_fallo = None
    limpiar_registro = None
    mostrar_registro = None

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------
sys.stdout.reconfigure(encoding="utf-8")
CONFIG_FILE = os.path.abspath("config.json")

FORBIDDEN_TOKENS = {
    "TOTAL", "CANTIDAD", "FACTURA", "MARCA", "DESCRIPCION", "DESCRIPCIÓN",
    "FECHA", "CONTRATO", "PRESENTACION", "PRESENTACIÓN", "SISTEMA",
    "ALEATORIO", "DICTAMEN", "PRODUCTO", "RELACION", "RELACIÓN",
    "MODELO", "ORIGEN", "CHINA", "MALASIA", "ALEMANIA", "RUMANIA",
    "ITALIA", "BRASIL", "DIMENSIONES", "CONTENIDO", "ETIQUETA", "CONTEN",
}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

# ------------------------------------------------------------
# FUNCIONES DE UTILIDAD
# ------------------------------------------------------------
def resolver_onedrive_path(ruta):
    ruta = os.path.abspath(ruta).replace("\\", "/")
    return ruta

def seleccionar_carpeta():
    Tk().withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de imágenes")
    if not carpeta:
        return None
    carpeta = carpeta.replace("\\", "/").strip()
    guardar_ruta(carpeta)
    return carpeta

def guardar_ruta(ruta):
    data = {"ruta_imagenes": ruta}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print(f"Carpeta guardada: {ruta}")

def obtener_ruta_carpeta():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            ruta = (data.get("ruta_imagenes") or "").strip()
            if os.path.isdir(ruta):
                return ruta
        except Exception as e:
            print(f"Error al leer config.json: {e}")
    return seleccionar_carpeta()

def validar_carpeta_imagenes(carpeta_imagenes: str):
    try:
        archivos = os.listdir(carpeta_imagenes)
    except Exception as e:
        print(f"Error al listar '{carpeta_imagenes}': {e}")
        return 0
    n = sum(1 for f in archivos if os.path.splitext(f)[1].lower() in IMG_EXTS)
    print(f"Imágenes detectadas: {n} en {carpeta_imagenes}")
    return n

def _sin_acentos(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def normalizar_cadena_alnum_mayus(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s or "").upper()

def contiene_digito(s: str) -> bool:
    return any(c.isdigit() for c in s or "")

# ------------------------------------------------------------
# INDEXADO DE IMÁGENES
# ------------------------------------------------------------
def indexar_imagenes(carpeta_imagenes: str):
    index = []
    for nombre in os.listdir(carpeta_imagenes):
        base, ext = os.path.splitext(nombre)
        if ext.lower() not in IMG_EXTS:
            continue
        path = os.path.join(carpeta_imagenes, nombre)
        index.append({
            "name": nombre,
            "base": base,
            "ext": ext,
            "base_norm": normalizar_cadena_alnum_mayus(base),
            "path": path,
        })
    return index

def norm_path_key(path: str) -> str:
    return os.path.normcase(os.path.normpath(path or ""))

def buscar_imagen_index(index, codigo_canonico: str, usadas_paths: set, usadas_bases: set) -> str | None:
    code = normalizar_cadena_alnum_mayus(codigo_canonico)
    if not code:
        return None
    exactos = [
        it for it in index
        if it["base_norm"] == code
        and norm_path_key(it["path"]) not in usadas_paths
        and it["base_norm"] not in usadas_bases
    ]
    if exactos:
        return exactos[0]["path"]
    parciales = [
        it for it in index
        if (code in it["base_norm"] or it["base_norm"] in code)
        and norm_path_key(it["path"]) not in usadas_paths
        and it["base_norm"] not in usadas_bases
    ]
    if not parciales:
        return None
    def score(it):
        bn = it["base_norm"]
        starts = bn.startswith(code)
        ends = bn.endswith(code)
        delta = abs(len(bn) - len(code))
        return (0 if starts or ends else 1, delta, bn)
    parciales.sort(key=score)
    return parciales[0]["path"]

# ------------------------------------------------------------
# EXTRACCIÓN DE CÓDIGOS
# ------------------------------------------------------------
def extraer_codigos(doc: Document):
    codigos = []
    patron_general = re.compile(r"[A-Za-z0-9][A-Za-z0-9.\-]{4,}", re.IGNORECASE)
    patron_bosch = re.compile(r"(?:No\.?\s*)?(\d(?:\s?\d){8,12})")

    for tabla in doc.tables:
        if not tabla.rows:
            continue
        idx_codigo = None
        for r in range(min(3, len(tabla.rows))):
            for i, celda in enumerate(tabla.rows[r].cells):
                t_norm = _sin_acentos(celda.text or "").upper()
                if ("CODIGO" in t_norm or "SKU" in t_norm or "CLAVE" in t_norm) and not any(
                    x in t_norm for x in ("FACTURA", "CANTIDAD", "TOTAL", "MARCA")
                ):
                    idx_codigo = i
                    break
            if idx_codigo is not None:
                break

        columnas = [idx_codigo] if idx_codigo is not None else range(len(tabla.rows[0].cells))
        for fila in tabla.rows[1:]:
            for j in columnas:
                texto = (fila.cells[j].text or "").strip()
                for m in patron_bosch.findall(texto):
                    num = re.sub(r"\s+", "", m)
                    if 8 <= len(num) <= 13:
                        codigos.append(num)
                for m in patron_general.findall(texto):
                    canon = normalizar_cadena_alnum_mayus(m)
                    if canon and contiene_digito(canon) and canon not in FORBIDDEN_TOKENS:
                        codigos.append(canon)
    return list(dict.fromkeys(codigos))

# ------------------------------------------------------------
# INSERCIÓN DE IMÁGENES CON TAMAÑO CONTROLADO
# ------------------------------------------------------------
H_MAX_W_CM = 4.36
H_MAX_H_CM = 6.37
V_MAX_W_CM = 8.13
V_MAX_H_CM = 4.84

def insertar_imagen_con_transparencia(run, img_path):
    """
    Inserta imágenes con límite de tamaño exacto por orientación.
    Usa solo una dimensión (width o height) para evitar desaparición en verticales.
    """
    try:
        with Image.open(img_path) as img:
            w_px, h_px = img.size
        w_in, h_in = w_px / 96.0, h_px / 96.0

        if w_px >= h_px:  # horizontal
            max_w_in = H_MAX_W_CM / 2.54
            max_h_in = H_MAX_H_CM / 2.54
        else:              # vertical
            max_w_in = V_MAX_W_CM / 2.54
            max_h_in = V_MAX_H_CM / 2.54

        width_factor = max_w_in / w_in
        height_factor = max_h_in / h_in
        scale = min(width_factor, height_factor, 1.0)
        new_w_in = w_in * scale
        new_h_in = h_in * scale

        if width_factor <= height_factor:
            run.add_picture(img_path, width=Inches(new_w_in))
        else:
            run.add_picture(img_path, height=Inches(new_h_in))

        run.add_text(" ")  # espacio en línea
    except Exception as e:
        print(f"Error al insertar {img_path}: {e}")

# ------------------------------------------------------------
# LÓGICA PRINCIPAL
# ------------------------------------------------------------
def insertar_imagenes_en_docx(ruta_doc):
    print(f"Procesando documento: {ruta_doc}")
    doc = Document(ruta_doc)
    carpeta = obtener_ruta_carpeta()
    if not carpeta:
        if registrar_fallo:
            registrar_fallo(os.path.basename(ruta_doc))
        return
    validar_carpeta_imagenes(carpeta)
    index = indexar_imagenes(carpeta)
    codigos = extraer_codigos(doc)
    if not codigos:
        if registrar_fallo:
            registrar_fallo(os.path.basename(ruta_doc))
        return
    else:
        print("Códigos detectados:", ", ".join(codigos))

    usadas_paths, usadas_bases = set(), set()
    imagen_insertada = False

    for p in doc.paragraphs:
        if "${etiqueta1}" in (p.text or ""):
            p.clear()
            run = p.add_run()
            for codigo in codigos:
                img_path = buscar_imagen_index(index, codigo, usadas_paths, usadas_bases)
                if img_path:
                    usadas_paths.add(norm_path_key(img_path))
                    usadas_bases.add(normalizar_cadena_alnum_mayus(os.path.splitext(os.path.basename(img_path))[0]))
                    insertar_imagen_con_transparencia(run, img_path)
                    imagen_insertada = True
                    print(f"Imagen insertada: {os.path.basename(img_path)}")
                else:
                    print(f"No se encontró imagen para {codigo}")
            break

    if not imagen_insertada and registrar_fallo:
        registrar_fallo(os.path.basename(ruta_doc))

    # Limpieza de etiquetas sobrantes
    etiqueta_pat = re.compile(r"\$\{etiqueta\d+\}", re.IGNORECASE)
    for p in list(doc.paragraphs):
        if etiqueta_pat.search(p.text or ""):
            try:
                p._element.getparent().remove(p._element)
            except Exception:
                pass

    doc.save(ruta_doc)
    print(f"Documento actualizado: {ruta_doc}\n")

# ------------------------------------------------------------
# PROCESAMIENTO POR LOTES
# ------------------------------------------------------------
def procesar_lote(carpeta_docs="docs"):
    if limpiar_registro:
        limpiar_registro()
    if not os.path.exists(carpeta_docs):
        print(f"No existe la carpeta '{carpeta_docs}'.")
        return
    archivos = [f for f in os.listdir(carpeta_docs) if f.endswith(".docx") and not f.startswith("~$")]
    if not archivos:
        print("No se encontraron archivos .docx.")
        return

    print(f"Procesando {len(archivos)} documentos...\n")
    for archivo in archivos:
        ruta = os.path.join(carpeta_docs, archivo)
        insertar_imagenes_en_docx(ruta)

    print("Procesamiento completado.\n")
    if mostrar_registro:
        mostrar_registro()
    log = os.path.abspath("documentos_sin_imagenes.txt")
    if os.path.exists(log):
        os.startfile(log)

# ------------------------------------------------------------
# EJECUCIÓN
# ------------------------------------------------------------
if __name__ == "__main__":
    procesar_lote()