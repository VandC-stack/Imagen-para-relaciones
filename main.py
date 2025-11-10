import os
import json
import sys
import re
import unicodedata
from tkinter import Tk, filedialog
from docx import Document
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement

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
# CONFIGURACIÓN Y ENTRADA/SALIDA
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
# UTILIDADES DE RUTA / PERSISTENCIA
# ------------------------------------------------------------
def resolver_onedrive_path(ruta):
    ruta = os.path.abspath(ruta).replace("\\", "/")
    if "OneDrive" in ruta:
        userprofile = os.environ.get("USERPROFILE", "")
        onedrive_local = os.path.join(userprofile, "OneDrive").replace("\\", "/")
        if ruta.startswith(onedrive_local):
            try:
                if os.path.exists(ruta):
                    return ruta
            except Exception:
                pass
            ruta_alt = ruta.replace("  ", " ")
            if os.path.exists(ruta_alt):
                return ruta_alt
    return ruta

def seleccionar_carpeta():
    Tk().withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de imágenes")
    if not carpeta:
        return None
    carpeta = carpeta.replace("\\", "/").strip()
    if not os.path.isdir(carpeta):
        print(f"No se encontró la carpeta: {carpeta}")
        return None
    carpeta = resolver_onedrive_path(carpeta)
    guardar_ruta(carpeta)
    return carpeta

def guardar_ruta(ruta):
    ruta = resolver_onedrive_path(ruta)
    data = {"ruta_imagenes": ruta}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print(f"Carpeta de imágenes guardada correctamente: {ruta}")

def obtener_ruta_carpeta():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                ruta = (data.get("ruta_imagenes") or "").strip()
                ruta = resolver_onedrive_path(ruta)
                print(f"Ruta cargada desde config.json: {ruta}")
                if os.path.isdir(ruta):
                    return ruta
                print(f"La ruta guardada no existe o no está descargada localmente: {ruta}")
        except Exception as e:
            print(f"Error al leer config.json: {e}")
    return seleccionar_carpeta()

def validar_carpeta_imagenes(carpeta_imagenes: str) -> int:
    try:
        archivos = os.listdir(carpeta_imagenes)
    except Exception as e:
        print(f"Error al listar '{carpeta_imagenes}': {e}")
        return 0
    n = sum(1 for f in archivos if os.path.splitext(f)[1].lower() in IMG_EXTS)
    print(f"Imágenes detectadas en la carpeta: {n} (ruta: {carpeta_imagenes})")
    return n

# ------------------------------------------------------------
# NORMALIZACIÓN DE TEXTO / CÓDIGOS
# ------------------------------------------------------------
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
    try:
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
    except Exception as e:
        print(f"Error al indexar imágenes en '{carpeta_imagenes}': {e}")
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
def extraer_codigos(doc: Document) -> list[str]:
    codigos: list[str] = []
    patron_general = re.compile(r"[A-Za-z0-9][A-Za-z0-9.\-]{4,}", re.IGNORECASE)
    patron_bosch = re.compile(r"(?:No\.?\s*)?(\d(?:\s?\d){8,12})")

    def limpiar(s: str) -> str:
        return (s or "").replace("\xa0", " ").replace("\n", " ").strip()

    for tabla in doc.tables:
        if not tabla.rows:
            continue
        idx_codigo = None
        for r in range(min(3, len(tabla.rows))):
            for i, celda in enumerate(tabla.rows[r].cells):
                t = limpiar(celda.text)
                t_norm = _sin_acentos(t).upper()
                t_norm = re.sub(r"\s+", " ", t_norm)
                if ("CODIGO" in t_norm or "CODIG" in t_norm or "SKU" in t_norm or "CLAVE" in t_norm) and not any(
                    x in t_norm for x in ("FACTURA", "CANTIDAD", "TOTAL", "MARCA")
                ):
                    idx_codigo = i
                    break
            if idx_codigo is not None:
                break

        columnas_a_revisar = [idx_codigo] if idx_codigo is not None else None

        for fila in tabla.rows[1:]:
            celdas = fila.cells
            indices = range(len(celdas)) if columnas_a_revisar is None else columnas_a_revisar
            for j in indices:
                if j >= len(celdas):
                    continue
                texto = limpiar(celdas[j].text)
                if not texto:
                    continue
                for m in patron_bosch.findall(texto):
                    num = re.sub(r"\s+", "", m)
                    if 8 <= len(num) <= 13:
                        codigos.append(num)
                for m in patron_general.findall(texto):
                    canon = normalizar_cadena_alnum_mayus(m)
                    if not canon:
                        continue
                    if not contiene_digito(canon):
                        continue
                    if canon in FORBIDDEN_TOKENS:
                        continue
                    if canon.isdigit():
                        if 8 <= len(canon) <= 13:
                            codigos.append(canon)
                    else:
                        if 5 <= len(canon) <= 24:
                            codigos.append(canon)
    vistos = set()
    resultado = []
    for c in codigos:
        if c not in vistos:
            vistos.add(c)
            resultado.append(c)
    return resultado

# ------------------------------------------------------------
# INSERCIÓN DE IMÁGENES
# ------------------------------------------------------------
def insertar_imagen_con_transparencia(run, img_path, ancho):
    run.add_text("\n" * 2)
    ancho_reducido = min(ancho * 0.9, 4.5)
    run.add_picture(img_path, width=Inches(ancho_reducido))
    run.add_text("\n" * 2)

def _insert_paragraph_after_table(doc: Document, table):
    tbl_element = table._element
    parent = tbl_element.getparent()
    idx = parent.index(tbl_element)
    p = doc.add_paragraph()
    parent.insert(idx + 1, p._element)
    return p

# ------------------------------------------------------------
# LÓGICA PRINCIPAL
# ------------------------------------------------------------
def insertar_imagenes_en_docx(ruta_doc: str):
    print(f"Procesando documento: {ruta_doc}")
    doc = Document(ruta_doc)
    carpeta_imagenes = obtener_ruta_carpeta()
    if not carpeta_imagenes:
        print("No se pudo obtener una carpeta válida para las imágenes.")
        if registrar_fallo:
            registrar_fallo(os.path.basename(ruta_doc))
        return
    validar_carpeta_imagenes(carpeta_imagenes)
    index = indexar_imagenes(carpeta_imagenes)
    codigos = extraer_codigos(doc)
    if not codigos:
        print("No se encontraron códigos válidos en las tablas.")
        if registrar_fallo:
            registrar_fallo(os.path.basename(ruta_doc))
        return
    else:
        print(f"Códigos detectados: {', '.join(codigos)}")
    ancho_pagina = 6.0
    espacio_entre = 0.15
    num_imgs = len(codigos)
    ancho_imagen = max(1.2, (ancho_pagina - espacio_entre * (num_imgs - 1)) / max(1, num_imgs))
    etiqueta_encontrada = False
    imagen_insertada = False
    usadas_paths = set()
    usadas_bases = set()
    for p in doc.paragraphs:
        if "${etiqueta1}" in (p.text or ""):
            etiqueta_encontrada = True
            p.clear()
            run = p.add_run()
            print(f"Inserción múltiple de imágenes ({num_imgs}) en etiqueta1.")
            run.add_text("\n" * 4)
            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen_index(index, codigo, usadas_paths, usadas_bases)
                if img_path:
                    usadas_paths.add(norm_path_key(img_path))
                    base_norm_img = normalizar_cadena_alnum_mayus(os.path.splitext(os.path.basename(img_path))[0])
                    usadas_bases.add(base_norm_img)
                    insertar_imagen_con_transparencia(run, img_path, ancho_imagen)
                    if i < num_imgs:
                        run.add_text("   ")
                    imagen_insertada = True
                    print(f"Imagen insertada: {os.path.basename(img_path)}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
            break
    if not etiqueta_encontrada:
        print("No se encontró ${etiqueta1}. Buscando tabla correcta para insertar las imágenes...")
        cod_set = set(normalizar_cadena_alnum_mayus(c) for c in codigos)
        tabla_objetivo = None
        for tabla in doc.tables:
            contiene_codigo = False
            for fila in tabla.rows:
                for celda in fila.cells:
                    cel_norm = normalizar_cadena_alnum_mayus(celda.text)
                    if any(c in cel_norm for c in cod_set):
                        contiene_codigo = True
                        break
                if contiene_codigo:
                    break
            if contiene_codigo:
                tabla_objetivo = tabla
                break
        if tabla_objetivo is not None:
            p = _insert_paragraph_after_table(doc, tabla_objetivo)
            p.alignment = 1
            p.paragraph_format.space_before = Inches(1.0)
            run = p.add_run()
            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen_index(index, codigo, usadas_paths, usadas_bases)
                if img_path:
                    usadas_paths.add(norm_path_key(img_path))
                    base_norm_img = normalizar_cadena_alnum_mayus(os.path.splitext(os.path.basename(img_path))[0])
                    usadas_bases.add(base_norm_img)
                    insertar_imagen_con_transparencia(run, img_path, ancho_imagen)
                    if i < num_imgs:
                        run.add_text("   ")
                    imagen_insertada = True
                    print(f"Imagen insertada (debajo de tabla): {os.path.basename(img_path)}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
        else:
            print("No se encontró ninguna tabla adecuada para insertar las imágenes.")
    if not imagen_insertada and registrar_fallo:
        registrar_fallo(os.path.basename(ruta_doc))
    doc.save(ruta_doc)
    print(f"Documento actualizado: {ruta_doc}\n")

# ------------------------------------------------------------
# PROCESAMIENTO POR LOTES
# ------------------------------------------------------------
def procesar_lote(carpeta_docs="docs"):
    if limpiar_registro:
        limpiar_registro()
    if not os.path.exists(carpeta_docs):
        print(f"La carpeta '{carpeta_docs}' no existe. Créala y coloca los documentos dentro.")
        return
    archivos = [
        f for f in os.listdir(carpeta_docs)
        if f.endswith(".docx") and not f.startswith("~$")
    ]
    if not archivos:
        print(f"No se encontraron archivos .docx en '{carpeta_docs}'.")
        return
    print(f"Se encontraron {len(archivos)} documentos en '{carpeta_docs}'. Iniciando procesamiento...\n")
    for archivo in archivos:
        ruta_doc = os.path.join(carpeta_docs, archivo)
        insertar_imagenes_en_docx(ruta_doc)
    print("Procesamiento por lotes completado.\n")
    if mostrar_registro:
        mostrar_registro()
    log_file = os.path.abspath("documentos_sin_imagenes.txt")
    if os.path.exists(log_file):
        os.startfile(log_file)

# ------------------------------------------------------------
# EJECUCIÓN
# ------------------------------------------------------------
if __name__ == "__main__":
    procesar_lote()