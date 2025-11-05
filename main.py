import os
import json
import sys
import re
from tkinter import Tk, filedialog
from docx import Document
from docx.shared import Inches

# ----------------------------------------
# Configuración general
# ----------------------------------------
sys.stdout.reconfigure(encoding='utf-8')
CONFIG_FILE = os.path.abspath("config.json")


# ----------------------------------------
# Gestión de carpeta y configuración
# ----------------------------------------
def seleccionar_carpeta():
    """Permite seleccionar carpeta y corrige nombres con espacios o rutas similares."""
    Tk().withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de imágenes")
    if not carpeta:
        return None

    carpeta = carpeta.replace("\\", "/").strip()
    carpeta = " ".join(carpeta.split())  # colapsa espacios múltiples

    # Si no existe exactamente, intenta encontrar una carpeta similar
    if not os.path.isdir(carpeta):
        base_dir = os.path.dirname(carpeta)
        nombre = os.path.basename(carpeta).replace(" ", "").lower()
        for item in os.listdir(base_dir):
            real = os.path.join(base_dir, item)
            if os.path.isdir(real) and item.replace(" ", "").lower() == nombre:
                carpeta = real
                break

    if not os.path.isdir(carpeta):
        print(f"No se encontró físicamente la carpeta seleccionada: {carpeta}")
        return None

    carpeta = os.path.normpath(carpeta)
    guardar_ruta(carpeta)
    return carpeta


def guardar_ruta(ruta):
    """Guarda la carpeta seleccionada en config.json."""
    data = {"ruta_imagenes": ruta}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def obtener_ruta_carpeta():
    """Obtiene la carpeta desde config.json o la solicita si no es válida."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                ruta = data.get("ruta_imagenes", "")
                if os.path.isdir(ruta):
                    return ruta
                else:
                    print(f"La ruta guardada no existe o fue movida: {ruta}")
                    return seleccionar_carpeta()
        except Exception as e:
            print(f"Error al leer config.json: {e}")
            return seleccionar_carpeta()
    else:
        return seleccionar_carpeta()


# ----------------------------------------
# Utilidades
# ----------------------------------------
def normalizar_codigo(codigo: str) -> str:
    """Convierte un código como 0.611.2A7.5G0 -> 06112A75G0"""
    return ''.join(filter(str.isalnum, codigo))


def buscar_imagen(carpeta: str, codigo: str):
    """
    Busca la imagen del código de forma tolerante:
    - Ignora mayúsculas y minúsculas
    - Ignora espacios invisibles
    - Acepta extensiones PNG, JPG, JPEG
    """
    codigo_limpio = codigo.strip().lower().replace(" ", "")
    try:
        archivos = os.listdir(carpeta)
    except Exception as e:
        print(f"Error al leer la carpeta: {carpeta} -> {e}")
        return None

    # Coincidencia exacta (tolerante a espacios y mayúsculas)
    for nombre in archivos:
        base, ext = os.path.splitext(nombre)
        base_limpio = base.strip().lower().replace(" ", "")
        if base_limpio == codigo_limpio and ext.lower() in [".png", ".jpg", ".jpeg"]:
            ruta = os.path.join(carpeta, nombre)
            print(f"Imagen encontrada: {ruta}")
            return ruta

    # Coincidencia parcial (por si hay variaciones mínimas)
    for nombre in archivos:
        base, ext = os.path.splitext(nombre)
        base_limpio = base.strip().lower().replace(" ", "")
        if codigo_limpio in base_limpio and ext.lower() in [".png", ".jpg", ".jpeg"]:
            print(f"Coincidencia parcial encontrada: {nombre}")
            return os.path.join(carpeta, nombre)

    print(f"No se encontró imagen para el código {codigo} en {carpeta}")
    return None


# ----------------------------------------
# Extracción de códigos desde tablas Word
# ----------------------------------------
def extraer_codigos(doc: Document):
    """Extrae códigos del tipo 0.611.2A7.5G0 desde las tablas."""
    codigos = []
    patron_codigo = re.compile(r"\d+\.\d+\.\w+\.\w+", re.IGNORECASE)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto = celda.text.strip()
                if patron_codigo.match(texto):
                    codigos.append(normalizar_codigo(texto))
    return list(dict.fromkeys(codigos))  # elimina duplicados preservando orden


# ----------------------------------------
# Lógica principal
# ----------------------------------------
def insertar_imagenes_en_docx(ruta_doc: str):
    print(f"Procesando documento: {ruta_doc}")
    doc = Document(ruta_doc)
    carpeta_imagenes = obtener_ruta_carpeta()
    if not carpeta_imagenes:
        print("No se pudo obtener una carpeta válida para las imágenes.")
        return

    codigos = extraer_codigos(doc)
    if not codigos:
        print("No se encontraron códigos válidos en las tablas.")
        return
    else:
        print(f"Códigos detectados: {', '.join(codigos)}")

    for p in doc.paragraphs:
        for i, codigo in enumerate(codigos, start=1):
            etiqueta = f"${{etiqueta{i}}}"
            if etiqueta in p.text:
                img_path = buscar_imagen(carpeta_imagenes, codigo)
                if img_path:
                    p.clear()
                    run = p.add_run()
                    try:
                        run.add_picture(img_path, width=Inches(2.5))
                        print(f"Imagen insertada: {os.path.basename(img_path)}")
                    except Exception as e:
                        print(f"Error al insertar imagen {img_path}: {e}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")

    doc.save(ruta_doc)
    print(f"Documento actualizado: {ruta_doc}\n")


def procesar_lote(carpeta_docs="docs"):
    """Procesa todos los documentos .docx dentro de la carpeta especificada."""
    if not os.path.exists(carpeta_docs):
        print(f"La carpeta '{carpeta_docs}' no existe. Créala y coloca los documentos dentro.")
        return

    archivos = [f for f in os.listdir(carpeta_docs) if f.endswith(".docx")]
    if not archivos:
        print(f"No se encontraron archivos .docx en '{carpeta_docs}'.")
        return

    print(f"Se encontraron {len(archivos)} documentos en '{carpeta_docs}'. Iniciando procesamiento...\n")
    for archivo in archivos:
        ruta_doc = os.path.join(carpeta_docs, archivo)
        insertar_imagenes_en_docx(ruta_doc)

    print("Procesamiento por lotes completado.")


# ----------------------------------------
# Ejecución principal
# ----------------------------------------
if __name__ == "__main__":
    procesar_lote()
