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
def _sin_acentos(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def _es_encabezado_codigo(texto: str) -> bool:
    """Detecta encabezado de la columna CODIGO (tolerante a acentos/espacios)."""
    t = _sin_acentos(texto or '').upper()
    t = re.sub(r'\s+', '', t)
    if any(pal in t for pal in ['FACTURA', 'CANTIDAD', 'MARCA', 'TOTAL']):
        return False
    return 'CODIG' in t  # CODIGO / CÓDIGO / CODIGOS

def normalizar_cadena_alnum_mayus(s: str) -> str:
    """Quita todo lo que no sea alfanumérico y pasa a MAYUS (para comparaciones robustas)."""
    return re.sub(r'[^A-Za-z0-9]', '', s or '').upper()

def buscar_imagen(carpeta_imagenes: str, codigo_canonico: str) -> str | None:
    """
    Busca la imagen correspondiente al código dentro de la carpeta (sin subcarpetas).
    Ignora mayúsculas/minúsculas tanto en el nombre del archivo como en la extensión.
    """
    codigo_canonico = (codigo_canonico or '').upper()
    try:
        for nombre in os.listdir(carpeta_imagenes):
            base, ext = os.path.splitext(nombre)
            if ext.lower() in ['.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp']:
                if base.upper() == codigo_canonico:
                    return os.path.join(carpeta_imagenes, nombre)
    except Exception as e:
        print(f"Error al leer la carpeta: {carpeta_imagenes} -> {e}")
    return None

# ----------------------------------------
# Extracción de códigos desde tablas Word (solo columna CODIGO)
# ----------------------------------------
def extraer_codigos(doc: Document) -> list[str]:
    """
    Extrae los códigos válidos de la columna CODIGO de las tablas,
    tolerando mayúsculas/minúsculas y letras al inicio del código.
    Requiere al menos una letra (evita facturas/cantidades) y longitud razonable.
    """
    codigos: list[str] = []
    patron_codigo = re.compile(r"[A-Za-z0-9][A-Za-z0-9.\-]{4,}", re.IGNORECASE)

    for tabla in doc.tables:
        if not tabla.rows:
            continue

        # Detectar índice de columna CODIGO
        idx_codigo = None
        for r in range(min(3, len(tabla.rows))):
            for i, celda in enumerate(tabla.rows[r].cells):
                if _es_encabezado_codigo(celda.text):
                    idx_codigo = i
                    break
            if idx_codigo is not None:
                break
        if idx_codigo is None:
            continue

        # Recorrer solo esa columna
        for fila in tabla.rows[1:]:
            if idx_codigo >= len(fila.cells):
                continue
            texto = fila.cells[idx_codigo].text or ''
            if not texto.strip():
                continue

            # Capturar candidatos en la celda
            for m in patron_codigo.findall(texto):
                canon = normalizar_cadena_alnum_mayus(m)
                if not any(c.isalpha() for c in canon):  # al menos una letra
                    continue
                if 5 <= len(canon) <= 20:
                    codigos.append(canon)

    # Quitar duplicados manteniendo orden
    return list(dict.fromkeys(codigos))

# ----------------------------------------
# Inserción de imágenes
# ----------------------------------------
def _insert_paragraph_after_table(table) -> Paragraph:
    """
    Inserta un párrafo inmediatamente después de 'table' y retorna el Paragraph.
    """
    new_p = OxmlElement("w:p")
    table._element.addnext(new_p)
    return Paragraph(new_p, table._element.getparent())

def insertar_imagenes_en_docx(ruta_doc: str):
    print(f"Procesando documento: {ruta_doc}")
    doc = Document(ruta_doc)
    carpeta_imagenes = obtener_ruta_carpeta()
    if not carpeta_imagenes:
        print("No se pudo obtener una carpeta válida para las imágenes.")
        return

    # 1) Obtener códigos SOLO con la función especializada
    codigos = extraer_codigos(doc)
    if not codigos:
        print("No se encontraron códigos válidos en las tablas.")
        return
    else:
        print(f"Códigos detectados: {', '.join(codigos)}")

    # Parámetros visuales
    ancho_pagina = 6.0
    espacio_entre = 0.15
    num_imgs = len(codigos)
    ancho_imagen = max(1.2, (ancho_pagina - espacio_entre * (num_imgs - 1)) / max(1, num_imgs))

    # 2) Si existe ${etiqueta1}, insertar ahí
    etiqueta_encontrada = False
    for p in doc.paragraphs:
        if "${etiqueta1}" in (p.text or ""):
            etiqueta_encontrada = True
            p.clear()
            run = p.add_run()
            print(f"Inserción múltiple de imágenes ({num_imgs}) en etiqueta1.")
            run.add_text("\n")  # pequeño desplazamiento vertical

            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen(carpeta_imagenes, codigo)
                if img_path:
                    try:
                        run.add_picture(img_path, width=Inches(ancho_imagen))
                        if i < num_imgs:
                            run.add_text("   ")
                        print(f"Imagen insertada: {os.path.basename(img_path)}")
                    except Exception as e:
                        print(f"Error al insertar imagen {img_path}: {e}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
            break

    # 3) Si no hay etiqueta, insertar inmediatamente debajo de la tabla que tiene los códigos
    if not etiqueta_encontrada:
        print("No se encontró ${etiqueta1}. Insertando imágenes debajo de la tabla con los códigos.")
        cod_set = set(codigos)
        tabla_objetivo = None

        for tabla in doc.tables:
            # ¿Esta tabla contiene alguno de los códigos?
            contiene = False
            for fila in tabla.rows:
                for celda in fila.cells:
                    cel_norm = normalizar_cadena_alnum_mayus(celda.text)
                    if any(c in cel_norm for c in cod_set):
                        contiene = True
                        break
                if contiene:
                    break
            if contiene:
                tabla_objetivo = tabla
                break

        if tabla_objetivo is not None:
            p = _insert_paragraph_after_table(tabla_objetivo)
            p.alignment = 1  # centrado
            run = p.add_run()
            run.add_text("\n")  # bajarlo un poco

            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen(carpeta_imagenes, codigo)
                if img_path:
                    try:
                        run.add_picture(img_path, width=Inches(ancho_imagen))
                        if i < num_imgs:
                            run.add_text("   ")
                        print(f"Imagen insertada (debajo de tabla): {os.path.basename(img_path)}")
                    except Exception as e:
                        print(f"Error al insertar imagen {img_path}: {e}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
        else:
            print("No se encontró ninguna tabla que contenga los códigos.")

    doc.save(ruta_doc)
    print(f"Documento actualizado: {ruta_doc}\n")

# ----------------------------------------
# Procesamiento por lotes
# ----------------------------------------
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
