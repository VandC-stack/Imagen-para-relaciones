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
# CONFIGURACIÓN Y ENTRADA/SALIDA
# ------------------------------------------------------------
sys.stdout.reconfigure(encoding='utf-8')
CONFIG_FILE = os.path.abspath("config.json")

# ------------------------------------------------------------
# MANEJO DE CONFIGURACIÓN
# ------------------------------------------------------------
def seleccionar_carpeta():
    """Permite seleccionar carpeta y corrige nombres con espacios."""
    Tk().withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de imágenes")
    if not carpeta:
        return None
    carpeta = carpeta.replace("\\", "/").strip()
    carpeta = " ".join(carpeta.split())
    if not os.path.isdir(carpeta):
        print(f"No se encontró la carpeta: {carpeta}")
        return None
    carpeta = os.path.normpath(carpeta)
    guardar_ruta(carpeta)
    return carpeta

def guardar_ruta(ruta):
    """Guarda la carpeta en config.json"""
    data = {"ruta_imagenes": ruta}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def obtener_ruta_carpeta():
    """Obtiene la carpeta desde config.json o la solicita."""
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

# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------
def _sin_acentos(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def _es_encabezado_codigo(texto: str) -> bool:
    """Detecta encabezado de la columna CODIGO de manera más precisa."""
    t = _sin_acentos(texto or '').upper().strip()
    t = re.sub(r'\s+', ' ', t)  # Normalizar espacios
    
    # Palabras clave que definitivamente identifican "CODIGO"
    palabras_codigo = ['CODIGO', 'CODIG', 'CLAVE', 'SKU', 'REFERENCIA']
    
    # Excluir otras columnas comunes
    excluir = ['FACTURA', 'CANTIDAD', 'MARCA', 'TOTAL', 'PRECIO', 'DESCRIPCION']
    
    return (any(pal in t for pal in palabras_codigo) and 
            not any(exc in t for exc in excluir))

def normalizar_cadena_alnum_mayus(s: str) -> str:
    """Convierte a mayúsculas y quita todo lo no alfanumérico."""
    return re.sub(r'[^A-Za-z0-9]', '', s or '').upper()

def buscar_imagen(carpeta_imagenes: str, codigo_canonico: str) -> str | None:
    """Busca la imagen ignorando mayúsculas/minúsculas."""
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

# ------------------------------------------------------------
# EXTRACCIÓN DE CÓDIGOS DESDE TABLAS
# ------------------------------------------------------------
def extraer_codigos(doc: Document) -> list[str]:
    """Extrae códigos de la columna CODIGO, tolerando letras al inicio y minúsculas."""
    codigos: list[str] = []
    patron_codigo = re.compile(r"[A-Za-z0-9][A-Za-z0-9.\-]{4,}", re.IGNORECASE)

    for tabla in doc.tables:
        if not tabla.rows:
            continue

        # Buscar índice de la columna CODIGO
        idx_codigo = None
        for r in range(min(3, len(tabla.rows))):
            for i, celda in enumerate(tabla.rows[r].cells):
                texto = celda.text.replace('\xa0', ' ').replace('\n', ' ').strip()
                if _es_encabezado_codigo(texto):
                    idx_codigo = i
                    break
            if idx_codigo is not None:
                break
        if idx_codigo is None:
            continue

        # Recorrer sólo esa columna
        for fila in tabla.rows[1:]:
            if idx_codigo >= len(fila.cells):
                continue
            texto = fila.cells[idx_codigo].text or ''
            texto = texto.replace('\xa0', ' ').replace('\n', ' ').strip()
            if not texto:
                continue

            for m in patron_codigo.findall(texto):
                canon = normalizar_cadena_alnum_mayus(m)
                if not any(c.isalpha() for c in canon):
                    continue
                if 5 <= len(canon) <= 20:
                    codigos.append(canon)

    return list(dict.fromkeys(codigos))

# ------------------------------------------------------------
# INSERCIÓN DE IMÁGENES
# ------------------------------------------------------------
def insertar_imagen_con_transparencia(run, img_path, ancho):
    """
    Inserta la imagen con espaciado para evitar que toque encabezados.
    No usa XML ni posicionamiento absoluto.
    """
    run.add_text("\n" * 2)  # margen superior controlado
    ancho_reducido = min(ancho * 0.9, 4.5)  # ancho máximo 4.5 pulgadas
    run.add_picture(img_path, width=Inches(ancho_reducido))
    run.add_text("\n" * 2)  # margen inferior


def _insert_paragraph_after_table(table) -> Paragraph:
    """Crea un párrafo justo debajo de una tabla."""
    new_p = OxmlElement("w:p")
    table._element.addnext(new_p)
    return Paragraph(new_p, table._element.getparent())

# ------------------------------------------------------------
# LÓGICA PRINCIPAL DE INSERCIÓN
# ------------------------------------------------------------
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

    ancho_pagina = 6.0
    espacio_entre = 0.15
    num_imgs = len(codigos)
    ancho_imagen = max(1.2, (ancho_pagina - espacio_entre * (num_imgs - 1)) / max(1, num_imgs))

    etiqueta_encontrada = False

    # --- Si existe ${etiqueta1}, usarla ---
    for p in doc.paragraphs:
        if "${etiqueta1}" in (p.text or ""):
            etiqueta_encontrada = True
            p.clear()
            run = p.add_run()
            print(f"Inserción múltiple de imágenes ({num_imgs}) en etiqueta1.")
            run.add_text("\n" * 4)

            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen(carpeta_imagenes, codigo)
                if img_path:
                    insertar_imagen_con_transparencia(run, img_path, ancho_imagen)
                    if i < num_imgs:
                        run.add_text("   ")
                    print(f"Imagen insertada: {os.path.basename(img_path)}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
            break

    # --- Si no hay etiqueta, insertar debajo de la tabla ---
    if not etiqueta_encontrada:
        print("No se encontró ${etiqueta1}. Insertando imágenes debajo de la tabla con los códigos.")
        cod_set = set(codigos)
        tabla_objetivo = None

        for tabla in doc.tables:
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
            p.alignment = 1
            p_format = p.paragraph_format
            p_format.space_before = Inches(1.0)  # baja 2.5 cm para no tapar encabezado
            run = p.add_run()

            for i, codigo in enumerate(codigos, start=1):
                img_path = buscar_imagen(carpeta_imagenes, codigo)
                if img_path:
                    insertar_imagen_con_transparencia(run, img_path, ancho_imagen)
                    if i < num_imgs:
                        run.add_text("   ")
                    print(f"Imagen insertada (debajo de tabla): {os.path.basename(img_path)}")
                else:
                    print(f"No se encontró la imagen para el código {codigo}")
        else:
            print("No se encontró ninguna tabla que contenga los códigos.")

    doc.save(ruta_doc)
    print(f"Documento actualizado: {ruta_doc}\n")

# ------------------------------------------------------------
# PROCESAMIENTO POR LOTES
# ------------------------------------------------------------
def procesar_lote(carpeta_docs="docs"):
    """Procesa todos los .docx dentro de la carpeta especificada."""
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

# ------------------------------------------------------------
# EJECUCIÓN
# ------------------------------------------------------------
if __name__ == "__main__":
    procesar_lote()
