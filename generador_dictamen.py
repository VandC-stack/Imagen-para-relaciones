import os
import json
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import unicodedata
import re
from docx import Document
from docx.oxml import parse_xml

def agregar_fondo_despues_de_renderizar(doc_path, image_path):
    """Agrega la imagen de fondo después de que el documento ha sido renderizado"""
    try:
        doc = Document(doc_path)
        section = doc.sections[0]
        header = section.header
        
        # Limpiar el header primero
        for element in header._element:
            header._element.remove(element)
        
        paragraph = header.add_paragraph()

        if not os.path.exists(image_path):
            print(f"⚠️ No se encontró la imagen: {image_path}")
            return False

        # Tamaño total de la página
        ancho_pagina = section.page_width
        alto_pagina = section.page_height

        # Crear relación para la imagen
        rel_id = header.part.relate_to(
            image_path,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            is_external=True
        )

        # Crear XML de imagen flotante
        background_xml = f'''
        <w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                 xmlns:v="urn:schemas-microsoft-com:vml"
                 xmlns:o="urn:schemas-microsoft-com:office:office"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <v:shape id="Fondo" 
                     style="position:absolute;margin-left:0;margin-top:0;width:{ancho_pagina.pt}pt;height:{alto_pagina.pt}pt;z-index:-251659264" 
                     o:allowincell="f" 
                     o:spid="_x0000_s1025" 
                     type="#_x0000_t75">
                <v:imagedata r:id="{rel_id}" o:title="Fondo"/>
            </v:shape>
        </w:pict>
        '''

        background_element = parse_xml(background_xml)
        paragraph._p.append(background_element)
        doc.save(doc_path)
        print(f"✅ Fondo agregado correctamente a: {doc_path}")
        return True
        
    except Exception as e:
        print(f"❌ Error al agregar fondo: {e}")
        return False

# --- CONFIGURACIÓN DE RUTAS ---
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "dictamenes_generados")
TEMPLATE_PATH = os.path.join(BASE_DIR, "Dictamen_plantilla.docx")
FONDO_PATH = os.path.join(BASE_DIR, "img", "Fondo.jpeg")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- FUNCIÓN PRINCIPAL ---
def generar_dictamen_desde_json(json_file: str):
    """ Genera uno o varios dictámenes Word (.docx) con el mismo diseño
    del machote, rellenando los valores de cada registro del JSON."""

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_PATH}")

    # Verificar que existe la imagen de fondo
    if not os.path.exists(FONDO_PATH):
        print(f"⚠️ Advertencia: No se encontró la imagen de fondo en: {FONDO_PATH}")

    # Leer archivo JSON
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        data = [data]

    generados = []

    for idx, registro in enumerate(data, 1):
        try:
            # 🔁 Cargar la plantilla para cada dictamen
            doc = DocxTemplate(TEMPLATE_PATH)

            # Preparar contexto (variables del machote)
            context = {
                "year": str(datetime.now().year),
                "lvxxxx": "hvxxxx",
                "normas": registro.get("Norma", "NOM-024"),
                "folio": registro.get("Folio", f"{idx:03d}"),
                "solicitud": registro.get("Solicitud", f"SOL{idx:03d}"),
                "lista": registro.get("Lista", ""),
                "verification": registro.get("FechaVerificacion", ""),
                "cliente": registro.get("Cliente", ""),
                "rfc": registro.get("RFC", ""),
                "producto": registro.get("Producto", ""),
                "pedimento": registro.get("Pedimento", ""),
                "fcostoadoraguargs": registro.get("FechaVerificacion", ""),
                "normades": registro.get("DescripcionNorma", ""),
                "TCantidad": registro.get("TotalCantidad", ""),
                "obs": registro.get("Observaciones", ""),
                "rowMarca": registro.get("Marca", ""),
                "rowCodigo": registro.get("Codigo", ""),
                "rowFactura": registro.get("Factura", ""),
                "rowCantidad": registro.get("Cantidad", ""),
                "firma1": registro.get("FirmaInspector", ""),
                "firma2": registro.get("FirmaSupervisor", ""),
                "nfirma1": registro.get("NombreInspector", ""),
                "nfirma2": registro.get("NombreSupervisor", ""),
                "fverificacion": registro.get("FechaVerificacion", ""),
                "femision": registro.get("FechaEmision", datetime.now().strftime("%d/%m/%Y")),
                "fverificacionlarga": registro.get("FechaVerificacionLarga", ""),
                "capitulo": registro.get("Capitulo", ""),
            }

            # Agregar imágenes si las hay
            for i in range(1, 11):
                img_key = f"img{i}"
                img_path = registro.get(img_key)
                if img_path and os.path.exists(img_path):
                    try:
                        context[img_key] = InlineImage(doc, img_path, width=Mm(50))
                    except Exception as img_err:
                        print(f"⚠ No se pudo cargar imagen {i}: {img_err}")
                        context[img_key] = ""
                else:
                    context[img_key] = ""

            # Rellenar el documento
            doc.render(context)

            # Guardar nuevo dictamen
            cliente = registro.get("Cliente", "Cliente").replace(" ", "_")
            file_name = f"Dictamen_{cliente}_{idx:03d}.docx"
            output_path = os.path.join(OUTPUT_DIR, file_name)

            doc.save(output_path)
            
            # 🔁 AGREGAR FONDO DESPUÉS DE RENDERIZAR
            if os.path.exists(FONDO_PATH):
                agregar_fondo_despues_de_renderizar(output_path, FONDO_PATH)
            else:
                print(f"⚠️ No se pudo agregar fondo, archivo no encontrado: {FONDO_PATH}")

            generados.append(output_path)
            print(f"✅ Generado: {output_path}")

        except Exception as e:
            print(f"⚠ Error en registro {idx}: {e}")
            continue

    print(f"\n=== {len(generados)} dictámenes generados correctamente ===")
    return generados

# --- EJECUCIÓN DIRECTA ---
if __name__ == "__main__":
    print("=== GENERADOR DE DICTÁMENES BOSCH ===\n")

    json_files = [f for f in os.listdir(DATA_DIR) if f.endswith(".json")]
    if not json_files:
        print("No se encontraron archivos JSON en /data.")
        exit()

    print("Archivos disponibles:")
    for i, f in enumerate(json_files, 1):
        print(f"  {i}. {f}")

    seleccion = input("\nSeleccione el número del archivo JSON a usar: ")
    try:
        idx = int(seleccion) - 1
        json_path = os.path.join(DATA_DIR, json_files[idx])
        generar_dictamen_desde_json(json_path)
    except Exception as e:
        print(f"Error: {e}")