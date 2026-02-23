from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, PageBreak, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas
import os
import json

# =========================================================
#   ESTILOS
# =========================================================
styles = getSampleStyleSheet()

style_normal = ParagraphStyle(
    "normal9",
    parent=styles["Normal"],
    fontSize=8,
    leading=11,
)

style_title = ParagraphStyle(
    "title9bold",
    parent=styles["Heading2"],
    fontSize=8,
    leading=11,
    spaceAfter=4,
    alignment=1,
    bold=True
)

style_subtitle = ParagraphStyle(
    "subtitle9bold",
    parent=styles["Heading3"],
    fontSize=8,
    leading=11,
    spaceAfter=3,
    bold=True
)

style_header_left = ParagraphStyle(
    "header_left",
    parent=styles["Normal"],
    fontSize=8,
    alignment=0,
    leading=11,
    bold=True
)

style_header_right = ParagraphStyle(
    "header_right",
    parent=styles["Normal"],
    fontSize=8,
    alignment=2,
    leading=11,
    bold=True
)

# =========================================================
#   FUNCIONES PARA TABLAS
# =========================================================
def build_table(title, subtitle, incisos, realiza="GT/SUB", especiales=None):
    """
    Genera UNA TABLA completa con:
    - Título combinado
    - Subtítulo (solo si existe)
    - Encabezados ajustados
    - Filas de incisos
    """

    if especiales is None:
        especiales = {}

    colWidths = [60*mm, 20*mm, 30*mm, 20*mm, 60*mm]

    data = []
    spans = []

    # --------------------------
    #  TÍTULO PRINCIPAL
    # --------------------------
    if title and title.strip() != "":
        data.append([Paragraph(f"<b>{title}</b>", style_title)])
        data[-1] += ["", "", "", ""]
        spans.append(('SPAN', (0, len(data)-1), (4, len(data)-1)))

    # --------------------------
    #  SUBTÍTULO (si aplica)
    # --------------------------
    if subtitle and subtitle.strip() != "":
        data.append([Paragraph(f"<b>{subtitle}</b>", style_subtitle)])
        data[-1] += ["", "", "", ""]
        spans.append(('SPAN', (0, len(data)-1), (4, len(data)-1)))


    # --------------------------
    #   ENCABEZADOS
    # --------------------------
    headers = [
        Paragraph("Actividad específica", style_normal),
        Paragraph("Realiza", style_normal),
        Paragraph("Conforme/<br/>No Conforme", style_normal),
        Paragraph("No Aplica", style_normal),
        Paragraph("Observaciones", style_normal),
    ]

    data.append(headers)

    # --------------------------
    # FILAS DE INCISOS
    # --------------------------
    for inciso in incisos:
        conf = especiales.get(inciso, {}).get("conf", "")
        na = especiales.get(inciso, {}).get("na", "")

        row = [
            Paragraph(inciso, style_normal),
            Paragraph(realiza, style_normal),
            Paragraph(conf, style_normal),
            Paragraph(na, style_normal),
            Paragraph("", style_normal)
        ]
        data.append(row)

    # --------------------------
    #  ESTILO DE TABLA
    # --------------------------
    table = Table(data, colWidths=colWidths)

    table_style = [
        ('BOX', (0,0), (-1,-1), 0.8, colors.black),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ] + spans

    table.setStyle(TableStyle(table_style))

    return table

# =========================================================
#   TABLA INICIAL
# =========================================================
def header_table(datos):

    colWidths = [50*mm, 80*mm, 60*mm]

    # ====================================================
    #   FIRMA Y NOMBRE EN TABLA INTERNA (SIN KeepTogether)
    # ====================================================
    firma_path = "Firmas/MTERREZ.png"

    if os.path.exists(firma_path):
        firma_img = Image(firma_path, width=40*mm, height=15*mm)
    else:
        firma_img = Paragraph("(Firma no encontrada)", style_normal)

    firma_table = Table(
        [
            [firma_img],
        ],
        colWidths=[55*mm]
    )

    firma_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))

    # ====================================================
    #   CONTENIDO GENERAL DEL ENCABEZADO
    # ====================================================
    info = [
        ["Número de solicitud:", datos.get("solicitud")],
        ["Servicio:", datos.get("servicio")],
        ["Fecha:", datos.get("fecha")],
        ["Cliente:", datos.get("cliente")],
        ["Supervisór:", "Mario Terrez González"],
    ]

    data = []

    # Primera fila con firma
    data.append([
        Paragraph("<b>Número de solicitud:</b>", style_normal),
        Paragraph(str(datos.get("solicitud") or ""), style_normal),
        firma_table   # firma únicamente aquí
    ])

    # Resto de filas
    for label, value in info[1:]:
        data.append([
            Paragraph(f"<b>{label}</b>", style_normal),
            Paragraph(str(value or ""), style_normal),
            ""  # vacío para SPAN
        ])

    # ====================================================
    #   TABLA SIN BORDES + SPAN
    # ====================================================
    table = Table(data, colWidths=colWidths)

    table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0, colors.white),
        ('GRID', (0,0), (-1,-1), 0, colors.white),

        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),

        ('VALIGN', (0,0), (-1,-1), 'TOP'),

        # SPAN de toda la tercera columna
        ('SPAN', (2,0), (2,4)),
    ]))

    return table

# =========================================================
#   ENCABEZADO Y PIE POR PÁGINA
# =========================================================
def add_header_footer(canvas: Canvas, doc):
    canvas.saveState()

    # Fondo
    bg_path = "img/Oficios.png"
    if os.path.exists(bg_path):
        bg = ImageReader(bg_path)
        canvas.drawImage(bg, 0, 0, width=letter[0], height=letter[1])

    # Encabezado izquierdo
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(40*mm, 260*mm, "FORMATO DE SUPERVISIÓN DE LAS ACTIVIDADES DE INSPECCIÓN DE LA UI.")

    # Encabezado derecho
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(200*mm, 268*mm, "PA-F-13A-00-3")

    # Contador de páginas
    page = canvas.getPageNumber()
    canvas.drawRightString(200*mm, 263*mm, f"Página {page}")

    canvas.restoreState()

# =========================================================
#   GENERAR DOCUMENTO
# =========================================================
def generar_supervision(datos, archivo="Supervision_Final.pdf"):

    story = []

    # ENCABEZADO
    story.append(header_table(datos))
    story.append(Spacer(1, 5))

    # =====================================================================
    #  TABLAS DEFINIDAS POR EL USUARIO (SIN CAMBIOS EN SU CONTENIDO)
    # =====================================================================

    incisos_1 = [
        "A) Menciona las normas para las que se solicita el servicio?",
        "B) Los datos referentes a la constitución de la empresa son correctos?",
        "C) Los datos del representante legal son correctos?",
        "D) Las tarifas son las acordadas?",
        "E) El cliente entregó el original firmado?"
    ]

    story.append(build_table(
        "1. ETAPA DEL SERVICIO A EVALUAR: EMISIÓN DE CONTRATO PARA INFORME TÉCNICO, DISEÑO, CONSTANCIA Y DICTAMEN",
        None,
        incisos_1
    ))


    incisos_21 = [
        "A) La norma en la cual se inspecciona el producto es correcta?",
        "B) La lista de inspección está correctamente llenada?",
        "C) Los resultados de inspección son correctos?",
        "D) Las mediciones y/o medidas de etiquetas son precisas?",
        "E) Se realizó con las muestras necesarias?",
        "F) La inspección la realizó personal distinto al del diseño?",
        "G) Aplica correctamente el procedimiento de muestreo?"
    ]

    story.append(build_table(
        "2. ETAPA DEL DESARROLLO DEL SERVICIO",
        "2.1 Inspecciones de resultados para constancia, informe técnico y diseño",
        incisos_21
    ))


    incisos_22 = [
        "A) Recibió la documentación necesaria para realizar el servicio?",
        "B) Tiene los elementos necesarios (normas, criterios, instrumentos)?",
        "C) La actividad se realizó conforme a prácticas comunes?",
        "D) Existe situación de riesgo que afecte la dictaminación?",
        "E) Existe impedimento logístico para realizar la inspección?",
        "F) La inspección la realizó personal distinto al del diseño?"
    ]

    especiales_22 = {
        incisos_22[5]: {"na": "N/A"},
        incisos_22[0]: {"conf": "C"},
        incisos_22[1]: {"conf": "C"},
        incisos_22[2]: {"conf": "C"},
        incisos_22[3]: {"conf": "C"},
        incisos_22[4]: {"conf": "C"},
        incisos_22[5]: {"na": "N/A"},
    }

    story.append(build_table(
        "2.2 Inspecciones para Dictamen",
        None,
        incisos_22,
        realiza="INSP",
        especiales=especiales_22
    ))


    # ETAPA 3
    incisos_31 = [
        "A) Producto, marca y contenido",
        "B) Descripción SKU / archivo",
        "C) Acompañado de diseño de etiqueta",
        "D) Observaciones (tamaño de fuente)",
        "E) Indicaciones de colocación",
        "F) Inspector responsable",
        "G) Norma ocupada"
    ]

    story.append(build_table(
        "3. Etapa Final del Final del Servicio: Evaluación de la Información contenida en los documentos finales emitidos",
        "3.1 Constancia",
        incisos_31
    ))

    story.append(build_table(
        None,
        "3.2 Dictamen",
        incisos_31
    ))

    story.append(build_table(
        None,
        "3.3 Informe Técnico",
        incisos_31
    ))

    # NOTAS
    notas = (
        "Notas:<br/>"
        "1.- Se deben llenar solamente los recuadros correspondientes según la etapa evaluada.<br/>"
        "2.- En “Realiza” se indica quién ejecuta la supervisión.<br/>"
        "3.- “Conforme/No conforme”: colocar C o NC según corresponda.<br/>"
        "4.- “N/A” indica que el punto no es susceptible de evaluación."
    )

    story.append(Spacer(1, 8))
    story.append(Paragraph(notas, style_normal))

    # CREAR DOCUMENTO
    doc = SimpleDocTemplate(
        archivo,
        pagesize=letter,
        topMargin=25*mm,
        bottomMargin=15*mm,
        leftMargin=15*mm,
        rightMargin=15*mm,
    )

    doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)

    print(f"✅ Archivo generado correctamente → {archivo}")

# =========================================================
#   EJEMPLO DE USO
# =========================================================
if __name__ == "__main__":
    datos_demo = {
        "solicitud": "006916/25, 006917/25, 006918/25",
        "servicio": "Dictamen",
        "fecha": "02/12/2025",
        "cliente": "ARTÍCULOS DEPORTIVOS DECATHLON S.A. DE C.V.",
        "supervisor": " "
    }

    # Genera un ejemplo localmente sólo cuando se ejecuta este script directamente.
    generar_supervision(datos_demo, os.path.join("Plantillas PDF", "Formato_Supervision.pdf"))

