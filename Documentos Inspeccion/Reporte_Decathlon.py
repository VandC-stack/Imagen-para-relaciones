"""
Generador de plantilla PDF para Reporte Decathlon.

Esta implementación usa reportlab (está en requirements.txt) para generar
una página con encabezado, datos del cliente, tabla de equipos, observaciones
y espacio para firmado.

Funciona con `lista_equipos` como lista de strings o lista de diccionarios
con campos como `equipo`, `estado`, `marca`, `modelo`.
"""

from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether, Frame, PageTemplate
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from datetime import datetime
import os


def generar_reporte_decathlon(nombre_cliente, folio_inspeccion, fecha_inspeccion, tabla_relacion=None, lista_equipos=None, observaciones_generales='', ruta_guardado='reporte.pdf', folio_visita=None, destinatario='Diana Cumpean', direccion=None):
    if direccion is None:
        direccion = {
            "CLIENTE": "ARTÍCULOS DEPORTIVOS DECATHLON SA DE CV.",
            "CALLE Y NO": "Avenida Ejército Nacional, No. Ext. 826",
            "MUNICIPIO O ALCALDÍA": "Polanco III Sección, C.P. 11540",
            "CIUDAD O ESTADO": "Miguel Hidalgo, Ciudad de México, México"
        }

    # Usar hoja "Carta" (letter) por defecto para salida en formato carta
    page_w, page_h = letter
    left_margin = 18 * mm
    right_margin = 18 * mm
    top_margin = 15 * mm
    bottom_margin = 15 * mm
    # espacio extra superior para páginas posteriores (desplaza la tabla continuada)
    extra_top_later = 12 * mm

    doc = BaseDocTemplate(ruta_guardado, pagesize=letter)

    # Definir frames para primera pagina y páginas posteriores
    frame_first_height = page_h - top_margin - bottom_margin
    frame_later_height = page_h - (top_margin + extra_top_later) - bottom_margin
    frame_width = page_w - left_margin - right_margin

    frame_first = Frame(left_margin, bottom_margin, frame_width, frame_first_height, id='frame_first')
    frame_later = Frame(left_margin, bottom_margin, frame_width, frame_later_height, id='frame_later')
    # Preparar membrete y handler de página antes de crear PageTemplates
    membrete_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'img', 'Membrete.jpg'))
    meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']

    def on_page(canvas, doc_):
        canvas.saveState()
        # Dibujar fondo (membrete) si existe: escalar y centrar para cubrir la hoja (Carta / Letter)
        if os.path.exists(membrete_path):
            try:
                page_w, page_h = doc_.pagesize
                # Usar ImageReader para obtener proporciones de la imagen
                img_reader = ImageReader(membrete_path)
                img_w, img_h = img_reader.getSize()
                if img_w > 0 and img_h > 0:
                    # Escalar para cubrir la página preservando la relación de aspecto (cover)
                    scale = max(page_w / img_w, page_h / img_h)
                    draw_w = img_w * scale
                    draw_h = img_h * scale
                    # Centrar la imagen en la página
                    x = (page_w - draw_w) / 2.0
                    y = (page_h - draw_h) / 2.0
                    canvas.drawImage(img_reader, x, y, width=draw_w, height=draw_h, preserveAspectRatio=False, mask='auto')
                else:
                    # Fallback: estirar a la página
                    canvas.drawImage(membrete_path, 0, 0, width=page_w, height=page_h, preserveAspectRatio=False, mask='auto')
            except Exception:
                pass
        canvas.restoreState()

    # PageTemplates: primero usa frame_first, luego cambiar a frame_later
    # Asignamos la función on_page como handler para dibujar el membrete en cada página
    first_pt = PageTemplate(id='First', frames=[frame_first], onPage=on_page)
    later_pt = PageTemplate(id='Later', frames=[frame_later], onPage=on_page)
    # Hacer que después de la primera página se use la plantilla 'Later'
    first_pt.nextTemplate = 'Later'
    doc.addPageTemplates([first_pt, later_pt])
    styles = getSampleStyleSheet()
    normal = styles['Normal']
    heading = styles['Heading1']

    elements = []

    # Header
    header_style = ParagraphStyle('header', parent=styles['Heading2'], alignment=1)
    elements.append(Paragraph('REPORTE DE INSPECCIÓN', header_style))
    elements.append(Spacer(1, 2*mm))

    # Preparar No ACUSE formateado: extraer número del folio de visita y el sufijo año desde la tabla_relacion
    right_style = ParagraphStyle('right', parent=normal, alignment=2, fontSize=10)
    acuse_display = None
    try:
        if folio_visita:
            import re
            digits = re.sub(r"\D", "", str(folio_visita))
            num = int(digits) if digits else 0
            num_formatted = f"{num:06d}"
            # extraer año de la primera SOLICITUD si está disponible
            year_suffix = None
            if tabla_relacion and isinstance(tabla_relacion, list):
                for r in tabla_relacion:
                    sol = r.get('SOLICITUD') or r.get('Solicitud') or r.get('solicitud')
                    if sol and '/' in str(sol):
                        parts = str(sol).split('/')
                        if len(parts) >= 2:
                            year_suffix = parts[-1]
                            break
            if year_suffix:
                acuse_display = f"No ACUSE {num_formatted}/{year_suffix}"
            else:
                acuse_display = f"No ACUSE {num_formatted}"
    except Exception:
        acuse_display = None
    if acuse_display:
        elements.append(Paragraph(f"<b>{acuse_display}</b>", right_style))
    # Fecha de impresión (fecha actual) en español
    meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
    now = datetime.now()
    fecha_imp = f"{now.day:02d} de {meses[now.month-1]} de {now.year}"
    elements.append(Paragraph(fecha_imp, right_style))
    elements.append(Spacer(1, 4*mm))

    # Cliente y dirección (lado izquierdo)
    cliente_para = f"<b>{nombre_cliente}</b><br/>{direccion.get('CALLE Y NO','')}<br/>{direccion.get('MUNICIPIO O ALCALDÍA','')}<br/>{direccion.get('CIUDAD O ESTADO','')}"
    elements.append(Paragraph(cliente_para, ParagraphStyle('left_cliente', parent=normal, fontSize=10)))
    elements.append(Spacer(1, 4*mm))

    # Si llega una tabla de relación (dictámenes), imprimirla con columnas específicas
    if tabla_relacion and isinstance(tabla_relacion, list) and len(tabla_relacion) > 0 and isinstance(tabla_relacion[0], dict):
        # Obtener fecha de verificación desde la tabla_relacion (primera encontrada)
        fecha_verificacion = None
        for rec in tabla_relacion:
            fv = rec.get('FECHA DE VERIFICACION') or rec.get('Fecha de Verificacion') or rec.get('FECHA_DE_VERIFICACION')
            if fv:
                fecha_verificacion = fv
                break

        # Texto introductorio con fecha de verificación (formato original ISO o dd/mm/yyyy)
        def formato_fecha_origen(f):
            if not f:
                return ''
            try:
                if '-' in f:
                    d = datetime.strptime(f[:10], '%Y-%m-%d')
                else:
                    d = datetime.strptime(f[:10], '%d/%m/%Y')
                return d.strftime('%d/%m/%Y')
            except Exception:
                return str(f)

        fecha_verif_print = formato_fecha_origen(fecha_verificacion)

        table_data = [["No. DE SOLICITUD", "No. DE DICTAMEN", "REFERENCIA / PEDIMENTO"]]

        # Deduplicar por folio de dictamen: mantener primer SOLICITUD y PEDIMENTO
        seen = {}
        for item in tabla_relacion:
            solicitud = item.get('SOLICITUD') or item.get('Solicitud') or item.get('solicitud') or ''
            folio = item.get('FOLIO') or item.get('folio') or item.get('Folio') or item.get('folio_dictamen') or ''
            pedimento = item.get('PEDIMENTO') or item.get('pedimento') or item.get('Referencia') or item.get('referencia') or ''
            key = str(folio)
            if key in seen:
                # If same folio, ensure we still have solicitud/pedimento populated
                if not seen[key][0] and solicitud:
                    seen[key] = (solicitud, pedimento)
                continue
            seen[key] = (solicitud, pedimento)

        # Construir filas a partir de seen, preservando orden de aparición in tabla_relacion
        folio_order = []
        for item in tabla_relacion:
            fol = str(item.get('FOLIO') or item.get('folio') or item.get('Folio') or '')
            if fol and fol not in folio_order:
                folio_order.append(fol)

        for fol in folio_order:
            sol, ped = seen.get(fol, ('', ''))
            table_data.append([str(sol), str(fol), str(ped)])
        tbl_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),
            ('TOPPADDING', (0,0), (-1,0), 4),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ])
        # Formatear fecha_inspeccion en formato español para el cuerpo
        def fecha_es(f):
            meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
            if isinstance(f, str):
                try:
                    d = datetime.fromisoformat(f)
                except Exception:
                    try:
                        d = datetime.strptime(f, '%Y-%m-%d')
                    except Exception:
                        d = None
            elif isinstance(f, datetime):
                d = f
            else:
                d = None
            if d:
                return d.strftime(f"%d de {meses[d.month-1]} de %Y")
            return str(f)

        fecha_inspeccion_es = fecha_es(fecha_inspeccion)
        # Usar fecha_verif_print en el cuerpo (fecha de verificación)
        fecha_cuerpo = fecha_verif_print or fecha_inspeccion_es
        elements.append(Spacer(1, 2*mm))
        elements.append(Paragraph(f"Atn: {destinatario}", ParagraphStyle('atn', parent=normal, fontSize=10)))
        elements.append(Spacer(1, 1*mm))
        elements.append(Paragraph(f"Estimada {destinatario.split()[0]}:", normal))
        elements.append(Spacer(1, 1*mm))
        elements.append(Paragraph(f"Te confirmo que derivado de la visita del día {fecha_cuerpo} se emitieron los siguientes dictámenes de cumplimiento:", normal))
        elements.append(Spacer(1, 2*mm))

        # Ajustar anchos de columnas según contenido (estimación simple)
        max_solicitud_len = max((len(r[0]) for r in table_data[1:]), default=10)
        max_folio_len = max((len(r[1]) for r in table_data[1:]), default=6)
        col1 = max(40*mm, min(80*mm, max_solicitud_len * 2.2 * mm))
        col2 = max(30*mm, min(60*mm, max_folio_len * 2.2 * mm))
        table = Table(table_data, colWidths=[col1, col2, None], repeatRows=1)
        table.setStyle(tbl_style)
        # Insertar la tabla inmediatamente después del texto introductorio
        elements.append(table)
    else:
        # Tabla de equipos (antigua funcionalidad)
        table_data = []
        # Determine headers based on contents
        if lista_equipos and isinstance(lista_equipos[0], dict):
            headers = [k.capitalize() for k in lista_equipos[0].keys()]
            table_data.append(headers)
            for item in lista_equipos:
                row = [str(item.get(k, '')) for k in lista_equipos[0].keys()]
                table_data.append(row)
        else:
            table_data.append(['#', 'Equipo'])
            for i, equipo in enumerate(lista_equipos or [], start=1):
                table_data.append([str(i), str(equipo)])

        tbl_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ])

        col_widths = None
        if table_data and len(table_data[0]) == 2:
            col_widths = [20*mm, None]

        table = Table(table_data, colWidths=col_widths)
        table.setStyle(tbl_style)
        elements.append(table)
    elements.append(Spacer(1, 6*mm))

    # Observaciones
    elements.append(Paragraph('<b>Observaciones Generales:</b>', styles['Heading3']))
    obs_style = ParagraphStyle('obs', parent=normal, fontSize=10, leading=12)
    for line in str(observaciones_generales or '').split('\n'):
        elements.append(Paragraph(line, obs_style))
    elements.append(Spacer(1, 4*mm))

    # Enlace y cierre solicitado
    elements.append(Paragraph('En el siguiente link podrás descargar dichos dictámenes:', obs_style))
    elements.append(Spacer(1, 2*mm))
    # Nota de cierre
    cierre_text = ('Sin más por el momento me reitero a tus apreciables órdenes para cualquier aclaración o comentario al respecto.')
    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph(cierre_text, obs_style))
    elements.append(Spacer(1, 12*mm))

    # Firma personalizada: alinear a la derecha
    elements.append(Spacer(1, 8*mm))
    atentamente = ' '.join(list('ATENTAMENTE.'))
    elements.append(Paragraph(atentamente, ParagraphStyle('atn_center', parent=styles['Heading3'], alignment=1)))
    elements.append(Spacer(1, 4*mm))

    # Construir bloque de firma alineado a la derecha usando tabla
    firma_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Firmas', 'AFLORES.png'))
    usable_width = page_w - left_margin - right_margin
    right_col_width = usable_width * 0.4
    left_col_width = usable_width - right_col_width

    # Preparar contenido derecho (imagen + texto) en una tabla interna
    right_rows = []
    if os.path.exists(firma_path):
        try:
            sig = Image(firma_path, width=50*mm, height=20*mm)
            right_rows.append([sig])
        except Exception:
            pass
    # Agregar textos de firma
    right_rows.append([Paragraph('<b>Arturo Flores Gómez</b>', ParagraphStyle('name_r', parent=normal, alignment=2, fontSize=11))])
    right_rows.append([Paragraph('Verificación y Control', ParagraphStyle('role_r', parent=normal, alignment=2))])
    right_rows.append([Paragraph('(0155) 55727984', ParagraphStyle('phone_r', parent=normal, alignment=2))])
    right_rows.append([Paragraph('www.vyc.com.mx', ParagraphStyle('web_r', parent=normal, alignment=2))])

    right_table = Table(right_rows, colWidths=[right_col_width])
    right_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))

    sig_block = Table([[ '', right_table ]], colWidths=[left_col_width, right_col_width])
    sig_block.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))

    # Mantener el bloque de firma junto y lo más pegado posible
    elements.append(KeepTogether([sig_block]))

    # Footer/background handled by PageTemplate.onPage (defined earlier)

    # BaseDocTemplate.build no acepta onFirstPage/onLaterPages kwargs; PageTemplate.onPage se usa en su lugar
    doc.build(elements)


if __name__ == '__main__':
    # Ejemplo de uso rápido
    ejemplo_equipos = [
        {'equipo': 'Bicicleta Estática', 'marca': 'FitCo', 'modelo': 'X200', 'estado': 'Operativa'},
        {'equipo': 'Cinta de Correr', 'marca': 'RunFast', 'modelo': 'R1000', 'estado': 'Mantenimiento'},
    ]
    # Ejemplo de tabla de relación (dictámenes)
    ejemplo_tabla_relacion = [
        {'solicitud': 'SOL-001', 'folio': 'D-000123', 'pedimento': 'PED-98765'},
        {'solicitud': 'SOL-002', 'folio': 'D-000124', 'pedimento': 'PED-98766'},
    ]
    generar_reporte_decathlon(
        'ARTÍCULOS DEPORTIVOS DECATHLON SA DE CV.',
        'FOLIO-000123',
        fecha_inspeccion='2025-12-12',
        tabla_relacion=ejemplo_tabla_relacion,
        lista_equipos=None,
        observaciones_generales='Se observó desgaste en la banda de la cinta de correr. Requiere mantenimiento preventivo.',
        ruta_guardado='reporte_decathlon_ejemplo.pdf',
        folio_visita='CP0000118',
        destinatario='Diana Cumpean'
    )
