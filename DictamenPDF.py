"""
Generador de Plantilla PDF - Versión Base
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib import colors
import os

# Tamaño carta en puntos
LETTER_WIDTH = 8.5 * inch
LETTER_HEIGHT = 11 * inch

class PDFGenerator:
    def __init__(self):
        self.doc = None
        self.elements = []
        self.styles = getSampleStyleSheet()
        self.total_pages = None
        
    def crear_estilos(self):
        """Crea los estilos personalizados para el documento"""
        
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=colors.black,
            alignment=1,
            spaceAfter=6,
            fontName='Helvetica-Bold'
        )
        
        self.code_style = ParagraphStyle(
            'CustomCode',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.black,
            alignment=1,
            spaceAfter=20,
            fontName='Helvetica'
        )
        
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=4,
            spaceAfter=12,
            fontName='Helvetica'
        )
        
        self.bold_style = ParagraphStyle(
            'CustomBold',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=4,
            spaceAfter=12,
            fontName='Helvetica-Bold'
        )
        
        self.label_style = ParagraphStyle(
            'CustomLabel',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.black,
            alignment=1,
            spaceAfter=8,
            fontName='Helvetica-Bold'
        )
        
        self.image_style = ParagraphStyle(
            'CustomImage',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=1,
            spaceAfter=15,
            fontName='Helvetica'
        )

    #Cuenta las hojas de forma automatica. 
    def _count_pages(self, canvas, doc):
        self.total_pages = canvas.getPageNumber()

    def agregar_primera_pagina(self):
        """Agrega el contenido de la primera página"""
        
        print("📄 Generando primera página...")
        
        # FECHAS
        cliente_text = '<b>Fecha de Inspección:</b> ${fverificacion}'
        self.elements.append(Paragraph(cliente_text, self.normal_style))
        
        rfc_text = '<b>Fecha de Emisión:</b> ${femision}'
        self.elements.append(Paragraph(rfc_text, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # CLIENTE Y RFC
        cliente_text = '<b>Cliente:</b> ${cliente}'
        self.elements.append(Paragraph(cliente_text, self.normal_style))
        
        rfc_text = '<b>RFC:</b> ${rfc}'
        self.elements.append(Paragraph(rfc_text, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # TEXTO PRINCIPAL
        texto_principal = (
            "De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, 68 y 140 de la Ley de Infraestructura de la "
            "Calidad; 50 del Reglamento de la Ley Federal de Metrología y Normalización; Punto 2.4.8 Fracción III ACUERDO por el que la "
            "Secretaría de Economía emite Reglas y criterios de carácter general en materia de comercio exterior; publicado en el Diario Oficial de la "
            "Federación el 09 de mayo de 2022 y posteriores modificaciones; esta Unidad de Inspección a solicitud de la persona moral denominada "
            "${cliente} dictamina el Producto: ${producto}; que la mercancía importada bajo el pedimento aduanal No. ${pedimento} de fecha "
            "${fverificacionlarga}, fue etiquetada conforme a los requisitos de Información Comercial en el capítulo ${capitulo} de la Norma Oficial Mexicana "
            "${norma} ${normades} Cualquier otro requisito establecido en la norma referida, es responsabilidad del titular de este Dictamen."
        )
        
        self.elements.append(Paragraph(texto_principal, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # TABLA DE PRODUCTOS
        productos_data = [
            ['MARCA', 'CÓDIGO', 'FACTURA', 'CANTIDAD'],
            ['${rowMarca}', '${rowCodigo}', '${rowFactura}', '${rowCantidad}']
        ]
        
        productos_table = Table(productos_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
        productos_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('BOLD', (0,0), (-1,0), True),
        ]))
        
        self.elements.append(productos_table)
        self.elements.append(Spacer(1, 0.2*inch))
        
        # TAMAÑO DEL LOTE
        lote_data = [
            ['TAMAÑO DEL LOTE', '${TCantidad}']
        ]
        
        lote_table = Table(lote_data, colWidths=[4.5*inch, 1.5*inch])
        lote_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (0,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('BOLD', (0,0), (0,0), True),
        ]))
        
        self.elements.append(lote_table)
        self.elements.append(Spacer(1, 0.2*inch))
        
        # OBSERVACIONES
        obs1_text = '<b>OBSERVACIONES:</b> La imagen amparada en el dictamen es una muestra de etiqueta que aplica para todos los modelos declarados en el presente dictamen lo anterior fue constatado durante la inspección.'
        self.elements.append(Paragraph(obs1_text, self.normal_style))
        
        obs2_text = '<b>OBSERVACIONES:</b> ${obs}'
        self.elements.append(Paragraph(obs2_text, self.normal_style))
        self.elements.append(Spacer(1, 0.3*inch))
        
    def agregar_segunda_pagina(self):
        print("📄 Generando segunda página...")

        # NO usar PageBreak, Platypus decide cuándo cortar
        # self.elements.append(PageBreak())

        self.elements.append, self.label_style

        etiquetas_linea1 = "${etiqueta1}   ${etiqueta2}   ${etiqueta3}   ${etiqueta4}   ${etiqueta5}"
        self.elements.append(Paragraph(etiquetas_linea1, self.image_style))

        etiquetas_linea2 = "${etiqueta6}   ${etiqueta7}   ${etiqueta8}   ${etiqueta9}   ${etiqueta10}"
        self.elements.append(Paragraph(etiquetas_linea2, self.image_style))

        self.elements.append(Spacer(1, 0.4 * inch))

        self.elements.append, self.label_style

        for i in range(1, 11):
            self.elements.append(Paragraph(f"${{img{i}}}", self.image_style))

        firmas_data = [
            ['${firma1}', '', '${firma2}'],
            ['${nfirma1}', '', '${nfirma2}'],
            ['Nombre del Inspector', '', 'Nombre del responsable de\nsupervisión UI']
        ]

        firmas_table = Table(firmas_data, colWidths=[2.8*inch, 0.4*inch, 2.8*inch])

        firmas_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('BOLD', (0,2), (-1,2), True),
            ('LINEBELOW', (0,0), (-1,-1), 0, colors.white),
            ('BOX', (0,0), (-1,-1), 0, colors.white),
            ('INNERGRID', (0,0), (-1,-1), 0, colors.white),
            ('LINEBELOW', (0,0), (0,0), 1, colors.black),
            ('LINEBELOW', (2,0), (2,0), 1, colors.black),
        ]))

        self.elements.append(firmas_table)

    def agregar_encabezado_pie_pagina(self, canvas, doc):
        """Agrega encabezado, pie de página y numeración"""
        
        canvas.saveState()
        
        # Fondo
        image_path = "img/Fondo.jpeg"
        if os.path.exists(image_path):
            try:
                canvas.drawImage(image_path, 0, 0, width=LETTER_WIDTH, height=LETTER_HEIGHT)
            except:
                pass
        
        # Encabezado
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(LETTER_WIDTH/2, LETTER_HEIGHT-60, "DICTAMEN DE CUMPLIMIENTO")
        
        canvas.setFont("Helvetica", 10)
        codigo_text = "${cadena_identificacion}"
        canvas.drawCentredString(LETTER_WIDTH/2, LETTER_HEIGHT-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion_total = self.doc.page  # total real generado
        numeracion = f"Página {pagina_actual} de {numeracion_total}"

        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(LETTER_WIDTH-72, LETTER_HEIGHT-50, numeracion)
        
        # Pie de página
        footer_text = "Este Dictamen de Cumplimiento se emitió por medios electrónicos, conforme al oficio de autorización DGN.312.05.2012.106 de fecha 10 de enero de 2012 expedido por la DGN a esta Unidad de Inspección."
        formato_text = "Formato: PT-F-208B-00-3"

        canvas.setFont("Helvetica", 7)

        lines = []
        words = footer_text.split()
        current_line = ""
        for word in words:
            test_line = current_line + " " + word if current_line else word
            if len(test_line) <= 150:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)

        line_height = 8
        start_y = 60

        for i, line in enumerate(lines):
            text_width = canvas.stringWidth(line, "Helvetica", 7)
            available_width = LETTER_WIDTH - 144
            if text_width < available_width * 0.8:
                x_position = (LETTER_WIDTH - text_width) / 2
            else:
                x_position = 72
            canvas.drawString(x_position, start_y - (i * line_height), line)

        canvas.drawRightString(LETTER_WIDTH - 72, start_y - (len(lines) * line_height) - 4, formato_text)

        canvas.restoreState()

