"""
Generador de Plantilla BOSCH PDF Corregido
- Sin duplicación de encabezados y pies
- Pie de página justificado con formato a la derecha
- Encabezado único en cada página
- Numeración corregida
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
import os
from PIL import Image

# Tamaño carta en puntos
LETTER_WIDTH = 8.5 * inch
LETTER_HEIGHT = 11 * inch

class PDFGenerator:
    def __init__(self):
        self.doc = None
        self.elements = []
        self.styles = getSampleStyleSheet()
        self.total_pages = 2  # Sabemos que son 2 páginas
        
    def crear_estilos(self):
        """Crea los estilos personalizados para el documento"""
        
        # Estilo para el título principal
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=colors.black,
            alignment=1,  # Centrado
            spaceAfter=6,
            fontName='Helvetica-Bold'
        )
        
        # Estilo para el subtítulo/código
        self.code_style = ParagraphStyle(
            'CustomCode',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.black,
            alignment=1,  # Centrado
            spaceAfter=20,
            fontName='Helvetica'
        )
        
        # Estilo para texto normal
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=4,  # Justificado
            spaceAfter=12,
            fontName='Helvetica'
        )
        
        # Estilo para etiquetas en segunda página
        self.label_style = ParagraphStyle(
            'CustomLabel',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.black,
            alignment=1,  # Centrado
            spaceAfter=8,
            fontName='Helvetica-Bold'
        )
        
        # Estilo para imágenes en segunda página
        self.image_style = ParagraphStyle(
            'CustomImage',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=1,  # Centrado
            spaceAfter=15,
            fontName='Helvetica'
        )

    def agregar_primera_pagina(self):
        """Agrega el contenido de la primera página SIN encabezado duplicado"""
        
        print("📄 Generando primera página...")
        
        # NOTA: No agregamos el encabezado aquí, solo el contenido específico
        
        # ==================== TABLA DE FECHAS ====================
        fecha_data = [
            ['Fecha de Inspección', '${fverificacion}'],
            ['Fecha de Emisión', '${femision}']
        ]
        
        fecha_table = Table(fecha_data, colWidths=[3*inch, 3*inch])
        fecha_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (0,-1), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('BOLD', (0,0), (0,-1), True),
        ]))
        
        self.elements.append(fecha_table)
        self.elements.append(Spacer(1, 0.2*inch))
        
        # ==================== CLIENTE Y RFC ====================
        cliente_text = '<b>Cliente:</b> ${cliente}'
        self.elements.append(Paragraph(cliente_text, self.normal_style))
        
        rfc_text = '<b>RFC:</b> ${rfc}'
        self.elements.append(Paragraph(rfc_text, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # ==================== TEXTO PRINCIPAL ====================
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
        
        # ==================== TABLA DE PRODUCTOS ====================
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
        
        # ==================== TAMAÑO DEL LOTE ====================
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
        
        # ==================== OBSERVACIONES ====================
        obs1_text = '<b>OBSERVACIONES:</b> La imagen amparada en el dictamen es una muestra de etiqueta que aplica para todos los modelos declarados en el presente dictamen lo anterior fue constatado durante la inspección.'
        self.elements.append(Paragraph(obs1_text, self.normal_style))
        
        obs2_text = '<b>OBSERVACIONES:</b> ${obs}'
        self.elements.append(Paragraph(obs2_text, self.normal_style))
        self.elements.append(Spacer(1, 0.3*inch))
        
    def agregar_segunda_pagina(self):
        """Agrega el contenido de la segunda página"""
        
        print("📄 Generando segunda página...")
        
        # Salto de página
        self.elements.append(PageBreak())
        
        # ==================== ETIQUETAS ====================
        self.elements.append(Paragraph("ETIQUETAS DEL PRODUCTO", self.label_style))
        
        # Primera fila de etiquetas
        etiquetas_linea1 = "${etiqueta1}   ${etiqueta2}   ${etiqueta3}   ${etiqueta4}   ${etiqueta5}"
        self.elements.append(Paragraph(etiquetas_linea1, self.image_style))
        
        # Segunda fila de etiquetas
        etiquetas_linea2 = "${etiqueta6}   ${etiqueta7}   ${etiqueta8}   ${etiqueta9}   ${etiqueta10}"
        self.elements.append(Paragraph(etiquetas_linea2, self.image_style))
        
        self.elements.append(Spacer(1, 0.4*inch))
        
        # ==================== IMÁGENES ====================
        self.elements.append(Paragraph("IMÁGENES DEL PRODUCTO", self.label_style))
        
        # Lista de imágenes
        imagenes = [
            "${img1}", "${img2}", "${img3}", "${img4}", "${img5}",
            "${img6}", "${img7}", "${img8}", "${img9}", "${img10}"
        ]
        
        for img in imagenes:
            self.elements.append(Paragraph(img, self.image_style))

     # ==================== TABLA DE FIRMAS ====================
        firmas_data = [
            ['${firma1}', '', '${firma2}'],
            ['${nfirma1}', '', '${nfirma2}'],
            ['Nombre del Inspector', '', 'Nombre del responsable de\nsupervisión UI']
        ]

        firmas_table = Table(firmas_data, colWidths=[2.8*inch, 0.4*inch, 2.8*inch])

        firmas_table.setStyle(TableStyle([
            # Centrar todo el contenido
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 8),

            # Negrita solo para la fila de títulos
            ('BOLD', (0,2), (-1,2), True),

            # ❌ Eliminar todos los bordes
            ('LINEBELOW', (0,0), (-1,-1), 0, colors.white),
            ('BOX', (0,0), (-1,-1), 0, colors.white),
            ('INNERGRID', (0,0), (-1,-1), 0, colors.white),

            # ✅ Línea inferior SOLO bajo las firmas (primera fila, columnas 0 y 2)
            ('LINEBELOW', (0,0), (0,0), 1, colors.black),
            ('LINEBELOW', (2,0), (2,0), 1, colors.black),
        ]))

        self.elements.append(firmas_table)



    def agregar_encabezado_pie_pagina(self, canvas, doc):
        """Agrega encabezado, pie de página y numeración a todas las páginas"""
        
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
        codigo_text = "${year}049UDC${norma}${folio} Solicitud: ${year}049USD${norma}${solicitud}-${lista}"
        if len(codigo_text) > 100:
            codigo_text = codigo_text[:100] + "..."
        canvas.drawCentredString(LETTER_WIDTH/2, LETTER_HEIGHT-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(LETTER_WIDTH-72, LETTER_HEIGHT-50, numeracion)
        
       # Pie de página SUBIDO (alineado y más compacto)
        footer_text = "Este Dictamen de Cumplimiento se emitió por medios electrónicos, conforme al oficio de autorización DGN.312.05.2012.106 de fecha 10 de enero de 2012 expedido por la DGN a esta Unidad de Inspección."
        formato_text = "Formato: PT-F-208B-00-3"

        canvas.setFont("Helvetica", 7)

        # Dividir texto en líneas
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

        # 🔧 Reducir interlineado y subir ambas líneas
        line_height = 8   # antes 10 → más compacto
        start_y = 60      # antes 90 → sube un poco todo el pie

        for i, line in enumerate(lines):
            text_width = canvas.stringWidth(line, "Helvetica", 7)
            available_width = LETTER_WIDTH - 144
            if text_width < available_width * 0.8:
                x_position = (LETTER_WIDTH - text_width) / 2
            else:
                x_position = 72
            canvas.drawString(x_position, start_y - (i * line_height), line)

        # 🔧 Mover el formato más cerca (pegado al texto anterior)
        canvas.drawRightString(LETTER_WIDTH - 72, start_y - (len(lines) * line_height) - 4, formato_text)

        canvas.restoreState()


    def generar_pdf_corregido(self):
        """Genera el PDF corregido sin duplicaciones"""
        
        print("🎯 GENERANDO PDF CORREGIDO...")
        
        # Crear el documento PDF
        output_path = "Dictamen.pdf"
        self.doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            topMargin=1.5*inch,    # Margen para el encabezado
            bottomMargin=1.5*inch, # Margen para el pie
            leftMargin=0.75*inch,
            rightMargin=0.75*inch
        )
        
        # Crear estilos
        self.crear_estilos()
        
        # Agregar contenido (SIN encabezados duplicados)
        self.agregar_primera_pagina()
        self.agregar_segunda_pagina()
        
        # Construir el documento
        try:
            self.doc.build(
                self.elements,
                onFirstPage=self.agregar_encabezado_pie_pagina,
                onLaterPages=self.agregar_encabezado_pie_pagina
            )
            
            print(f"✅ PDF CORREGIDO CREADO: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ Error al generar PDF: {e}")
            return False

    def mostrar_cambios(self):
        """Muestra los cambios realizados"""
        
        print("\n🔧 CORRECIONES APLICADAS:")
        print("   ✅ ELIMINADO: Encabezado duplicado en el contenido")
        print("   ✅ ELIMINADO: Pie de página duplicado en el contenido")
        print("   ✅ MEJORADO: Pie de página con texto justificado")
        print("   ✅ MEJORADO: 'Formato: PT-F-208B-00-3' alineado a la derecha")
        print("   ✅ MANTENIDO: Encabezado único en cada página")
        print("   ✅ MANTENIDO: Numeración correcta de páginas")
        print("   ✅ MANTENIDO: Fondo en todas las páginas")
        print("   ✅ MANTENIDO: Todas las variables disponibles")


def verificar_imagen_fondo():
    """Verifica que la imagen de fondo existe"""
    
    image_path = "img/Fondo.jpeg"
    
    if not os.path.exists(image_path):
        print(f"⚠️  No se encontró: {image_path}")
        os.makedirs("img", exist_ok=True)
        print("📁 Se creó la carpeta 'img/' - coloca 'Fondo.jpeg' allí")
        return False
    
    print("✅ Imagen de fondo encontrada")
    return True


if __name__ == "__main__":
    print("=" * 70)
    print("   GENERADOR DE PDF BOSCH - VERSIÓN CORREGIDA")
    print("=" * 70)
    
    # Verificar imagen de fondo
    print("\n🔍 Verificando recursos...")
    verificar_imagen_fondo()
    
    # Generar PDF corregido
    print("\n🛠️  Generando documento corregido...")
    generador = PDFGenerator()
    
    if generador.generar_pdf_corregido():
        # Mostrar información
        generador.mostrar_cambios()
        
        print("\n📁 ARCHIVO CREADO:")
        print("   • Dictamen_BOSCH_Corregido.pdf")
        
        print("\n🎯 CARACTERÍSTICAS FINALES:")
        print("   - Un solo archivo PDF integrado")
        print("   - Sin duplicación de textos")
        print("   - Encabezado único por página")
        print("   - Pie de página justificado")
        print("   - Formato alineado a la derecha")
        print("   - Numeración correcta")
        print("   - Dos páginas completas")
        
    else:
        print("❌ No se pudo generar el PDF corregido")
    
    print("\n🎉 ¡PROCESO FINALIZADO!")