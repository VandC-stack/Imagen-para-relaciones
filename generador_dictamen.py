"""
Ejecutor de Plantilla BOSCH PDF con Datos Reales
- Llena automáticamente todas las variables
- Genera PDF listo para usar
- Mantiene el formato profesional
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
import os
from PIL import Image
from datetime import datetime

# Tamaño carta en puntos
LETTER_WIDTH = 8.5 * inch
LETTER_HEIGHT = 11 * inch

class PDFGeneratorFilled:
    def __init__(self, datos):
        self.doc = None
        self.elements = []
        self.styles = getSampleStyleSheet()
        self.total_pages = 2
        self.datos = datos  # Diccionario con todos los datos
        
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

    def reemplazar_variables(self, texto):
        """Reemplaza las variables en el texto con los datos reales"""
        for key, value in self.datos.items():
            variable = "${" + key + "}"
            texto = texto.replace(variable, str(value))
        return texto

    def agregar_primera_pagina(self):
        """Agrega el contenido de la primera página con datos reales"""
        
        print("📄 Generando primera página con datos reales...")
        
        # ==================== TABLA DE FECHAS ====================
        fecha_data = [
            ['Fecha de Inspección', self.datos.get('fverificacion', '')],
            ['Fecha de Emisión', self.datos.get('femision', '')]
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
        cliente_text = f"<b>Cliente:</b> {self.datos.get('cliente', '')}"
        self.elements.append(Paragraph(cliente_text, self.normal_style))
        
        rfc_text = f"<b>RFC:</b> {self.datos.get('rfc', '')}"
        self.elements.append(Paragraph(rfc_text, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # ==================== TEXTO PRINCIPAL ====================
        texto_principal = (
            "De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, 68 y 140 de la Ley de Infraestructura de la "
            "Calidad; 50 del Reglamento de la Ley Federal de Metrología y Normalización; Punto 2.4.8 Fracción III ACUERDO por el que la "
            "Secretaría de Economía emite Reglas y criterios de carácter general en materia de comercio exterior; publicado en el Diario Oficial de la "
            "Federación el 09 de mayo de 2022 y posteriores modificaciones; esta Unidad de Inspección a solicitud de la persona moral denominada "
            f"{self.datos.get('cliente', '')} dictamina el Producto: {self.datos.get('producto', '')}; que la mercancía importada bajo el pedimento aduanal No. {self.datos.get('pedimento', '')} de fecha "
            f"{self.datos.get('fverificacionlarga', '')}, fue etiquetada conforme a los requisitos de Información Comercial en el capítulo {self.datos.get('capitulo', '')} de la Norma Oficial Mexicana "
            f"{self.datos.get('norma', '')} {self.datos.get('normades', '')} Cualquier otro requisito establecido en la norma referida, es responsabilidad del titular de este Dictamen."
        )
        
        self.elements.append(Paragraph(texto_principal, self.normal_style))
        self.elements.append(Spacer(1, 0.2*inch))
        
        # ==================== TABLA DE PRODUCTOS ====================
        productos_data = [
            ['MARCA', 'CÓDIGO', 'FACTURA', 'CANTIDAD'],
            [
                self.datos.get('rowMarca', ''), 
                self.datos.get('rowCodigo', ''), 
                self.datos.get('rowFactura', ''), 
                self.datos.get('rowCantidad', '')
            ]
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
            ['TAMAÑO DEL LOTE', self.datos.get('TCantidad', '')]
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
        
        obs2_text = f"<b>OBSERVACIONES:</b> {self.datos.get('obs', '')}"
        self.elements.append(Paragraph(obs2_text, self.normal_style))
        self.elements.append(Spacer(1, 0.3*inch))
        
    def agregar_segunda_pagina(self):
        """Agrega el contenido de la segunda página con datos reales"""
        
        print("📄 Generando segunda página con datos reales...")
        
        # Salto de página
        self.elements.append(PageBreak())
        
        # ==================== ETIQUETAS ====================
        self.elements.append(Paragraph("ETIQUETAS DEL PRODUCTO", self.label_style))
        
        # Primera fila de etiquetas
        etiquetas_linea1 = f"{self.datos.get('etiqueta1', '')}   {self.datos.get('etiqueta2', '')}   {self.datos.get('etiqueta3', '')}   {self.datos.get('etiqueta4', '')}   {self.datos.get('etiqueta5', '')}"
        self.elements.append(Paragraph(etiquetas_linea1, self.image_style))
        
        # Segunda fila de etiquetas
        etiquetas_linea2 = f"{self.datos.get('etiqueta6', '')}   {self.datos.get('etiqueta7', '')}   {self.datos.get('etiqueta8', '')}   {self.datos.get('etiqueta9', '')}   {self.datos.get('etiqueta10', '')}"
        self.elements.append(Paragraph(etiquetas_linea2, self.image_style))
        
        self.elements.append(Spacer(1, 0.4*inch))
        
        # ==================== IMÁGENES ====================
        self.elements.append(Paragraph("IMÁGENES DEL PRODUCTO", self.label_style))
        
        # Lista de imágenes (solo mostrar las que tienen contenido)
        imagenes = [
            self.datos.get('img1', ''), self.datos.get('img2', ''), self.datos.get('img3', ''), 
            self.datos.get('img4', ''), self.datos.get('img5', ''), self.datos.get('img6', ''), 
            self.datos.get('img7', ''), self.datos.get('img8', ''), self.datos.get('img9', ''), 
            self.datos.get('img10', '')
        ]
        
        for img in imagenes:
            if img.strip():  # Solo agregar si no está vacío
                self.elements.append(Paragraph(img, self.image_style))

        # ==================== TABLA DE FIRMAS ====================
        firmas_data = [
            [self.datos.get('firma1', ''), '', self.datos.get('firma2', '')],
            [self.datos.get('nfirma1', ''), '', self.datos.get('nfirma2', '')],
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
        codigo_text = f"{self.datos.get('year', '')}049UDC{self.datos.get('norma', '')}{self.datos.get('folio', '')} Solicitud: {self.datos.get('year', '')}049USD{self.datos.get('norma', '')}{self.datos.get('solicitud', '')}-{self.datos.get('lista', '')}"
        if len(codigo_text) > 100:
            codigo_text = codigo_text[:100] + "..."
        canvas.drawCentredString(LETTER_WIDTH/2, LETTER_HEIGHT-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(LETTER_WIDTH-72, LETTER_HEIGHT-50, numeracion)
        
        # Pie de página
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

        # Reducir interlineado y subir ambas líneas
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

        # Formato alineado a la derecha
        canvas.drawRightString(LETTER_WIDTH - 72, start_y - (len(lines) * line_height) - 4, formato_text)

        canvas.restoreState()

    def generar_pdf_lleno(self):
        """Genera el PDF completamente lleno con datos reales"""
        
        print("🎯 GENERANDO PDF CON DATOS REALES...")
        
        # Crear el documento PDF
        output_path = "Dictamen_Completo.pdf"
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
        
        # Agregar contenido con datos reales
        self.agregar_primera_pagina()
        self.agregar_segunda_pagina()
        
        # Construir el documento
        try:
            self.doc.build(
                self.elements,
                onFirstPage=self.agregar_encabezado_pie_pagina,
                onLaterPages=self.agregar_encabezado_pie_pagina
            )
            
            print(f"✅ PDF COMPLETO CREADO: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ Error al generar PDF: {e}")
            return False

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

def crear_datos_ejemplo():
    """Crea un conjunto de datos de ejemplo para llenar la plantilla"""
    
    datos = {
        # Información general
        'year': '2024',
        'norma': 'NOM-001',
        'folio': '12345',
        'solicitud': '67890',
        'lista': 'A',
        
        # Fechas
        'fverificacion': '15/01/2024',
        'femision': '20/01/2024',
        'fverificacionlarga': '15 de enero de 2024',
        
        # Cliente
        'cliente': 'ELECTRODOMÉSTICOS BOSCH MÉXICO S.A. DE C.V.',
        'rfc': 'EBM240115ABC',
        
        # Producto
        'producto': 'REFIRGERADOR BOSCH MODELO KIR56B80',
        'pedimento': '202401151234567',
        'capitulo': '4',
        'normades': 'ESPECIFICACIONES DE SEGURIDAD',
        
        # Tabla de productos
        'rowMarca': 'BOSCH',
        'rowCodigo': 'KIR56B80',
        'rowFactura': 'FAC-2024-001',
        'rowCantidad': '150',
        'TCantidad': '150 unidades',
        
        # Observaciones
        'obs': 'El producto cumple con todos los requisitos establecidos en la norma aplicable.',
        
        # Etiquetas (10 espacios para etiquetas)
        'etiqueta1': 'ETQ-001',
        'etiqueta2': 'ETQ-002', 
        'etiqueta3': 'ETQ-003',
        'etiqueta4': 'ETQ-004',
        'etiqueta5': 'ETQ-005',
        'etiqueta6': 'ETQ-006',
        'etiqueta7': 'ETQ-007',
        'etiqueta8': 'ETQ-008',
        'etiqueta9': 'ETQ-009',
        'etiqueta10': 'ETQ-010',
        
        # Imágenes (10 espacios para imágenes)
        'img1': 'Imagen frontal del producto',
        'img2': 'Imagen posterior del producto',
        'img3': 'Etiqueta de especificaciones',
        'img4': 'Diagrama de conexiones',
        'img5': '',  # Vacío intencional
        'img6': '',  # Vacío intencional
        'img7': '',  # Vacío intencional
        'img8': '',  # Vacío intencional
        'img9': '',  # Vacío intencional
        'img10': '', # Vacío intencional
        
        # Firmas
        'firma1': '________________________',
        'firma2': '________________________',
        'nfirma1': 'Ing. Juan Pérez Hernández',
        'nfirma2': 'Lic. María García López'
    }
    
    return datos

def crear_datos_personalizados():
    """Permite al usuario ingresar datos personalizados"""
    
    print("\n📝 INGRESO DE DATOS PERSONALIZADOS")
    print("=" * 50)
    
    datos = {}
    
    # Información general
    datos['year'] = input("Año (ej: 2024): ") or '2024'
    datos['norma'] = input("Norma (ej: NOM-001): ") or 'NOM-001'
    datos['folio'] = input("Folio: ") or '12345'
    datos['solicitud'] = input("Solicitud: ") or '67890'
    datos['lista'] = input("Lista: ") or 'A'
    
    # Fechas
    datos['fverificacion'] = input("Fecha de verificación (dd/mm/aaaa): ") or '15/01/2024'
    datos['femision'] = input("Fecha de emisión (dd/mm/aaaa): ") or '20/01/2024'
    datos['fverificacionlarga'] = input("Fecha de verificación larga: ") or '15 de enero de 2024'
    
    # Cliente
    datos['cliente'] = input("Nombre del cliente: ") or 'ELECTRODOMÉSTICOS BOSCH MÉXICO S.A. DE C.V.'
    datos['rfc'] = input("RFC: ") or 'EBM240115ABC'
    
    # Producto
    datos['producto'] = input("Nombre del producto: ") or 'REFIRGERADOR BOSCH MODELO KIR56B80'
    datos['pedimento'] = input("Número de pedimento: ") or '202401151234567'
    datos['capitulo'] = input("Capítulo: ") or '4'
    datos['normades'] = input("Descripción de norma: ") or 'ESPECIFICACIONES DE SEGURIDAD'
    
    # Tabla de productos
    datos['rowMarca'] = input("Marca: ") or 'BOSCH'
    datos['rowCodigo'] = input("Código: ") or 'KIR56B80'
    datos['rowFactura'] = input("Factura: ") or 'FAC-2024-001'
    datos['rowCantidad'] = input("Cantidad: ") or '150'
    datos['TCantidad'] = input("Tamaño del lote: ") or '150 unidades'
    
    # Observaciones
    datos['obs'] = input("Observaciones: ") or 'El producto cumple con todos los requisitos establecidos en la norma aplicable.'
    
    print("\n¿Desea usar valores por defecto para el resto? (s/n): ")
    if input().lower() != 'n':
        # Usar valores por defecto para el resto
        datos_default = crear_datos_ejemplo()
        for key in ['etiqueta1', 'etiqueta2', 'etiqueta3', 'etiqueta4', 'etiqueta5',
                   'etiqueta6', 'etiqueta7', 'etiqueta8', 'etiqueta9', 'etiqueta10',
                   'img1', 'img2', 'img3', 'img4', 'firma1', 'firma2', 'nfirma1', 'nfirma2']:
            datos[key] = datos_default[key]
    else:
        # Ingresar todos los datos manualmente
        for i in range(1, 11):
            datos[f'etiqueta{i}'] = input(f"Etiqueta {i}: ") or f'ETQ-{i:03d}'
        
        for i in range(1, 11):
            datos[f'img{i}'] = input(f"Descripción imagen {i}: ") or ''
        
        datos['firma1'] = '________________________'
        datos['firma2'] = '________________________'
        datos['nfirma1'] = input("Nombre inspector: ") or 'Ing. Juan Pérez Hernández'
        datos['nfirma2'] = input("Nombre responsable: ") or 'Lic. María García López'
    
    return datos

if __name__ == "__main__":
    print("=" * 70)
    print("   EJECUTOR DE PLANTILLA BOSCH PDF - DATOS REALES")
    print("=" * 70)
    
    # Verificar imagen de fondo
    print("\n🔍 Verificando recursos...")
    verificar_imagen_fondo()
    
    # Selección de modo
    print("\n🎯 SELECCIONE MODO DE EJECUCIÓN:")
    print("   1. Usar datos de ejemplo")
    print("   2. Ingresar datos personalizados")
    
    opcion = input("\nSeleccione opción (1/2): ").strip()
    
    if opcion == "2":
        datos = crear_datos_personalizados()
    else:
        datos = crear_datos_ejemplo()
        print("\n✅ Usando datos de ejemplo...")
    
    # Mostrar resumen de datos
    print("\n📊 RESUMEN DE DATOS:")
    print(f"   Cliente: {datos['cliente']}")
    print(f"   Producto: {datos['producto']}")
    print(f"   RFC: {datos['rfc']}")
    print(f"   Fechas: {datos['fverificacion']} / {datos['femision']}")
    
    # Generar PDF con datos reales
    print("\n🛠️  Generando documento con datos reales...")
    generador = PDFGeneratorFilled(datos)
    
    if generador.generar_pdf_lleno():
        print("\n🎉 ¡PDF COMPLETADO EXITOSAMENTE!")
        print("\n📁 ARCHIVO CREADO:")
        print("   • Dictamen_Completo.pdf")
        
        print("\n✅ CARACTERÍSTICAS:")
        print("   - Todas las variables reemplazadas")
        print("   - Formato profesional mantenido")
        print("   - Fondo incluido")
        print("   - 2 páginas completas")
        print("   - Listo para usar y imprimir")
        
    else:
        print("❌ No se pudo generar el PDF completo")
    
    print("\n¡PROCESO FINALIZADO!")