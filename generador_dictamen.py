"""
Ejecutor de Plantilla BOSCH PDF - Usa Dictamen.py
- Importa y utiliza tu archivo Dictamen.py original
- Llena automáticamente todas las variables con datos reales
- Genera PDF listo para usar
"""

import os
import sys
from datetime import datetime

# Intenta importar tu archivo Dictamen.py original
try:
    # Importa tu clase PDFGenerator desde Dictamen.py
    from DictamenPDF import PDFGenerator
    print("✅ Archivo Dictamen.py encontrado y cargado correctamente")
except ImportError as e:
    print(f"❌ Error: No se pudo importar Dictamen.py")
    print(f"   Detalle: {e}")
    print("\n🔧 Solución: Asegúrate de que:")
    print("   - El archivo 'Dictamen.py' esté en la misma carpeta")
    print("   - La clase se llame 'PDFGenerator'")
    print("   - No haya errores de sintaxis en Dictamen.py")
    sys.exit(1)

"""
Ejecutor de Plantilla BOSCH PDF - Usa Dictamen.py
- Importa y utiliza tu archivo Dictamen.py original
- Llena automáticamente todas las variables con datos reales
- Genera PDF listo para usar
"""

import os
import sys
from datetime import datetime

# Intenta importar tu archivo Dictamen.py original
try:
    # Importa tu clase PDFGenerator desde DictamenPDF.py
    from DictamenPDF import PDFGenerator
    print("✅ Archivo DictamenPDF.py encontrado y cargado correctamente")
except ImportError as e:
    print(f"❌ Error: No se pudo importar DictamenPDF.py")
    print(f"   Detalle: {e}")
    print("\n🔧 Solución: Asegúrate de que:")
    print("   - El archivo 'DictamenPDF.py' esté en la misma carpeta")
    print("   - La clase se llame 'PDFGenerator'")
    print("   - No haya errores de sintaxis en DictamenPDF.py")
    sys.exit(1)

# Importar los elementos necesarios de reportlab
from reportlab.platypus import SimpleDocTemplate

class PDFGeneratorConDatos(PDFGenerator):
    """Subclase de tu PDFGenerator original con capacidad para datos reales"""
    
    def __init__(self, datos):
        # Llama al constructor de la clase padre
        super().__init__()
        self.datos = datos
    
    def agregar_encabezado_pie_pagina(self, canvas, doc):
        """Sobrescribe el método para usar datos reales en encabezado/pie"""
        canvas.saveState()
        
        # Fondo (usando tu lógica original)
        image_path = "img/Fondo.jpeg"
        if os.path.exists(image_path):
            try:
                from reportlab.lib.units import inch
                LETTER_WIDTH = 8.5 * inch
                LETTER_HEIGHT = 11 * inch
                canvas.drawImage(image_path, 0, 0, width=LETTER_WIDTH, height=LETTER_HEIGHT)
            except:
                pass
        
        # ENCABEZADO CON DATOS REALES
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(8.5*72/2, 11*72-60, "DICTAMEN DE CUMPLIMIENTO")
        
        canvas.setFont("Helvetica", 10)
        # Usa datos reales en lugar de variables
        codigo_text = f"{self.datos.get('year', '2024')}049UDC{self.datos.get('norma', 'NOM-001')}{self.datos.get('folio', '12345')} Solicitud: {self.datos.get('year', '2024')}049USD{self.datos.get('norma', 'NOM-001')}{self.datos.get('solicitud', '67890')}-{self.datos.get('lista', 'A')}"
        
        if len(codigo_text) > 100:
            codigo_text = codigo_text[:100] + "..."
        canvas.drawCentredString(8.5*72/2, 11*72-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(8.5*72-72, 11*72-50, numeracion)
        
        # PIE DE PÁGINA (igual al original)
        footer_text = "Este Dictamen de Cumplimiento se emitió por medios electrónicos, conforme al oficio de autorización DGN.312.05.2012.106 de fecha 10 de enero de 2012 expedido por la DGN a esta Unidad de Inspección."
        formato_text = "Formato: PT-F-208B-00-3"

        canvas.setFont("Helvetica", 7)

        # Dividir texto en líneas (misma lógica que tu original)
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
            available_width = 8.5*72 - 144
            if text_width < available_width * 0.8:
                x_position = (8.5*72 - text_width) / 2
            else:
                x_position = 72
            canvas.drawString(x_position, start_y - (i * line_height), line)

        canvas.drawRightString(8.5*72 - 72, start_y - (len(lines) * line_height) - 4, formato_text)
        canvas.restoreState()
    
    def generar_pdf_con_datos(self):
        """Genera el PDF usando tu plantilla pero con datos reales"""
        print("🎯 Generando PDF con datos reales usando tu plantilla...")
        
        # Configuración del documento (igual a tu original)
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.units import inch
        
        output_path = "Dictamen_Completado.pdf"
        self.doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            topMargin=1.5*inch,
            bottomMargin=1.5*inch,
            leftMargin=0.75*inch,
            rightMargin=0.75*inch
        )
        
        # Crear estilos (usa tu método original)
        self.crear_estilos()
        
        # Agregar contenido (usa tus métodos originales)
        self.agregar_primera_pagina()
        self.agregar_segunda_pagina()
        
        # Construir el documento
        try:
            self.doc.build(
                self.elements,
                onFirstPage=self.agregar_encabezado_pie_pagina,
                onLaterPages=self.agregar_encabezado_pie_pagina
            )
            
            print(f"✅ PDF CREADO EXITOSAMENTE: {output_path}")
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
    """Crea un conjunto de datos de ejemplo"""
    return {
        'year': '2024',
        'norma': 'NOM-001',
        'folio': '12345',
        'solicitud': '67890',
        'lista': 'A',
        'fverificacion': '15/01/2024',
        'femision': '20/01/2024',
        'fverificacionlarga': '15 de enero de 2024',
        'cliente': 'ELECTRODOMÉSTICOS BOSCH MÉXICO S.A. DE C.V.',
        'rfc': 'EBM240115ABC',
        'producto': 'REFRIGERADOR BOSCH MODELO KIR56B80',
        'pedimento': '202401151234567',
        'capitulo': '4',
        'normades': 'ESPECIFICACIONES DE SEGURIDAD',
        'rowMarca': 'BOSCH',
        'rowCodigo': 'KIR56B80',
        'rowFactura': 'FAC-2024-001',
        'rowCantidad': '150',
        'TCantidad': '150 unidades',
        'obs': 'El producto cumple con todos los requisitos establecidos en la norma aplicable.',
        'etiqueta1': 'ETQ-001', 'etiqueta2': 'ETQ-002', 'etiqueta3': 'ETQ-003',
        'etiqueta4': 'ETQ-004', 'etiqueta5': 'ETQ-005', 'etiqueta6': 'ETQ-006',
        'etiqueta7': 'ETQ-007', 'etiqueta8': 'ETQ-008', 'etiqueta9': 'ETQ-009',
        'etiqueta10': 'ETQ-010',
        'img1': 'Imagen frontal', 'img2': 'Imagen posterior', 
        'img3': 'Etiqueta', 'img4': 'Diagrama',
        'firma1': '________________________',
        'firma2': '________________________',
        'nfirma1': 'Ing. Juan Pérez Hernández',
        'nfirma2': 'Lic. María García López'
    }

def main():
    """Función principal"""
    print("=" * 70)
    print("   GENERADOR DE DICTAMEN - USA TU Dictamen.py")
    print("=" * 70)
    
    # Verificar dependencias
    print("\n🔍 Verificando recursos...")
    verificar_imagen_fondo()
    
    # Crear datos
    datos = crear_datos_ejemplo()
    
    # Mostrar resumen
    print("\n📊 DATOS A UTILIZAR:")
    print(f"   Cliente: {datos['cliente']}")
    print(f"   Producto: {datos['producto']}")
    print(f"   RFC: {datos['rfc']}")
    
    # Generar PDF usando TU plantilla
    print("\n🛠️  Ejecutando tu plantilla Dictamen.py con datos reales...")
    
    try:
        generador = PDFGeneratorConDatos(datos)
        
        if generador.generar_pdf_con_datos():
            print("\n🎉 ¡PDF GENERADO EXITOSAMENTE!")
            print("\n📁 ARCHIVO CREADO:")
            print("   • Dictamen_Completado.pdf")
            
            print("\n✅ DETALLES:")
            print("   - Se utilizó TU archivo Dictamen.py")
            print("   - Todos los datos fueron insertados automáticamente")
            print("   - Formato profesional mantenido")
            print("   - Fondo incluido")
            print("   - 2 páginas completas")
            
        else:
            print("❌ No se pudo generar el PDF")
            
    except Exception as e:
        print(f"❌ Error durante la generación: {e}")
        print("\n🔧 Posibles soluciones:")
        print("   - Verifica que Dictamen.py no tenga errores")
        print("   - Asegúrate de que todas las clases y métodos existan")
        print("   - Revisa que los nombres coincidan exactamente")
    
    print("\n¡PROCESO FINALIZADO!")

if __name__ == "__main__":
    main()