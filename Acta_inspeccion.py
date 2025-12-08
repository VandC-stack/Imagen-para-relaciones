# -- Acta de inspección -- #
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import os
import json
from datetime import datetime
class ActaPDFGenerator:
    def __init__(self, datos, path_firmas_json="data/Firmas.json"):
        """
        Inicializa el generador de PDF para oficio
        """
        self.datos = datos
        self.width, self.height = letter  # 612 x 792 puntos
        self.firmas_data = self.cargar_firmas(path_firmas_json)
        
        # Inicializar posición vertical (desde la parte superior)
        self.cursor_y = self.height - 40  # Empezamos desde arriba con margen
    
    def cargar_firmas(self, path_firmas_json):
        """Carga los datos de las firmas desde el archivo JSON"""
        try:
            if os.path.exists(path_firmas_json):
                with open(path_firmas_json, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return []
        except Exception as e:
            print(f"⚠️ Error al cargar firmas: {e}")
            return []
    
    def dibujar_fondo(self, c):
        """Dibuja la imagen de fondo"""
        fondo_path = "img/Oficios.png"
        if os.path.exists(fondo_path):
            try:
                img = ImageReader(fondo_path)
                c.drawImage(img, 0, 0, width=self.width, height=self.height)
            except Exception as e:
                print(f"⚠️ Error al cargar imagen de fondo: {e}")
    
    def dibujar_paginacion(self, c):
        """Dibuja la paginación"""
        c.setFont("Helvetica", 8)
        c.drawRightString(self.width - 30, self.height - 30, "PT-F-208A-00-1")
        c.drawRightString(self.width - 40, self.height - 40, "Página 1 de 1")
    
    def dibujar_encabezado(self, c):
        """Encabezado centrado arriba del documento"""
        titulo1 = "ACTA DE INSPECCIÓN DE LA UNIDAD DE INSPECCIÓN "
        

        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(self.width / 2, self.cursor_y, titulo1)
        self.cursor_y -= 70

    def dibujar_tabla_superior(self, c):
        """Tabla superior de 4 columnas para ACTA DE INSPECCIÓN (sin bordes)"""

        x_start = 25 * mm
        row_height = 12

        # Anchos de columna
        col_w1 = 40 * mm   # Fecha de inspección (inicio / termino / título)
        col_w2 = 40 * mm   # Día
        col_w3 = 25 * mm   # Hora
        col_w4 = 80 * mm   # Normas

        # =====================================================
        #   ENCABEZADOS
        # =====================================================

        c.setFont("Helvetica-Bold", 10)

        c.drawString(x_start, self.cursor_y, "Fecha de inspección")
        c.drawString(x_start + col_w1, self.cursor_y, "Día")
        c.drawString(x_start + col_w1 + col_w2, self.cursor_y, "Hora")
        c.drawString(x_start + col_w1 + col_w2 + col_w3, self.cursor_y,
                    "Normas para las que solicita el servicio")

        self.cursor_y -= row_height

        # =====================================================
        #   FILA: INICIO
        # =====================================================

        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Inicio")

        c.setFont("Helvetica", 10)
        c.drawString(x_start + col_w1, self.cursor_y,
                    self.datos.get("fecha_inicio", "DD/MM/YYYY"))
        c.drawString(x_start + col_w1 + col_w2, self.cursor_y,
                    self.datos.get("hora_inicio", "09:00"))

        # Normas (primera línea)
        normas = self.datos.get("normas", [])
        if normas:
            c.drawString(x_start + col_w1 + col_w2 + col_w3 + 5,
                        self.cursor_y, normas[0])

        self.cursor_y -= row_height

        # =====================================================
        #   FILA: TÉRMINO
        # =====================================================

        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Término")

        c.setFont("Helvetica", 10)
        c.drawString(x_start + col_w1, self.cursor_y,
                    self.datos.get("fecha_termino", "DD/MM/YYYY"))
        c.drawString(x_start + col_w1 + col_w2, self.cursor_y,
                    self.datos.get("hora_termino", "18:00"))

        # Resto de normas
        self.cursor_y -= row_height

        if len(normas) > 1:
            c.setFont("Helvetica", 10)
            for norma in normas[1:]:
                c.drawString(x_start + col_w1 + col_w2 + col_w3 + 5,
                            self.cursor_y, norma)
                self.cursor_y -= row_height

        # Espacio final
        self.cursor_y -= 10

    def dibujar_datos_empresa(self, c):
        """Dibuja los datos de la empresa visitada sin bordes"""
        x_start = 25 * mm
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Datos del lugar donde se realiza la Inspección de Información Comercial:")
        self.cursor_y -= 25
        
        # Títulos y valores en dos columnas
        campos = [
            ("Empresa Visitada:", self.datos.get('empresa_visitada', '')),
            ("Calle y No.:", self.datos.get('calle_numero', '')),
            ("Colonia o Población:", self.datos.get('colonia', '')),
            ("Municipio o Alcaldía:", self.datos.get('municipio', '')),
            ("Ciudad o Estado:", self.datos.get('ciudad_estado', ''))
        ]
        
        for titulo, valor in campos:
            # Título en negrita
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x_start, self.cursor_y, titulo)
            
            # Valor
            c.setFont("Helvetica", 10)
            # Truncar valor si es muy largo
            if len(valor) > 60:
                valor = valor[:57] + "..."
            c.drawString(x_start + 60*mm, self.cursor_y, valor)
            self.cursor_y -= 15
        
        self.cursor_y -= 20  # Espacio después de la sección
    
    def dibujar_tabla_firmas(self, c):
        """Tabla completa de firmas con 4 columnas por 2 filas, sin bordes"""
        
        # Configuración inicial
        x = 25 * mm
        ancho_total = 165 * mm

        # Anchos de columna
        ancho_nombre = 60 * mm     # Columnas 0 y 2: nombres
        ancho_firma = 22.5 * mm    # Columnas 1 y 3: firmas

        # Posiciones x
        col0 = x
        col1 = col0 + ancho_nombre
        col2 = col1 + ancho_firma
        col3 = col2 + ancho_nombre

        c.setFont("Helvetica", 9)
        y = self.cursor_y - 15

        # ==========================
        # OBTENER INSPECTOR REAL
        # ==========================
        # inspectores = self.datos.get("inspectores", [])
        # inspector_nombre = inspectores[0] if inspectores else ""

        # OBTENER INSPECTOR DESDE DATOS #
        inspector_nombre = (
            self.datos.get("NOMBRE_DE_INSPECTOR")
            or self.datos.get("inspector")
            or self.datos.get("INSPECTOR")
            or ""
        )



        # ==========================================================
        # 1. PRIMERA FILA: CLIENTE E INSPECTOR
        # ==========================================================

        # Títulos del cliente (dos líneas)
        c.drawString(col0, y, "Nombre del cliente o responsable")
        c.drawString(col0, y - 7, "de atender la visita")

        # Títulos de firma del cliente e inspector
        c.drawString(col1, y, "Firma")
        c.drawString(col2, y, "Nombre del Inspector")
        c.drawString(col3, y, "Firma")

        # Espaciado general para bajar a las líneas de firma
        y -= 25

        # Líneas de firma (cliente e inspector)
        c.line(col1, y, col1 + ancho_firma - 5, y)
        c.line(col3, y, col3 + ancho_firma - 5, y)

        # Bajar para escribir nombres
        y -= 25

        # Nombre del inspector
        if inspector_nombre:
            if len(inspector_nombre) > 20:
                c.setFont("Helvetica", 8)
            c.drawString(col2 + 2, y, inspector_nombre)
            c.setFont("Helvetica", 9)

        # ==========================
        # FIRMA DEL INSPECTOR (IMAGEN)
        # ==========================
        firma_path = self.obtener_firma_inspector(inspector_nombre)

        if firma_path:
            try:
                firma_img = ImageReader(firma_path)
                firma_ancho = ancho_firma - 5
                firma_alto = 20
                c.drawImage(
                    firma_img,
                    col3,
                    y - 10,
                    width=firma_ancho,
                    height=firma_alto,
                    preserveAspectRatio=True,
                    mask='auto'
                )
            except Exception as e:
                print(f"⚠️ Error al cargar firma del inspector: {e}")
        else:
            c.line(col3, y, col3 + ancho_firma - 5, y)

        # Espacio para siguiente sección
        y -= 30

        # ==========================================================
        # 2. SEGUNDA FILA: TESTIGOS
        # ==========================================================
        c.drawString(col0, y, "Nombre (Testigo 1)")
        c.drawString(col1, y, "Firma")
        c.drawString(col2, y, "Nombre (Testigo 2)")
        c.drawString(col3, y, "Firma")

        y -= 15

        # Líneas firma testigos
        c.line(col1, y, col1 + ancho_firma - 5, y)
        c.line(col3, y, col3 + ancho_firma - 5, y)

        y -= 40

        # ==========================================================
        # 3. NOTAS Y OBSERVACIONES
        # ==========================================================
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(x + ancho_total / 2, y, "NOTAS Y OBSERVACIONES:")

        y -= 20

        # Observaciones Cliente
        c.setFont("Helvetica", 9)
        c.drawString(col0, y, "Observaciones (Cliente):")
        y -= 10

        for _ in range(3):
            c.line(col0, y, col0 + ancho_total - 10, y)
            y -= 15

        y -= 10

        # Observaciones Inspector
        c.drawString(col0, y, "Observaciones (Inspector):")
        y -= 10

        for _ in range(3):
            c.line(col0, y, col0 + ancho_total - 10, y)
            y -= 15

        y -= 20

        # ==========================================================
        # 4. ACTA Y C.P.
        # ==========================================================
        acta = self.datos.get("acta", "C.P.12345")
        cp = self.datos.get("cp", "CP07890")

        c.drawString(col0, y, f"Acta: {acta}    C.P.: {cp}")

        # Actualizar cursor general
        self.cursor_y = y - 25

    def obtener_firma_inspector(self, inspector_nombre):
        """
        Devuelve la ruta de la firma del inspector según Firmas.json.
        """
        if not inspector_nombre:
            print("⚠️ Nombre de inspector vacío.")
            return None

        inspector_normalizado = inspector_nombre.lower().strip()

        for f in self.firmas_data:

            # DETECTAR EL NOMBRE (incluye 'NOMBRE DE INSPECTOR')
            posible_nombre = (
                f.get("NOMBRE DE INSPECTOR") or
                f.get("nombre") or
                f.get("inspector") or
                f.get("nombre_inspector") or
                f.get("name") or
                ""
            )

            if posible_nombre.lower().strip() == inspector_normalizado:

                # DETECTAR LA RUTA (incluye 'IMAGEN')
                posible_ruta = (
                    f.get("IMAGEN") or
                    f.get("FIRMA") or   # tu JSON trae esto, pero es el código, no la imagen
                    f.get("ruta") or
                    f.get("path") or
                    ""
                )

                # Si la ruta es algo como "ASANCHEZ", convertirla en archivo
                if posible_ruta and "." not in posible_ruta:
                    posible_ruta = os.path.join("Firmas", posible_ruta + ".png")

                if posible_ruta and os.path.exists(posible_ruta):
                    return posible_ruta

        # Buscar por nombre de archivo directo
        nombre_archivo = inspector_nombre.replace(" ", "").upper() + ".png"
        ruta_directa = os.path.join("Firmas", nombre_archivo)

        if os.path.exists(ruta_directa):
            return ruta_directa

        print(f"⚠️ No se encontró firma para: {inspector_nombre}")
        return None

    def generar(self, nombre_archivo="Acta.pdf"):
        """Genera el archivo PDF"""
        c = canvas.Canvas(nombre_archivo, pagesize=letter)
        
        # Resetear cursor al inicio
        self.cursor_y = self.height - 40
        
        # Dibujar fondo (si existe)
        self.dibujar_fondo(c)
        
        # Dibujar paginación
        self.dibujar_paginacion(c)
        
        # Dibujar encabezado
        self.dibujar_encabezado(c)
        
        # Dibujar tabla superior
        self.dibujar_tabla_superior(c)
        
        # Dibujar datos empresa
        self.dibujar_datos_empresa(c)
        
        # Dibujar tabla de firmas
        self.dibujar_tabla_firmas(c)
        
        
        # Guardar PDF
        c.save()
        print(f"✅ PDF generado exitosamente: {nombre_archivo}")
        return nombre_archivo

# Función principal para usar desde tu aplicación
def generar_acta_pdf(datos, ruta_salida="Acta.pdf"):
    """
    Genera un PDF de oficio con los datos proporcionados
    """
    # Validar datos mínimos requeridos
    datos_requeridos = [
        'fecha_inspeccion_inicio', 'fecha_inspeccion_termino', 'normas',
        'empresa_visitada', 'calle_numero', 'colonia',
        'municipio', 'ciudad_estado', 'fecha_confirmacion',
        'medio_confirmacion', 'inspectores', 'observaciones',
        'num_solicitudes'
    ]
    
    # Si falta algún dato, usar valores por defecto
    for campo in datos_requeridos:
        if campo not in datos:
            if campo == 'normas':
                datos[campo] = []
            elif campo == 'inspectores':
                datos[campo] = []
            else:
                datos[campo] = ''
    
    # Asegurar que las normas sean una lista
    if isinstance(datos.get('normas'), str):
        datos['normas'] = [n.strip() for n in datos['normas'].split(',') if n.strip()]
    
    # Generar PDF
    generador = ActaPDFGenerator(datos)
    return generador.generar(ruta_salida)

# Función para preparar datos desde la tabla de relación
def preparar_datos_desde_visita(datos_visita, firmas_json_path="data/Firmas.json"):
    """
    Prepara los datos para el oficio a partir de los datos de una visita
    """
    # Cargar firmas
    firmas_data = []
    if os.path.exists(firmas_json_path):
        with open(firmas_json_path, 'r', encoding='utf-8') as f:
            firmas_data = json.load(f)
    
    # Obtener inspectores
    inspectores = []
    if 'supervisores_tabla' in datos_visita and datos_visita['supervisores_tabla']:
        inspectores = [s.strip() for s in datos_visita['supervisores_tabla'].split(',')]
    elif 'nfirma1' in datos_visita and datos_visita['nfirma1']:
        inspectores = [datos_visita['nfirma1']]
    
    # Preparar datos para el PDF
    datos_acta = {
        'fecha_inspeccion': datos_visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'normas': datos_visita.get('norma', '').split(', ') if datos_visita.get('norma') else [],
        'empresa_visitada': datos_visita.get('cliente', ''),
        'calle_numero': datos_visita.get('direccion', ''),
        'colonia': datos_visita.get('colonia', ''),
        'municipio': datos_visita.get('municipio', ''),
        'ciudad_estado': datos_visita.get('ciudad_estado', ''),
        'fecha_confirmacion': datos_visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'medio_confirmacion': 'correo electrónico',
        'inspectores': inspectores,
        'observaciones': datos_visita.get('observaciones', 'Sin observaciones')
        
    }
    
    return datos_acta

# Ejemplo de uso
if __name__ == "__main__":
    # Datos de ejemplo
    datos = {
        "fecha_inicio": "02/12/2025",
        "hora_inicio": "09:00",
        "fecha_termino": "02/12/2025",
        "hora_termino": "18:00",
        "normas": [
            "NOM-050-SCFI-2004",
            "NOM-142-SSA1/SCFI-2014",
            "NOM-004-SE-2021"
            ],
        'empresa_visitada': 'ARTICULOS DEPORTIVOS S.A. DE C.V.',
        'calle_numero': 'AVENIDA PRINCIPAL 123',
        'colonia': 'CENTRO',
        'municipio': 'BENITO JUAREZ',
        'ciudad_estado': 'CIUDAD DE MEXICO, CDMX',
        'firma_inspector': 'Firmas/AFLORES.png',
        'NOMBRE_DE_INSPECTOR': 'Arturo Flores Gómez',


    }
    # Crear carpetas si no existen
    os.makedirs("img", exist_ok=True)
    os.makedirs("Firmas", exist_ok=True)
    os.makedirs("data", exist_ok=True)
    
    # Generar PDF
    generar_acta_pdf(datos, "Plantillas PDF/Acta_inspeccion.pdf")


