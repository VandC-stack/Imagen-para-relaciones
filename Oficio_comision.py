from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import os
import json
from datetime import datetime

class OficioPDFGenerator:
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
        c.setFont("Helvetica", 9)
        c.drawRightString(self.width - 20, self.height - 20, "Página 1 de 1")
    
    def dibujar_encabezado(self, c):
        """Encabezado centrado arriba del documento"""
        titulo1 = "OFICIO DE COMISIÓN"
        titulo2 = "PT-F-208W-00-1"

        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(self.width / 2, self.cursor_y, titulo1)
        self.cursor_y -= 20

        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(self.width / 2, self.cursor_y, titulo2)
        self.cursor_y -= 40  # Espacio después del encabezado
    
    def dibujar_tabla_superior(self, c):
        """Tabla superior de 3 columnas x 2 filas, sin bordes"""

        x_start = 25 * mm
        col1_w = 45 * mm   # ancho columna 1 (títulos)
        col2_w = 50 * mm   # ancho columna 2 (valores)
        col3_w = 70 * mm   # ancho columna 3 (normas)

        # ===============================
        # FILA 1
        # ===============================

        # Columna 1 – Título: No. de Oficio
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "No. de Oficio:")

        # Columna 2 – Valor No. oficio
        c.setFont("Helvetica", 10)
        c.drawString(x_start + col1_w, self.cursor_y, self.datos.get('no_oficio', 'AC0001'))

        # Columna 3 – Normas (título)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start + col1_w + col2_w, self.cursor_y, "Normas:")

        self.cursor_y -= 15  # bajar para fila 2

        # ===============================
        # FILA 2
        # ===============================

        # Columna 1 – Título: Fecha de inspección
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Fecha de Inspección:")

        # Columna 2 – Valor fecha inspección
        c.setFont("Helvetica", 10)
        c.drawString(x_start + col1_w, self.cursor_y, self.datos.get('fecha_inspeccion', 'DD/MM/AAAA'))

        # Columna 3 – Lista de normas
        c.setFont("Helvetica", 10)
        
        normas = self.datos.get("normas", [])
        norma_y = self.cursor_y

        for norma in normas:
            c.drawString(x_start + col1_w + col2_w + 5, norma_y, f"• {norma}")
            norma_y -= 10

        # Ajustar cursor según número de normas
        self.cursor_y = min(self.cursor_y, norma_y - 10)

        # Espacio final para evitar empalmes
        self.cursor_y -= 15

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
    
    def dibujar_cuerpo(self, c):
        """Dibuja el cuerpo del texto del oficio"""
        x_start = 25 * mm
        max_width = 165 * mm

        # =============================
        # 1) Texto introductorio
        # =============================
        texto_intro = (
            f"Estimados Señores: De acuerdo a la confirmación de fecha: "
            f"{self.datos.get('fecha_confirmacion', 'DD/MM/YYYY')} "
            f"recibida de su parte vía: {self.datos.get('medio_confirmacion', 'correo electrónico')}, "
            "me permito informarle por esta vía que el Inspector asignado para llevar "
            "a cabo la inspección es/son el/los señor(es): "
        )

        c.setFont("Helvetica", 10)
        text_obj = c.beginText(x_start, self.cursor_y)

        for linea in self._dividir_texto(c, texto_intro, max_width):
            text_obj.textLine(linea)

        c.drawText(text_obj)
        self.cursor_y = text_obj.getY() - 15

        # =============================
        # 2) Lista de inspectores
        # =============================
        inspectores = self.datos.get('inspectores', [])
        for inspector in inspectores:
            c.drawString(x_start + 10*mm, self.cursor_y, f"• {inspector}")
            self.cursor_y -= 12

        self.cursor_y -= 10

        # =============================
        # 3) PÁRRAFOS FINALES EXACTOS
        # =============================

        parrafos = [
            "Quién(es) se encuentra(n) acreditado(s) y es/son el/los único(s) autorizado(s) "
            "para llevar a cabo las actividades propias de inspección objeto de este servicio.",

            "De antemano le agradecemos las facilidades que se le den para llevar a cabo "
            "correctamente las actividades de inspección y se firme de conformidad por el "
            "responsable de atender la inspección este documento.",

            "Cualquier anormalidad o queja durante el servicio comunicarse al: "
            "5531430039 ó arturo.flores@vyc.com.mx",

            "Atentamente"
        ]

        for p in parrafos:
            text_parrafo = c.beginText(x_start, self.cursor_y)
            text_parrafo.setFont("Helvetica", 10)

            lineas = self._dividir_texto(c, p, max_width)
            for linea in lineas:
                text_parrafo.textLine(linea)

            c.drawText(text_parrafo)
            self.cursor_y = text_parrafo.getY() - 15

    def dibujar_firma(self, c):
        """Dibuja la sección de firma en formato:
        Atentamente
        [Firma]
        Arturo Flores Gómez
        """

        x_start = 25 * mm
        max_width = 165 * mm

        # =============================
        # 1) "Atentamente"
        # =============================
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_start, self.cursor_y, "Atentamente")
        self.cursor_y -= 15

        # =============================
        # 2) Imagen de Firma AFLORES.png
        # =============================
        firma_path = "Firmas/AFLORES.png"

        if os.path.exists(firma_path):
            try:
                c.drawImage(
                    firma_path,
                    x_start,                # izquierda
                    self.cursor_y - 18*mm,  # debajo del texto
                    width=45*mm,
                    height=18*mm,
                    preserveAspectRatio=True
                )
            except Exception as e:
                print(f"⚠️ No se pudo cargar la imagen de firma: {e}")
        else:
            print(f"⚠️ No existe la firma en: {firma_path}")

        # Desplazar cursor debajo de la imagen
        self.cursor_y -= 22*mm

        # =============================
        # 3) Nombre del responsable
        # =============================
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_start, self.cursor_y, "ARTURO FLORES GÓMEZ")

        # Espacio final
        self.cursor_y -= 20

    def dibujar_observaciones(self, c):
        """Dibuja las observaciones y número de solicitudes"""
        x_start = 25 * mm
        
        # Observaciones
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Observaciones (Inspector):")
        self.cursor_y -= 15
        
        observaciones = self.datos.get('observaciones', '')
        c.setFont("Helvetica", 10)
        
        # Dividir observaciones si son muy largas
        lineas_obs = self._dividir_texto(c, observaciones, 150*mm)
        for linea in lineas_obs:
            c.drawString(x_start, self.cursor_y, linea)
            self.cursor_y -= 12
        
        self.cursor_y -= 10
        
        # No. de Solicitudes
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "No. de Solicitudes a Inspeccionar:")
        self.cursor_y -= 15
        
        c.setFont("Helvetica", 10)
        c.drawString(x_start, self.cursor_y, self.datos.get('num_solicitudes', ''))
        self.cursor_y -= 40  # Espacio antes de la tabla de firmas
    
    def dibujar_tabla_firmas(self, c):
        """Dibuja la sección de firmas SIN bordes, sin empalmes y con texto pequeño."""
        x_start = 25 * mm

        # Asegurar espacio en la página
        if self.cursor_y < 60:
            self.cursor_y = 60

        ancho_total = 165 * mm

        # Columnas ajustadas
        col_izq = 75 * mm     # Responsable de atender la visita
        col_der = 75 * mm     # Inspector

        # Altura inicial
        y_text = self.cursor_y - 10 * mm

        # Letra más pequeña
        c.setFont("Helvetica", 8)

        # =====================================================
        # Función interna para escribir texto en columnas sin empalme
        # =====================================================
        def write_wrapped(texto, x, y, max_width):
            palabras = texto.split()
            linea = ""
            for palabra in palabras:
                test = (linea + " " + palabra).strip()
                if c.stringWidth(test, "Helvetica", 8) > max_width:
                    c.drawString(x, y, linea)
                    y -= 4 * mm
                    linea = palabra
                else:
                    linea = test
            if linea:
                c.drawString(x, y, linea)
                y -= 4 * mm
            return y

        # =====================================================
        # COLUMNA IZQUIERDA
        # =====================================================
        texto_izq = "Nombre y Firma del responsable de atender la visita"
        y_final_izq = write_wrapped(texto_izq, x_start + 5*mm, y_text, col_izq)

        # =====================================================
        # COLUMNA DERECHA
        # =====================================================
        texto_der = "Nombre y Firma del Inspector"
        y_final_der = write_wrapped(
            texto_der,
            x_start + col_izq + 15*mm,   # separación entre columnas
            y_text,
            col_der
        )

        # Ajustar cursor a la posición más baja de ambas columnas
        self.cursor_y = min(y_final_izq, y_final_der) - 5 * mm

    def _dividir_texto(self, c, texto, max_width):
        """Divide texto en líneas según el ancho máximo"""
        palabras = texto.split()
        lineas = []
        linea_actual = ""
        
        for palabra in palabras:
            test_linea = f"{linea_actual} {palabra}" if linea_actual else palabra
            if c.stringWidth(test_linea, "Helvetica", 10) < max_width:
                linea_actual = test_linea
            else:
                if linea_actual:
                    lineas.append(linea_actual)
                linea_actual = palabra
        
        if linea_actual:
            lineas.append(linea_actual)
        
        return lineas
    
    def generar(self, nombre_archivo="Oficio.pdf"):
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
        
        # Dibujar cuerpo
        self.dibujar_cuerpo(c)
        
        # Dibujar firma (como se solicita)
        self.dibujar_firma(c)
        
        # Dibujar observaciones
        self.dibujar_observaciones(c)
        
        # Dibujar tabla de firmas
        self.dibujar_tabla_firmas(c)
        
        # Guardar PDF
        c.save()
        print(f"✅ PDF generado exitosamente: {nombre_archivo}")
        return nombre_archivo

# Función principal para usar desde tu aplicación
def generar_oficio_pdf(datos, ruta_salida="Oficio.pdf"):
    """
    Genera un PDF de oficio con los datos proporcionados
    """
    # Validar datos mínimos requeridos
    datos_requeridos = [
        'no_oficio', 'fecha_inspeccion', 'normas',
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
    generador = OficioPDFGenerator(datos)
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
    datos_oficio = {
        'no_oficio': datos_visita.get('folio_acta', 'AC' + datos_visita.get('folio_visita', '0000')),
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
        'observaciones': datos_visita.get('observaciones', 'Sin observaciones'),
        'num_solicitudes': datos_visita.get('num_solicitudes', 'Sin especificar')
    }
    
    return datos_oficio

# Ejemplo de uso
if __name__ == "__main__":
    # Datos de ejemplo
    datos_ejemplo = {
        'no_oficio': '2025-001',
        'fecha_inspeccion': '02/12/2025',
        'normas': ['NOM-004-SE-2021'],
        'empresa_visitada': 'ARTICULOS DEPORTIVOS S.A. DE C.V.',
        'calle_numero': 'AVENIDA PRINCIPAL 123',
        'colonia': 'CENTRO',
        'municipio': 'BENITO JUAREZ',
        'ciudad_estado': 'CIUDAD DE MEXICO, CDMX',
        'fecha_confirmacion': '02/12/2025',
        'medio_confirmacion': 'correo electrónico',
        'inspectores': ['GABRIEL RAMIREZ CASTILLO','MARCOS URIEL FLORES GÓMEZ','DAVID ALCANTARA'],
        'observaciones': 'NINGUNA',
        'num_solicitudes': '006916/25'
    }
    
    # Crear carpetas si no existen
    os.makedirs("img", exist_ok=True)
    os.makedirs("Firmas", exist_ok=True)
    os.makedirs("data", exist_ok=True)
    
    # Generar PDF
    generar_oficio_pdf(datos_ejemplo, "Oficio_comision.pdf")


