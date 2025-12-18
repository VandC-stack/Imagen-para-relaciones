# -- Constancia de Conformidad -- #
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import os
import json
from datetime import datetime


class ConstanciaPDFGenerator:
    def __init__(self, datos, path_firmas_json="data/Firmas.json"):
        """Inicializa el generador de PDF para constancia de conformidad"""
        self.datos = datos
        self.width, self.height = letter
        self.firmas_data = self.cargar_firmas(path_firmas_json)
        self.cursor_y = self.height - 40

    def cargar_firmas(self, path_firmas_json):
        """Carga los datos de las firmas desde el archivo JSON"""
        try:
            if os.path.exists(path_firmas_json):
                with open(path_firmas_json, 'r', encoding='utf-8') as f:
                    return json.load(f)
            # Fallback a APPDATA
            try:
                alt_base = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GeneradorDictamenes')
                alt_path = os.path.join(alt_base, os.path.basename(path_firmas_json))
                if os.path.exists(alt_path):
                    with open(alt_path, 'r', encoding='utf-8') as f:
                        return json.load(f)
            except Exception:
                pass
            return []
        except Exception as e:
            print(f"⚠️ Error al cargar firmas: {e}")
            return []

    def dibujar_fondo(self, c):
        """Dibuja la imagen de fondo"""
        fondo_path = "img/Fondo.jpeg"
        if os.path.exists(fondo_path):
            try:
                img = ImageReader(fondo_path)
                c.drawImage(img, 0, 0, width=self.width, height=self.height)
            except Exception as e:
                print(f"⚠️ Error al cargar imagen de fondo: {e}")

    def dibujar_paginacion(self, c):
        """Dibuja la paginación"""
        c.setFont("Helvetica", 8)
        c.drawRightString(self.width - 30, self.height - 30, "PT-F-208C-00-1")
        c.drawRightString(self.width - 40, self.height - 40, "Página 1")

    def dibujar_encabezado(self, c):
        """Encabezado centrado"""
        titulo = "CONSTANCIA DE CONFORMIDAD"
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(self.width / 2, self.cursor_y, titulo)
        self.cursor_y -= 50

    def dibujar_datos_basicos(self, c):
        """Dibuja datos básicos (Folio, Fecha, Cliente, RFC)"""
        x_start = 25 * mm
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, f"Folio: {self.datos.get('folio_constancia','')}")
        self.cursor_y -= 15
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, f"Fecha de Emisión: {self.datos.get('fecha_emision','')}")
        self.cursor_y -= 15
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, f"Cliente: {self.datos.get('cliente','')}")
        self.cursor_y -= 15
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, f"R.F.C.: {self.datos.get('rfc','')}")
        self.cursor_y -= 25

    def dibujar_cuerpo_legal(self, c):
        """Dibuja el texto legal de la constancia"""
        x_start = 25 * mm
        max_width = 165 * mm
        
        norma = self.datos.get('norma','')
        nombre_norma = self.datos.get('nombre_norma','')
        
        texto_legal = (
            f"De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, 68 y 140 "
            f"de la Ley de Infraestructura de la Calidad; 50 del Reglamento de la Ley Federal de Metrología y Normalización; "
            f"Punto 2.4.8 Fracción I ACUERDO por el que la Secretaría de Economía emite Reglas y criterios de carácter general "
            f"en materia de comercio exterior; publicado en el Diario Oficial de la Federación el 09 de mayo de 2022 y posteriores "
            f"modificaciones; esta Unidad de Inspección, hace constar que la Información Comercial contenida en el producto cuya "
            f"etiqueta se muestra aparece en esta Constancia, cumple con la Norma Oficial Mexicana {norma} ({nombre_norma}), "
            f"modificación del 27 de marzo de 2020, ACUERDO por el cual se establecen los Criterios para la implementación, "
            f"verificación y vigilancia, así como para la evaluación de la conformidad de la Modificación a la Norma Oficial Mexicana "
            f"{norma} ({nombre_norma}), publicada el 27 de marzo de 2020 y la Nota Aclaratoria que emiten la Secretaría de Economía "
            f"y la Secretaría de Salud a través de la Comisión Federal para la Protección contra Riesgos Sanitarios a la Modificación "
            f"a la Norma Oficial Mexicana {norma}."
        )
        
        c.setFont("Helvetica", 9)
        text_obj = c.beginText(x_start, self.cursor_y)
        lineas = self._dividir_texto(c, texto_legal, max_width, font_size=9)
        for linea in lineas:
            text_obj.textLine(linea)
        c.drawText(text_obj)
        self.cursor_y = text_obj.getY() - 20

    def dibujar_condiciones(self, c):
        """Dibuja las condiciones de la constancia"""
        x_start = 25 * mm
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Condiciones de la Constancia")
        self.cursor_y -= 15
        
        condiciones = [
            "1. Este documento sólo ampara la información contenida en el producto cuya etiqueta se presenta en esta Constancia.",
            "2. Cualquier modificación a la etiqueta debe ser sometida a la consideración de la Unidad de Inspección Acreditada y Aprobada en los términos de la Ley de Infraestructura de la Calidad, para que inspeccione su cumplimiento con la Norma Oficial Mexicana aplicable.",
            "3. Esta Constancia sólo ampara el cumplimiento con la Norma Oficial Mexicana aplicable."
        ]
        
        c.setFont("Helvetica", 9)
        for condicion in condiciones:
            text_obj = c.beginText(x_start, self.cursor_y)
            lineas = self._dividir_texto(c, condicion, 160*mm, font_size=9)
            for linea in lineas:
                text_obj.textLine(linea)
            c.drawText(text_obj)
            self.cursor_y = text_obj.getY() - 5
        
        self.cursor_y -= 10

    def dibujar_producto(self, c):
        """Dibuja info del producto"""
        x_start = 25 * mm
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, f"Producto: {self.datos.get('producto','')}")
        self.cursor_y -= 20

    def dibujar_tabla_relacion(self, c):
        """Dibuja tabla de relación marca/modelo"""
        x_start = 25 * mm
        col_w = 75 * mm
        row_h = 12
        
        # Encabezado
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "RELACIÓN CORRESPONDIENTE")
        self.cursor_y -= 15
        
        # Bordes tabla
        y_header = self.cursor_y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, y_header, "MARCA")
        c.drawString(x_start + col_w, y_header, "MODELO")
        
        # Líneas divisoras
        c.line(x_start, y_header + 5, x_start + col_w + col_w, y_header + 5)
        c.line(x_start, y_header - 15, x_start + col_w + col_w, y_header - 15)
        c.line(x_start, y_header - 45, x_start + col_w + col_w, y_header - 45)
        
        # Línea vertical divisoria
        c.line(x_start + col_w, y_header + 5, x_start + col_w, y_header - 45)
        
        # Llenar datos si existen
        marca = self.datos.get('marca','')
        modelo = self.datos.get('modelo','')
        c.setFont("Helvetica", 9)
        c.drawString(x_start + 5, y_header - 10, marca)
        c.drawString(x_start + col_w + 5, y_header - 10, modelo)
        
        self.cursor_y = y_header - 50

    def dibujar_observaciones(self, c):
        """Dibuja observaciones finales"""
        x_start = 25 * mm
        max_width = 165 * mm
        
        obs = (
            "OBSERVACIONES: EN CUMPLIMIENTO CON LOS PUNTOS 4.2.6 Y 4.2.7 DE LA NORMA LOS DATOS DE FECHA DE CONSUMO "
            "PREFERENTE Y LOTE SE ENCUENTRAN DECLARADOS EN EL ENVASE DEL PRODUCTO. ESTE PRODUCTO FUE INSPECCIONADO EN "
            "CUMPLIMIENTO BAJO LA FASE 2 DE LA NOM CON VIGENCIA AL 31 DE DICIEMBRE DE 2027 Y FASE 3 DE LA NOM CON "
            "ENTRADA EN VIGOR A PARTIR DEL 01 DE ENERO DEL 2028."
        )
        
        c.setFont("Helvetica", 8)
        text_obj = c.beginText(x_start, self.cursor_y)
        lineas = self._dividir_texto(c, obs, max_width, font_size=8)
        for linea in lineas:
            text_obj.textLine(linea)
        c.drawText(text_obj)
        self.cursor_y = text_obj.getY() - 30

    def dibujar_firma(self, c):
        """Dibuja sección de firma"""
        x_start = 25 * mm
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_start, self.cursor_y, "Atentamente")
        self.cursor_y -= 20
        
        # Intentar cargar firma
        firma_path = "Firmas/AFLORES.png"
        if os.path.exists(firma_path):
            try:
                c.drawImage(firma_path, x_start, self.cursor_y - 15*mm, width=45*mm, height=15*mm, preserveAspectRatio=True)
            except Exception:
                pass
        
        self.cursor_y -= 20*mm
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_start, self.cursor_y, "ARTURO FLORES GÓMEZ")
        self.cursor_y -= 15
        c.setFont("Helvetica", 9)
        c.drawString(x_start, self.cursor_y, "Inspector Acreditado")

    def _dividir_texto(self, c, texto, max_width, font_size=10):
        """Divide texto en líneas según el ancho máximo"""
        palabras = texto.split()
        lineas = []
        linea_actual = ""
        
        for palabra in palabras:
            test_linea = f"{linea_actual} {palabra}" if linea_actual else palabra
            if c.stringWidth(test_linea, "Helvetica", font_size) < max_width:
                linea_actual = test_linea
            else:
                if linea_actual:
                    lineas.append(linea_actual)
                linea_actual = palabra
        
        if linea_actual:
            lineas.append(linea_actual)
        
        return lineas

    def generar(self, nombre_archivo="Constancia.pdf"):
        """Genera el archivo PDF"""
        c = canvas.Canvas(nombre_archivo, pagesize=letter)
        
        # Resetear cursor
        self.cursor_y = self.height - 40
        
        # Dibujar fondo
        try:
            self.dibujar_fondo(c)
        except Exception:
            pass
        
        # Dibujar paginación
        try:
            self.dibujar_paginacion(c)
        except Exception:
            pass
        
        # Dibujar secciones
        self.dibujar_encabezado(c)
        self.dibujar_datos_basicos(c)
        self.dibujar_cuerpo_legal(c)
        self.dibujar_condiciones(c)
        self.dibujar_producto(c)
        self.dibujar_tabla_relacion(c)
        self.dibujar_observaciones(c)
        self.dibujar_firma(c)
        
        # Guardar
        c.save()
        print(f"✅ Constancia generada exitosamente: {nombre_archivo}")
        return nombre_archivo


def generar_constancia_pdf(datos, ruta_salida="Constancia.pdf"):
    """Genera un PDF de constancia con los datos proporcionados"""
    datos_requeridos = [
        'folio_constancia', 'fecha_emision', 'cliente', 'rfc', 'norma', 'nombre_norma',
        'producto', 'marca', 'modelo'
    ]
    
    for campo in datos_requeridos:
        if campo not in datos:
            datos[campo] = ''
    
    generador = ConstanciaPDFGenerator(datos)
    return generador.generar(ruta_salida)


def generar_constancia_desde_visita(folio_visita=None, ruta_salida=None):
    """Genera una constancia a partir de datos de visita y tabla_de_relacion.json"""
    base_dir = os.path.join(os.path.dirname(__file__), '..')
    data_dir = os.path.join(base_dir, 'data')
    historial_path = os.path.join(data_dir, 'historial_visitas.json')
    tabla_path = os.path.join(data_dir, 'tabla_de_relacion.json')
    
    # Cargar historial
    if not os.path.exists(historial_path):
        raise FileNotFoundError(f"No se encontró {historial_path}")
    
    with open(historial_path, 'r', encoding='utf-8') as f:
        historial = json.load(f)
    
    visitas = historial.get('visitas', []) if isinstance(historial, dict) else historial
    visita = None
    if folio_visita:
        for v in visitas:
            if v.get('folio_visita') == folio_visita:
                visita = v
                break
    if visita is None and visitas:
        visita = visitas[-1]
    
    if visita is None:
        raise ValueError('No hay visitas en el historial')
    
    # Cargar Clientes.json para RFC
    clientes = {}
    clientes_path = os.path.join(data_dir, 'Clientes.json')
    try:
        if os.path.exists(clientes_path):
            with open(clientes_path, 'r', encoding='utf-8') as f:
                cl = json.load(f)
                if isinstance(cl, list):
                    for c in cl:
                        clientes[c.get('CLIENTE','').upper()] = c
    except Exception:
        pass
    
    # Cargar Normas.json para nombre_norma
    normas = {}
    normas_path = os.path.join(data_dir, 'Normas.json')
    try:
        if os.path.exists(normas_path):
            with open(normas_path, 'r', encoding='utf-8') as f:
                nm = json.load(f)
                if isinstance(nm, list):
                    for n in nm:
                        normas[n.get('NOM','')] = n.get('NOMBRE','')
    except Exception:
        pass
    
    # Obtener primer registro de tabla para datos
    producto = ''
    marca = ''
    modelo = ''
    fecha_verificacion = None
    if os.path.exists(tabla_path):
        try:
            with open(tabla_path, 'r', encoding='utf-8') as f:
                tabla = json.load(f)
                if isinstance(tabla, list) and tabla:
                    r = tabla[0]
                    producto = r.get('DESCRIPCION','')
                    marca = r.get('MARCA','')
                    modelo = r.get('CODIGO','')
                    fecha_verificacion = r.get('FECHA DE VERIFICACION','')
        except Exception:
            pass
    
    norma_str = visita.get('norma','').split(',')[0].strip() if visita.get('norma') else ''
    nombre_norma = normas.get(norma_str, '')
    
    cliente = visita.get('cliente','')
    cliente_info = clientes.get(cliente.upper(), {})
    rfc = cliente_info.get('RFC','')
    
    datos_constancia = {
        'folio_constancia': visita.get('folio_visita',''),
        'fecha_emision': fecha_verificacion or visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'cliente': cliente,
        'rfc': rfc,
        'norma': norma_str,
        'nombre_norma': nombre_norma,
        'producto': producto,
        'marca': marca,
        'modelo': modelo
    }
    
    if not ruta_salida:
        fol = visita.get('folio_visita', 'constancia')
        ruta_salida = os.path.join(base_dir, f'Constancia_{fol}.pdf')
    
    generar_constancia_pdf(datos_constancia, ruta_salida)
    return ruta_salida


if __name__ == "__main__":
    datos = {
        'folio_constancia': 'UCC12345',
        'fecha_emision': '12/12/2025',
        'cliente': 'ARTICULOS DEPORTIVOS DECATHLON SA DE CV',
        'rfc': 'ADD150727S34',
        'norma': 'NOM-004-SE-2021',
        'nombre_norma': 'Información Comercial- etiquetado de productos textiles, prendas de vestir, sus accesorios y ropa de casa',
        'producto': 'ARTICULOS DEPORTIVOS',
        'marca': 'DECATHLON',
        'modelo': 'VAR001'
    }
    
    os.makedirs("img", exist_ok=True)
    os.makedirs("Firmas", exist_ok=True)
    
    generar_constancia_pdf(datos, "Constancia.pdf")




