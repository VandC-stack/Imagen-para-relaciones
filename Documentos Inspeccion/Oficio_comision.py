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
            # Intentar ruta proporcionada
            if os.path.exists(path_firmas_json):
                with open(path_firmas_json, 'r', encoding='utf-8') as f:
                    return json.load(f)

            # Fallback: si la app está empacada o la ruta no existe, buscar en APPDATA/GeneradorDictamenes
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
        page_num = getattr(self, 'page_num', 1)
        c.drawRightString(self.width - 20, self.height - 20, f"Página {page_num}")
    
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
        
        # Preparar display de colonia + C.P. si existe
        colonia_val = self.datos.get('colonia', '') or ''
        cp_val = str(self.datos.get('cp', '') or '')
        if cp_val:
            if colonia_val:
                colonia_display = f"{colonia_val}  C.P.: {cp_val}"
            else:
                colonia_display = f"C.P.: {cp_val}"
        else:
            colonia_display = colonia_val

        # Títulos y valores en dos columnas
        campos = [
            ("Empresa Visitada:", self.datos.get('empresa_visitada', '')),
            ("Calle y No.:", self.datos.get('calle_numero', '')),
            ("Colonia o Población:", colonia_display),
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
        """Dibuja la sección de firmas sin empalmes y sin partirse entre páginas."""
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
        import os

        # Posición base para empezar a dibujar la tabla de firmas
        # Dejamos un margen inferior mínimo de 80 mm
        margen_inferior_minimo = 80 * mm
        
        # Calcular espacio requerido para la tabla de firmas
        inspectores = self.datos.get("inspectores", []) or []
        num_inspectores = len([i for i in inspectores if i])
        
        # Altura aproximada requerida (reducida para que quede más junto):
        # - Encabezados: 15mm (reducido de 25mm)
        # - Por cada inspector: 20mm (reducido de 25mm)
        altura_requerida = 15 * mm + (num_inspectores * 20 * mm)
        
        # Verificar si hay suficiente espacio en la página actual
        if self.cursor_y < margen_inferior_minimo + altura_requerida:
            # No hay suficiente espacio, crear nueva página
            self.page_num = getattr(self, "page_num", 1) + 1
            c.showPage()
            
            # Dibujar fondo en nueva página
            self.dibujar_fondo(c)
            self.dibujar_paginacion(c)
            
            # Resetear cursor para nueva página
            self.cursor_y = self.height - 100  # Empezar más arriba en la nueva página (reducido de 120)
        
        # Ahora dibujar la tabla de firmas
        x_start = 25 * mm
        
        col_izq = 50 * mm
        col_der = 50 * mm

        y_text = self.cursor_y - 8 * mm  # Reducido de 10mm
        c.setFont("Helvetica", 8)

        # -------------------------------------------------
        # Función para texto envuelto (con interlineado reducido)
        # -------------------------------------------------
        def write_wrapped(texto, x, y, max_width):
            palabras = texto.split()
            linea = ""
            for palabra in palabras:
                test = (linea + " " + palabra).strip()
                if c.stringWidth(test, "Helvetica", 8) > max_width:
                    c.drawString(x, y, linea)
                    y -= 3 * mm  # Reducido de 4mm
                    linea = palabra
                else:
                    linea = test
            if linea:
                c.drawString(x, y, linea)
                y -= 3 * mm  # Reducido de 4mm
            return y

        # Encabezados
        y_i = write_wrapped(
            "Nombre y Firma del responsable de atender la visita",
            x_start + 5 * mm, y_text, col_izq
        )

        y_d = write_wrapped(
            "Nombre y Firma del Inspector",
            x_start + col_izq + 15 * mm, y_text, col_der
        )

        # Cursor después del encabezado (reducido)
        current_y = min(y_i, y_d) - 12 * mm  # Reducido de 20mm

        # -------------------------------------------------
        # Datos de firmas
        # -------------------------------------------------
        firma_w = 50 * mm
        firma_h = 18 * mm
        inspector_x = x_start + col_izq + 15 * mm
        firma_x = inspector_x + 60 * mm

        # -------------------------------------------------
        # Obtener firma
        # -------------------------------------------------
        def obtener_firma(nombre):
            nombre_norm = nombre.strip().lower()
            for f in getattr(self, "firmas_data", []):
                if f.get("NOMBRE DE INSPECTOR", "").strip().lower() == nombre_norm:
                    ruta = f.get("IMAGEN") or f.get("FIRMA")
                    if ruta and os.path.exists(ruta):
                        return ruta
                    posible = os.path.join("Firmas", os.path.basename(ruta)) if ruta else None
                    if posible and os.path.exists(posible):
                        return posible

            nombre_arch = nombre.replace(" ", "").upper() + ".png"
            ruta = os.path.join("Firmas", nombre_arch)
            return ruta if os.path.exists(ruta) else None

        # -------------------------------------------------
        # DIBUJAR INSPECTORES (con espaciado reducido)
        # -------------------------------------------------
        for insp in inspectores:
            if not insp:
                continue

            # Nombre (con fuente ligeramente más grande para mejor legibilidad)
            c.setFont("Helvetica", 10)  # Aumentado de 9 a 10
            c.drawString(inspector_x, current_y, insp)

            # Firma debajo del nombre (espacio reducido)
            firma_y = current_y - 5 * mm - firma_h  # Reducido de 6mm a 5mm

            ruta = obtener_firma(insp)
            if ruta:
                try:
                    img = ImageReader(ruta)
                    c.drawImage(
                        img,
                        firma_x,
                        firma_y,
                        width=firma_w,
                        height=firma_h,
                        preserveAspectRatio=True,
                        mask="auto"
                    )
                except Exception:
                    # Dibujar línea si no hay imagen
                    c.line(firma_x, firma_y + 2 * mm, firma_x + firma_w, firma_y + 2 * mm)
            else:
                # Dibujar línea si no hay imagen
                c.line(firma_x, firma_y + 2 * mm, firma_x + firma_w, firma_y + 2 * mm)

            # Espacio entre inspectores (reducido significativamente)
            current_y -= 25 * mm  # Reducido de 35mm a 25mm

        # Actualizar cursor global
        self.cursor_y = current_y
        
        # Añadir pie de página si estamos en la página de firmas
        if getattr(self, 'page_num', 1) > 1:
            # Añadir texto de prohibición en la página de firmas
            c.setFont("Helvetica", 7)
            texto_prohibicion = "SE PROHIBE LA REPRODUCCIÓN TOTAL O PARCIAL DE ESTE DOCUMENTO SIN PREVIA AUTORIZACIÓN POR ESCRITO DE LA GERENCIA TÉCNICA DE VERIFICACIÓN & CONTROLIVA, S.C."
            c.drawCentredString(self.width / 2, 15 * mm, texto_prohibicion)  # Reducido de 20mm a 15mm   

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

        # Inicializar contador de páginas y cursor
        self.page_num = 1
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

    # Intentar obtener dirección/datos desde Clientes.json
    calle = datos_visita.get('direccion','')
    colonia = datos_visita.get('colonia','')
    municipio = datos_visita.get('municipio','')
    ciudad_estado = datos_visita.get('ciudad_estado','')
    numero_contrato = ''
    rfc = ''
    clientes_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Clientes.json')
    try:
        if os.path.exists(clientes_path):
            with open(clientes_path, 'r', encoding='utf-8') as cf:
                clientes = json.load(cf)
                if isinstance(clientes, list):
                    for c in clientes:
                        if str(c.get('CLIENTE','')).strip().upper() == str(datos_visita.get('cliente','')).strip().upper():
                            calle = c.get('CALLE Y NO') or c.get('CALLE','') or calle
                            colonia = c.get('COLONIA O POBLACION') or c.get('COLONIA','') or colonia
                            municipio = c.get('MUNICIPIO O ALCADIA') or c.get('MUNICIPIO','') or municipio
                            ciudad_estado = c.get('CIUDAD O ESTADO') or c.get('CIUDAD/ESTADO') or ciudad_estado
                            numero_contrato = c.get('NÚMERO_DE_CONTRATO','')
                            rfc = c.get('RFC','')
                            break
    except Exception:
        pass

    # CP: preferir valor en visita, si no, intentar extraer del final de la direccion
    cp = datos_visita.get('cp') or datos_visita.get('CP') or datos_visita.get('codigo_postal','')
    if not cp and calle:
        last = str(calle).split(',')[-1].strip()
        s = ''.join(ch for ch in last if ch.isdigit())
        if s:
            cp = s

    colonia_mostrada = colonia
    if cp:
        colonia_mostrada = f"{colonia_mostrada} {cp}" if colonia_mostrada else str(cp)

    # Preparar datos para el PDF
    datos_oficio = {
        'no_oficio': datos_visita.get('folio_acta', 'AC' + datos_visita.get('folio_visita', '0000')),
        'fecha_inspeccion': datos_visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'normas': datos_visita.get('norma', '').split(', ') if datos_visita.get('norma') else [],
        'empresa_visitada': datos_visita.get('cliente', ''),
        'calle_numero': calle,
        'colonia': colonia_mostrada,
        'municipio': municipio,
        'ciudad_estado': ciudad_estado,
        'fecha_confirmacion': datos_visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'medio_confirmacion': 'correo electrónico',
        'inspectores': inspectores,
        'observaciones': datos_visita.get('observaciones', 'Sin observaciones'),
        'num_solicitudes': datos_visita.get('num_solicitudes', 'Sin especificar'),
        'NUMERO_DE_CONTRATO': numero_contrato,
        'RFC': rfc,
        'cp': str(cp) if cp is not None else ''
    }

    return datos_oficio

# Ejemplo de uso
if __name__ == "__main__":
    # Datos de ejemplo
    datos_ejemplo = {
        'no_oficio': '2025-001',
        'fecha_inspeccion': '02/12/2025',
        'normas': ['NOM-004-SE-2021'],
        'empresa_visitada': 'ARTICULOS DEPORTIVOS DECATHLON SA DE CV',
        'calle_numero': 'Parque industrial advance II',
        'colonia': 'Capula, 09876',
        'municipio': 'Tepotzotlán',
        'ciudad_estado': 'Estado de México',
        'cp': '09876',
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
    generar_oficio_pdf(datos_ejemplo, "Plantillas PDF/Oficio_comision.pdf")


