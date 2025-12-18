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

            # Fallback: buscar en APPDATA\GeneradorDictamenes
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
        c.setFont("Helvetica", 8)
        # Código/clave documento a la derecha
        c.drawRightString(self.width - 30, self.height - 30, "PT-F-208A-00-1")
        # Contador de página (solo número actual)
        page_num = getattr(self, 'page_num', 1)
        c.drawRightString(self.width - 40, self.height - 40, f"Página {page_num}")
    
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
        """Dibuja la sección de firmas en el orden solicitado con mejor espaciado

        Orden:
        - Nombre y Firma del cliente o responsable de atender la visita
        - Nombre y Firma (Testigo 1)
        - Nombre y Firma del Inspector (uno o varios)

        Esta versión sólo ajusta el orden y la disposición de nombres/firma;
        no modifica el resto de campos del documento.
        """
        x = 25 * mm
        ancho_total = 165 * mm

        c.setFont("Helvetica", 9)
        y = self.cursor_y - 15

        # Helper para asegurarse de que hay espacio suficiente en la página;
        # si no, crea una nueva página, dibuja el fondo y la paginación,
        # incrementando el contador de páginas.
        def ensure_space(pos_y, min_space=80):
            if pos_y < min_space:
                try:
                    self.page_num = getattr(self, 'page_num', 1) + 1
                except Exception:
                    self.page_num = 2
                c.showPage()
                try:
                    self.dibujar_fondo(c)
                except Exception:
                    pass
                try:
                    self.dibujar_paginacion(c)
                except Exception:
                    pass
                # resetear un cursor en la nueva página (margen superior)
                return self.height - 60
            return pos_y

        firma_ancho = 55 * mm  # Aumentar ancho
        firma_alto = 20  # Aumentar alto

        # Helper para dibujar nombre + firma (imagen o línea)
        def dibujar_nombre_y_firma(label, nombre, pos_y):
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x, pos_y, label)
            pos_y -= 12
            c.setFont("Helvetica", 9)
            nombre_text = nombre or ''
            # mostrar nombre (truncado si muy largo)
            if len(nombre_text) > 60:
                nombre_text = nombre_text[:57] + '...'
            c.drawString(x, pos_y, nombre_text)

            # intentar firma (buscar en Firmas.json)
            firma_path = None
            if nombre_text:
                firma_path = self.obtener_firma_inspector(nombre_text)

            if firma_path and os.path.exists(firma_path):
                try:
                    img = ImageReader(firma_path)
                    c.drawImage(img, x + 90 * mm, pos_y - 10, width=firma_ancho, height=firma_alto, preserveAspectRatio=True, mask='auto')
                except Exception as e:
                    print(f"⚠️ Error cargando firma {firma_path}: {e}")
                    c.line(x + 90 * mm, pos_y, x + 90 * mm + firma_ancho, pos_y)
            else:
                # línea de firma
                c.line(x + 90 * mm, pos_y, x + 90 * mm + firma_ancho, pos_y)

            return pos_y - 30  # Aumentar espaciamiento

        # 1) Cliente / responsable
        cliente_nombre = self.datos.get('empresa_visitada') or self.datos.get('cliente') or ''
        y = ensure_space(y)
        y = dibujar_nombre_y_firma('Nombre y Firma del cliente o responsable de atender la visita', cliente_nombre, y)

        # 2) Testigo 1
        testigo1 = self.datos.get('testigo1') or self.datos.get('testigo_1') or ''
        y = ensure_space(y)
        y = dibujar_nombre_y_firma('Nombre y Firma (Testigo 1)', testigo1, y)

        # 3) Inspector(es)
        inspectores = self.datos.get('inspectores', []) or []
        if not inspectores:
            nd = self.datos.get('NOMBRE_DE_INSPECTOR')
            if nd:
                inspectores = [s.strip() for s in nd.split(',') if s.strip()]

        # Si hay varios inspectores, listarlos uno por uno
        if inspectores:
            for insp in inspectores:
                y = ensure_space(y)
                y = dibujar_nombre_y_firma('Nombre y Firma del Inspector', insp, y)
        else:
            # Si no hay inspectores, dejar un espacio vacío para firma
            y = ensure_space(y)
            y = dibujar_nombre_y_firma('Nombre y Firma del Inspector', '', y)

        # Espacio para siguiente sección
        y -= 6

        # NOTAS Y OBSERVACIONES (mantener comportamiento previo)
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(x + ancho_total / 2, y, "NOTAS Y OBSERVACIONES:")

        y -= 20

        # Observaciones Cliente
        c.setFont("Helvetica", 9)
        c.drawString(x, y, "Observaciones (Cliente):")
        y -= 10

        for _ in range(3):
            c.line(x, y, x + ancho_total - 10, y)
            y -= 15

        y -= 10

        # Observaciones Inspector
        c.drawString(x, y, "Observaciones (Inspector):")
        y -= 10

        for _ in range(3):
            c.line(x, y, x + ancho_total - 10, y)
            y -= 15

        y -= 20

        # ACTA Y C.P. (mantener)
        acta = self.datos.get("acta", "C.P.12345")
        cp = self.datos.get("cp", "CP07890")

        c.drawString(x, y, f"Acta: {acta}    C.P.: {cp}")

        # Actualizar cursor general
        self.cursor_y = y - 25

    def _dibujar_tabla_productos_canvas(self, c, productos):
        """Dibuja una tabla simple en canvas con los campos solicitados.
        Campos mostrados: SOLICITUD, PEDIMENTO, FACTURA, CODIGO, CANTIDAD, EVALUACIÓN
        """
        # Coordenadas y medidas
        left = 20 * mm
        top = self.height - 90
        row_h = 12
        # Columnas más compactas para que se vean más juntas
        col_widths = [40 * mm, 40 * mm, 40 * mm, 40 * mm, 20 * mm, 20 * mm]

        headers = ["No. De Solicitud", "No. De Pedimento", "Factura", "Código", "Piezas", "Eval."]

        # Dibujar encabezados
        c.setFont("Helvetica-Bold", 7)
        x = left
        for i, h in enumerate(headers):
            c.drawString(x + 2, top, h)
            x += col_widths[i]

        y = top - row_h
        c.setFont("Helvetica", 7)

        # Iterar productos y pintar filas (paginar si necesario)
        for idx, prod in enumerate(productos):
            if y < 50:  # nueva página
                # incrementar contador de páginas y crear nueva
                try:
                    self.page_num = getattr(self, 'page_num', 1) + 1
                except Exception:
                    self.page_num = 2
                c.showPage()
                # dibujar fondo y paginación en nueva página
                try:
                    self.dibujar_fondo(c)
                except Exception:
                    pass
                try:
                    self.dibujar_paginacion(c)
                except Exception:
                    pass

                y = self.height - 90
                c.setFont("Helvetica-Bold", 7)
                x = left
                for i, h in enumerate(headers):
                    c.drawString(x + 2, y, h)
                    x += col_widths[i]
                y -= row_h
                c.setFont("Helvetica", 7)

            x = left
            # Obtener valores con fallback
            solicitud = str(prod.get('SOLICITUD', prod.get('SOLICITUD', '')))
            pedimento = str(prod.get('PEDIMENTO', prod.get('PEDIMENTO', '')))
            factura = str(prod.get('FACTURA', prod.get('FACTURA', '')))
            codigo = str(prod.get('CODIGO', prod.get('CODIGO', '')))
            cantidad = str(prod.get('CANTIDAD', prod.get('CANTIDAD', '')))
            evaluacion = prod.get('EVALUACION', prod.get('EVALUACIÓN', 'C') ) or 'C'

            values = [solicitud, pedimento, factura, codigo, cantidad, evaluacion]
            for i, val in enumerate(values):
                # Truncar si es muy largo
                txt = str(val)
                # Ajustar font y ancho máximo por columna
                max_chars = int(col_widths[i] / 3.8)
                if len(txt) > max_chars:
                    txt = txt[:max_chars-3] + '...'
                c.drawString(x + 2, y, txt)
                x += col_widths[i]

            y -= row_h

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

        # inicializar contador de páginas
        self.page_num = 1

        # Resetear cursor al inicio
        self.cursor_y = self.height - 40

        # Dibujar fondo y paginación en la primera página
        try:
            self.dibujar_fondo(c)
        except Exception:
            pass
        try:
            self.dibujar_paginacion(c)
        except Exception:
            pass

        # Dibujar encabezado
        self.dibujar_encabezado(c)

        # Dibujar tabla superior
        self.dibujar_tabla_superior(c)

        # Dibujar datos empresa
        self.dibujar_datos_empresa(c)

        # Dibujar tabla de firmas
        self.dibujar_tabla_firmas(c)

        # Si hay tabla de productos en los datos, añadir una segunda hoja
        productos = self.datos.get('tabla_productos', []) or []
        if productos:
            # terminar la primera página y crear la siguiente
            # aumentar contador
            try:
                self.page_num = getattr(self, 'page_num', 1) + 1
            except Exception:
                self.page_num = 2
            c.showPage()
            # Dibujar fondo y paginación en la segunda hoja
            try:
                self.dibujar_fondo(c)
            except Exception:
                pass
            try:
                self.dibujar_paginacion(c)
            except Exception:
                pass

            # Dibujar encabezado simple en la segunda hoja
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(self.width / 2, self.height - 40, "LISTA DE PRODUCTOS - DETALLE")
            # Dibujar la tabla de productos
            self._dibujar_tabla_productos_canvas(c, productos)

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
    
    # Intentar completar dirección y datos desde Clientes.json
    calle = datos_visita.get('direccion', '')
    colonia = datos_visita.get('colonia', '')
    municipio = datos_visita.get('municipio', '')
    ciudad_estado = datos_visita.get('ciudad_estado', '')
    numero_contrato = ''
    rfc = ''
    curp = ''
    clientes_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Clientes.json')
    try:
        if os.path.exists(clientes_path):
            with open(clientes_path, 'r', encoding='utf-8') as cf:
                clientes = json.load(cf)
                # Clientes.json puede ser lista
                if isinstance(clientes, list):
                    for c in clientes:
                        # comparar por nombre de cliente (case-insensitive)
                        if str(c.get('CLIENTE','')).strip().upper() == str(datos_visita.get('cliente','')).strip().upper():
                            calle = c.get('CALLE Y NO') or c.get('CALLE','') or calle
                            colonia = c.get('COLONIA O POBLACION') or c.get('COLONIA','') or colonia
                            municipio = c.get('MUNICIPIO O ALCADIA') or c.get('MUNICIPIO','') or municipio
                            ciudad_estado = c.get('CIUDAD O ESTADO') or c.get('CIUDAD/ESTADO') or ciudad_estado
                            numero_contrato = c.get('NÚMERO_DE_CONTRATO','')
                            rfc = c.get('RFC','')
                            curp = c.get('CURP','') or curp
                            break
    except Exception:
        pass

    # Preparar datos para el PDF
    datos_acta = {
        'fecha_inspeccion': datos_visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'normas': datos_visita.get('norma', '').split(', ') if datos_visita.get('norma') else [],
        'empresa_visitada': datos_visita.get('cliente', ''),
        'calle_numero': calle,
        'colonia': colonia,
        'municipio': municipio,
        'ciudad_estado': ciudad_estado,
        'fecha_confirmacion': datos_visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'medio_confirmacion': 'correo electrónico',
        'inspectores': inspectores,
        'NOMBRE_DE_INSPECTOR': (datos_visita.get('supervisores_tabla') or datos_visita.get('nfirma1') or '').strip(),
        'observaciones': datos_visita.get('observaciones', 'Sin observaciones'),
        'NUMERO_DE_CONTRATO': numero_contrato,
        'RFC': rfc,
        'CURP': curp
        
    }
    
    return datos_acta


def generar_acta_desde_visita(folio_visita=None, ruta_salida=None):
    """Genera un acta a partir de la información en data/historial_visitas.json y
    data/tabla_de_relacion.json. Si `folio_visita` es None toma la última visita.
    """
    base_dir = os.path.join(os.path.dirname(__file__), '..')
    data_dir = os.path.join(base_dir, 'data')
    historial_path = os.path.join(data_dir, 'historial_visitas.json')
    # Preferir backups recientes: si existen backups en tabla_relacion_backups, usar el más reciente
    tabla_path = os.path.join(data_dir, 'tabla_de_relacion.json')
    backups_dir = os.path.join(data_dir, 'tabla_relacion_backups')
    if os.path.exists(backups_dir):
        try:
            backup_files = [os.path.join(backups_dir, f) for f in os.listdir(backups_dir) if f.lower().endswith('.json')]
            if backup_files:
                tabla_path = max(backup_files, key=os.path.getmtime)
                print(f"📁 Usando backup de tabla de relación para acta: {tabla_path}")
        except Exception:
            pass

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

    # Determinar lista de folios asociados a la visita
    folios_list = []
    # 1) intentar cargar archivo en data/folios_visitas/folios_<numeric>.json
    folio_num = ''.join([c for c in visita.get('folio_visita','') if c.isdigit()])
    folios_file = os.path.join(data_dir, 'folios_visitas', f'folios_{folio_num}.json')
    if os.path.exists(folios_file):
        try:
            with open(folios_file, 'r', encoding='utf-8') as ff:
                data = json.load(ff)
                # esperar lista de folios en data
                if isinstance(data, list):
                    folios_list = [int(x) for x in data if str(x).isdigit()]
                elif isinstance(data, dict) and 'folios' in data:
                    folios_list = [int(x) for x in data.get('folios', []) if str(x).isdigit()]
        except Exception:
            folios_list = []

    # 2) fallback: parsear visita['folios_utilizados'] si existe (ej: '046294 - 046302')
    if not folios_list:
        fu = visita.get('folios_utilizados') or visita.get('folios_utilizados', '')
        if fu and isinstance(fu, str):
            if '-' in fu:
                parts = [p.strip() for p in fu.split('-')]
                try:
                    start = int(parts[0])
                    end = int(parts[1]) if len(parts) > 1 else start
                    folios_list = list(range(start, end+1))
                except Exception:
                    folios_list = []
            elif ',' in fu:
                vals = [p.strip() for p in fu.split(',')]
                for v in vals:
                    if v.isdigit():
                        folios_list.append(int(v))

    # Cargar tabla_de_relacion (o backup seleccionado) y filtrar registros por folio
    productos = []
    fecha_verificacion = None
    if os.path.exists(tabla_path):
        try:
            with open(tabla_path, 'r', encoding='utf-8') as tf:
                tabla = json.load(tf)
                # tabla puede ser lista de dicts
                for rec in tabla:
                    fol = rec.get('FOLIO')
                    try:
                        fol_int = int(fol) if fol is not None and str(fol).isdigit() else None
                    except Exception:
                        fol_int = None
                    if folios_list and fol_int in folios_list:
                        productos.append(rec)
                        if not fecha_verificacion and rec.get('FECHA DE VERIFICACION'):
                            fecha_verificacion = rec.get('FECHA DE VERIFICACION')
                # si no encontramos por folios, intentar usar primer registro si existe
                if not productos and isinstance(tabla, list) and tabla:
                    # intentar extraer fecha de verificacion del primer registro
                    first = tabla[0]
                    if first.get('FECHA DE VERIFICACION') and not fecha_verificacion:
                        fecha_verificacion = first.get('FECHA DE VERIFICACION')
        except Exception:
            productos = []

    # Preparar datos para el acta
    normas = visita.get('norma', '')
    normas_list = [n.strip() for n in normas.split(',')] if normas else []

    # Formatear fecha_verificacion a dd/mm/YYYY si viene en formato ISO
    fecha_formateada = None
    if fecha_verificacion:
        try:
            # soportar formatos como YYYY-MM-DD o dd/mm/YYYY
            if '-' in fecha_verificacion:
                dt = datetime.strptime(fecha_verificacion[:10], '%Y-%m-%d')
            else:
                dt = datetime.strptime(fecha_verificacion[:10], '%d/%m/%Y')
            fecha_formateada = dt.strftime('%d/%m/%Y')
        except Exception:
            fecha_formateada = fecha_verificacion

    datos_acta = {
        # Fecha de inicio/termino extraída de tabla_de_relacion (FECHA DE VERIFICACION)
        'fecha_inicio': fecha_formateada or visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'hora_inicio': '09:00',
        'fecha_termino': fecha_formateada or visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'hora_termino': '18:00',
        'normas': normas_list,
        'empresa_visitada': visita.get('cliente', ''),
        'calle_numero': visita.get('direccion', ''),
        'colonia': visita.get('colonia', ''),
        'municipio': visita.get('municipio', ''),
        'ciudad_estado': visita.get('ciudad_estado', ''),
        'inspectores': [s.strip() for s in (visita.get('supervisores_tabla') or visita.get('nfirma1') or '').split(',') if s.strip()],
        'NOMBRE_DE_INSPECTOR': (visita.get('supervisores_tabla') or visita.get('nfirma1') or '').strip(),
        'observaciones': visita.get('observaciones', ''),
        'acta': visita.get('folio_acta', ''),
        'cp': visita.get('folio_visita', ''),
        'tabla_productos': productos
    }

    # Determinar ruta de salida
    if not ruta_salida:
        fol = visita.get('folio_visita', 'acta')
        ruta_salida = os.path.join(os.path.dirname(__file__), '..', f'Acta_{fol}.pdf')

    # Generar PDF
    generar_acta_pdf(datos_acta, ruta_salida)
    return ruta_salida

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


