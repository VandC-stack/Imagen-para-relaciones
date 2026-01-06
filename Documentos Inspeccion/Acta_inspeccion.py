# -- Acta de inspección -- #
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import os
import json
import sys
from pathlib import Path

# Simple debug logger for acta generation; writes to data/acta_debug.log next to exe when possible
def _log_acta(msg: str):
    try:
        base = None
        if getattr(sys, 'frozen', False):
            base = Path(sys.executable).parent
        else:
            base = Path(__file__).resolve().parent
        data_dir = base / 'data'
        data_dir.mkdir(parents=True, exist_ok=True)
        log_path = data_dir / 'acta_debug.log'
        with open(log_path, 'a', encoding='utf-8') as lf:
            lf.write(f"[{datetime.now().isoformat()}] {msg}\n")
    except Exception:
        # Best-effort logging only
        pass
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
        # Reducir espacio después del encabezado para compactar el diseño
        self.cursor_y -= 45

    def dibujar_tabla_superior(self, c):
        """Tabla superior de 4 columnas para ACTA DE INSPECCIÓN (sin bordes)"""

        x_start = 20 * mm
        # Usar filas más compactas
        row_height = 9

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

        # Espacio final reducido
        self.cursor_y -= 5

    def dibujar_datos_empresa(self, c):
        """Dibuja los datos de la empresa visitada sin bordes"""
        x_start = 20 * mm
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_start, self.cursor_y, "Datos del lugar donde se realiza la Inspección de Información Comercial:")
        # Espacio reducido antes de la lista de datos
        self.cursor_y -= 18
        
        # Preparar display de colonia + C.P. si existe
        colonia_val = (self.datos.get('colonia') or '')
        cp_val = str(self.datos.get('cp') or '')

        # Normalizar y formatear la colonia para mostrar siempre 'C.P.: <n>'
        def _clean_colonia_and_format_cp(colonia_text, cp_text):
            try:
                import re
                if not cp_text:
                    return colonia_text or ''
                digits = re.sub(r'\D', '', str(cp_text))
                if not digits:
                    return colonia_text or ''

                # Si la colonia ya contiene la etiqueta C.P. (en cualquier variante), dejarla
                if re.search(r'c\.?p\.?', colonia_text, flags=re.IGNORECASE):
                    # Normalizar a 'C.P.: <cp>' si es necesario
                    if digits and not re.search(re.escape(digits), colonia_text):
                        return (colonia_text.strip() + '  C.P.: ' + cp_text).strip()
                    return colonia_text

                # Si la colonia termina con el número de CP, eliminarlo y agregar la etiqueta C.P.
                pattern = r'[\s,;:\-]*' + re.escape(digits) + r'$'
                nueva = re.sub(pattern, '', colonia_text).strip()
                if nueva:
                    return f"{nueva}  C.P.: {cp_text}"
                else:
                    return f"C.P.: {cp_text}"
            except Exception:
                return colonia_text or ''

        colonia_display = _clean_colonia_and_format_cp(colonia_val, cp_val)

        # Títulos y valores en dos columnas
        campos = [
            ("Empresa Visitada:", self.datos.get('empresa_visitada', '')),
            ("Calle y No.:", self.datos.get('calle_numero', '')),
            ("Colonia o Población:", colonia_display),
            ("Municipio o Alcaldía:", self.datos.get('municipio', '')),
            ("Ciudad o Estado:", self.datos.get('ciudad_estado', ''))
        ]
        
        for titulo, valor in campos:
            # Título en negrita (pequeño)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x_start, self.cursor_y, titulo)

            # Valor: usar tamaño de letra ligeramente más pequeño y mostrar exactamente
            c.setFont("Helvetica", 9)
            c.drawString(x_start + 48 * mm, self.cursor_y, str(valor))
            # Espacio reducido entre líneas de datos
            self.cursor_y -= 8

        # Espacio reducido después de la sección
        self.cursor_y -= 10
    
    def dibujar_tabla_firmas(self, c):
        """Dibuja la sección de firmas en el orden solicitado con mejor espaciado

        Orden:
        - Nombre y Firma del cliente o responsable de atender la visita
        - Nombre y Firma (Testigo 1)
        - Nombre y Firma del Inspector (uno o varios)

        Esta versión sólo ajusta el orden y la disposición de nombres/firma;
        no modifica el resto de campos del documento.
        """
        x = 20 * mm
        ancho_total = 155 * mm

        c.setFont("Helvetica", 9)
        # Posición inicial para firmas más compacta
        y = self.cursor_y - 6

        # Ajustar tamaño de las firmas para legibilidad y ajuste
        firma_ancho = 55 * mm
        firma_alto = 20 * mm

        # Helper para asegurarse de que hay espacio suficiente en la página;
        # si no, crea una nueva página, dibuja el fondo y la paginación,
        # incrementando el contador de páginas. Calcula espacio mínimo en
        # base al alto de la firma y líneas de texto para evitar recortes.
        def ensure_space(pos_y, min_space=None):
            # Calcular un min_space razonable si no se pasó
            default_needed = int(12 + firma_alto + (4 * mm) + 8)
            if min_space is None:
                min_space = default_needed

            # Si no hay espacio suficiente, crear nueva página y volver a dibujar
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
                # Redibujar encabezado y tabla superior para mantener formato
                try:
                    self.cursor_y = self.height - 40
                    self.dibujar_encabezado(c)
                    self.dibujar_tabla_superior(c)
                except Exception:
                    pass
                # Devolver cursor justo debajo de la tabla superior
                return self.cursor_y - 8
            return pos_y

        # Helper para dibujar nombre + firma (imagen o línea)
        def dibujar_nombre_y_firma(label, nombre, pos_y):
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x, pos_y, label)
            pos_y -= 12
            c.setFont("Helvetica", 9)
            # usar el nombre completo para búsqueda de firma, pero truncar
            # solo la versión que mostramos en el PDF para evitar fallos
            full_name = nombre or ''
            display_name = full_name
            if len(display_name) > 60:
                display_name = display_name[:57] + '...'
            c.drawString(x, pos_y, display_name)

            # intentar firma (buscar en Firmas.json) usando el nombre completo
            firma_path = None
            if full_name:
                firma_path = self.obtener_firma_inspector(full_name)

            if firma_path and os.path.exists(firma_path):
                try:
                    img = ImageReader(firma_path)
                    c.drawImage(img, x + 85 * mm, pos_y - (firma_alto / 2), width=firma_ancho, height=firma_alto, preserveAspectRatio=True, mask='auto')
                except Exception as e:
                    print(f"⚠️ Error cargando firma {firma_path}: {e}")
                    c.line(x + 90 * mm, pos_y, x + 90 * mm + firma_ancho, pos_y)
            else:
                # línea de firma
                c.line(x + 85 * mm, pos_y, x + 85 * mm + firma_ancho, pos_y)

            # Retornar nueva posición dejando espacio suficiente según el alto de la firma
            separation = (firma_alto + 4 * mm)
            return pos_y - separation

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

        # Espacio para siguiente sección reducido
        y -= 4

        # NOTAS Y OBSERVACIONES (mantener comportamiento previo)
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(x + ancho_total / 2, y, "NOTAS Y OBSERVACIONES:")

        y -= 14

        # Observaciones Cliente
        c.setFont("Helvetica", 9)
        c.drawString(x, y, "Observaciones (Cliente):")
        y -= 10

        for _ in range(3):
            c.line(x, y, x + ancho_total - 10, y)
            y -= 12

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
        # Mostrar C.P. del sistema (folio de la visita) en la parte inferior.
        # Este valor es distinto al código postal que se muestra junto a la colonia.
        folio_visita = self.datos.get('folio_visita') or self.datos.get('folio') or ''
        sys_cp = ''.join(ch for ch in str(folio_visita) if ch.isdigit())
        # Si no hay folio de visita, conservar el CP postal como fallback
        if not sys_cp:
            sys_cp = str(self.datos.get('cp', ''))

        c.drawString(x, y, f"Acta: {acta}    C.P.: {sys_cp}")

        # Actualizar cursor general
        self.cursor_y = y - 25

    def _dibujar_tabla_productos_canvas(self, c, productos):
        """Dibuja una tabla simple en canvas con los campos solicitados.
        Campos mostrados: SOLICITUD, PEDIMENTO, FACTURA, CODIGO, PIEZAS, EVALUACIÓN
        """
        # Coordenadas y medidas
        left = 20 * mm
        top = self.height - 90
        row_h = 12
        # Columnas: ajustar anchos para que todas las columnas quepan en página
        # Margen izquierdo/desplazamiento aproximado se considera al dibujar
        # Totales aproximados en mm deben ser menores al ancho de página (~216mm)
        col_widths = [28 * mm, 38 * mm, 50 * mm, 50 * mm, 12 * mm, 12 * mm]

        headers = ["No. Solicitud", "No. De Pedimento", "Factura", "Código", "Piezas", "Eval."]

        # Dibujar encabezados
        c.setFont("Helvetica-Bold", 7)
        x = left
        for i, h in enumerate(headers):
            c.drawString(x + 2, top, h)
            x += col_widths[i]

        y = top - row_h
        c.setFont("Helvetica", 7)

        import re

        def _norm_key(k):
            return re.sub(r'[^0-9a-zA-Záéíóúüñ]', '', str(k).lower())

        # Iterar productos y pintar filas (paginar si necesario)
        for idx, prod in enumerate(productos):
            if y < 50:  # nueva página
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

                y = self.height - 90
                c.setFont("Helvetica-Bold", 7)
                x = left
                for i, h in enumerate(headers):
                    c.drawString(x + 2, y, h)
                    x += col_widths[i]
                y -= row_h
                c.setFont("Helvetica", 7)

            x = left
            # Normalizar claves para búsqueda flexible (case-insensitive y sin signos)
            norm = {}
            if isinstance(prod, dict):
                for k, v in prod.items():
                    norm[_norm_key(k)] = v
            else:
                # Si el registro no es dict, representarlo como string en la columna código
                norm[_norm_key('codigo')] = prod

            def find_value(candidates, default=''):
                for cand in candidates:
                    nk = _norm_key(cand)
                    if nk in norm and norm[nk] not in (None, ''):
                        return norm[nk]
                return default

            solicitud = str(find_value(['SOLICITUD', 'no solicitud', 'no.solicitud', 'no. de solicitud', 'no_solicitud', 'SOLICITUD_NO'], ''))
            pedimento = str(find_value(['PEDIMENTO', 'no pedimento', 'no.pedimento', 'no. de pedimento', 'nopedimento', 'num pedimento', 'numero pedimento'], ''))
            factura = str(find_value(['FACTURA', 'factura', 'FACTURAS', 'NO FACTURA', 'NUM FACTURA', 'NFACTURA', 'n° factura'], ''))
            codigo = str(find_value(['CODIGO', 'CÓDIGO', 'codigo', 'SKU', 'ITEM', 'REFERENCIA'], ''))
            # Obtener piezas desde el campo CONTENIDO (si existe). Mantener CANTIDAD internamente si se necesita.
            piezas = str(find_value(['CONTENIDO', 'DESCRIPCION', 'DESCRIPCIÓN', 'CONTENIDO MERCANCIA', 'CONTENIDO_PRODUCTO', 'DESCRIPCION_PRODUCTO'], ''))
            cantidad = str(find_value(['CANTIDAD', 'PIEZAS', 'piezas', 'cantidad', 'UNIDADES', 'CANT'], ''))
            evaluacion = str(find_value(['EVALUACION', 'EVALUACIÓN', 'Eval', 'EVAL'], '')) or 'C'

            # Mostrar Piezas (valor de CONTENIDO) y no la columna 'Contenido'
            values = [solicitud, pedimento, factura, codigo, piezas, evaluacion]
            for i, val in enumerate(values):
                txt = '' if val is None else str(val)
                # Ajustar font y ancho máximo por columna
                try:
                    max_chars = int(col_widths[i] / 3.8)
                except Exception:
                    max_chars = 20
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
    
    # Intentar completar dirección y datos desde la visita o Clientes.json
    direccion_compuesta = datos_visita.get('direccion', '')
    calle_numero = datos_visita.get('calle_numero', '')
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
                            if not calle_numero:
                                calle_numero = c.get('CALLE Y NO') or c.get('CALLE','') or calle_numero
                            colonia = colonia or c.get('COLONIA O POBLACION') or c.get('COLONIA','') or colonia
                            municipio = municipio or c.get('MUNICIPIO O ALCADIA') or c.get('MUNICIPIO','') or municipio
                            ciudad_estado = ciudad_estado or c.get('CIUDAD O ESTADO') or c.get('CIUDAD/ESTADO') or ciudad_estado
                            numero_contrato = numero_contrato or c.get('NÚMERO_DE_CONTRATO','')
                            rfc = rfc or c.get('RFC','')
                            curp = curp or c.get('CURP','')
                            break
    except Exception:
        pass

    # Si no tenemos `calle_numero` pero sí `direccion_compuesta`, intentar separar
    if not calle_numero and direccion_compuesta:
        parts = [p.strip() for p in direccion_compuesta.split(',') if p.strip()]
        if parts:
            calle_numero = parts[0]
            if not colonia and len(parts) > 1:
                colonia = parts[1]
            if not municipio and len(parts) > 2:
                municipio = parts[2]
            if not ciudad_estado and len(parts) > 3:
                ciudad_estado = parts[3]

    # CP: preferir valor en visita, si no, intentar extraer del final de la direccion compuesta
    cp = datos_visita.get('cp') or datos_visita.get('CP') or datos_visita.get('codigo_postal','')
    if not cp and direccion_compuesta:
        last = direccion_compuesta.split(',')[-1].strip()
        s = ''.join(ch for ch in last if ch.isdigit())
        if s:
            cp = s

    colonia_mostrada = colonia
    if cp:
        colonia_mostrada = f"{colonia_mostrada} {cp}" if colonia_mostrada else cp

    # Preparar datos para el PDF
    datos_acta = {
        'fecha_inspeccion': datos_visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'normas': datos_visita.get('norma', '').split(', ') if datos_visita.get('norma') else [],
        'empresa_visitada': datos_visita.get('cliente', ''),
        'calle_numero': calle_numero,
        'colonia': colonia_mostrada,
        'municipio': municipio,
        'ciudad_estado': ciudad_estado,
        'fecha_confirmacion': datos_visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'medio_confirmacion': 'correo electrónico',
        'inspectores': inspectores,
        'NOMBRE_DE_INSPECTOR': (datos_visita.get('supervisores_tabla') or datos_visita.get('nfirma1') or '').strip(),
        'observaciones': datos_visita.get('observaciones', 'Sin observaciones'),
        'NUMERO_DE_CONTRATO': numero_contrato,
        'RFC': rfc,
        'CURP': curp,
        'cp': str(cp) if cp is not None else ''
        
    }
    
    return datos_acta

def generar_acta_desde_visita(folio_visita=None, ruta_salida=None):
    """Genera un acta a partir de la información en data/historial_visitas.json y
    data/tabla_de_relacion.json. Si `folio_visita` es None toma la última visita.
    """
    # Resolver data_dir de forma robusta: preferir carpeta junto al exe,
    # luego ruta bundle (_MEIPASS), luego ruta relativa al paquete y cwd.
    candidates = []
    try:
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            candidates.append(os.path.join(exe_dir, 'data'))
    except Exception:
        pass

    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        candidates.append(os.path.join(meipass, 'data'))

    # data relative to the source tree
    candidates.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data')))
    # current working directory
    candidates.append(os.path.join(os.getcwd(), 'data'))

    # APPDATA fallback
    try:
        appdata_dir = os.path.join(os.environ.get('APPDATA', ''), 'GeneradorDictamenes')
        candidates.append(os.path.join(appdata_dir, 'data'))
    except Exception:
        pass

    data_dir = None
    for c in candidates:
        _log_acta(f"Checking candidate data dir: {c}")
        if c and os.path.exists(c):
            data_dir = c
            _log_acta(f"Selected data_dir: {data_dir}")
            break

    if data_dir is None:
        # No candidate existed; default to first candidate or './data'
        data_dir = candidates[0] if candidates else os.path.abspath('data')
        _log_acta(f"No existing data dir found; defaulting to: {data_dir}")

    historial_path = os.path.join(data_dir, 'historial_visitas.json')
    _log_acta(f"Historial path resolved to: {historial_path} (exists={os.path.exists(historial_path)})")
    # Default tabla_path; we may override with a per-visit backup later
    tabla_path = os.path.join(data_dir, 'tabla_de_relacion.json')
    backups_dir = os.path.join(data_dir, 'tabla_relacion_backups')

    # If historial doesn't exist at resolved path, try a few fallbacks
    if not os.path.exists(historial_path):
        alt_candidates = [
            os.path.join(os.getcwd(), 'data', 'historial_visitas.json'),
            os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data', 'historial_visitas.json')),
            os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'historial_visitas.json'))
        ]
        found_alt = None
        for a in alt_candidates:
            _log_acta(f"Trying alt historial: {a} (exists={os.path.exists(a)})")
            if os.path.exists(a):
                found_alt = a
                break
        if found_alt:
            historial_path = found_alt
            _log_acta(f"Using alternative historial_path: {historial_path}")
        else:
            _log_acta(f"Historial not found in any candidate. Last checked: {historial_path}")
            raise FileNotFoundError(f"No se encontró {historial_path} o alternativas")

    # Load historial with diagnostics
    try:
        with open(historial_path, 'r', encoding='utf-8') as f:
            historial = json.load(f)
    except Exception as e:
        _log_acta(f"Error leyendo historial JSON: {e}; path={historial_path}")
        raise

    # soportar formato {'visitas': [...]} o lista simple
    visitas = historial.get('visitas', []) if isinstance(historial, dict) else historial
    _log_acta(f"Loaded historial: type={type(historial).__name__}, visitas_count={len(visitas) if hasattr(visitas, '__len__') else 'unknown'}")
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
                # Formatos soportados:
                # - lista simple de números/strings
                # - dict {'folios': [...]}
                # - lista de dicts (cada dict con clave 'FOLIOS' o 'FOLIO')
                if isinstance(data, list):
                    # lista simple?
                    if data and all(not isinstance(x, dict) for x in data):
                        folios_list = [int(x) for x in data if str(x).strip().isdigit()]
                    else:
                        # lista de registros -> extraer campo FOLIOS/FOLIO
                        extracted = []
                        for rec in data:
                            if not isinstance(rec, dict):
                                continue
                            val = None
                            for key in ('FOLIOS', 'FOLIO', 'folios', 'folio'):
                                if key in rec and rec.get(key) is not None:
                                    val = rec.get(key)
                                    break
                            if val is None:
                                continue
                            # val puede ser '000867' u objeto. Extraer dígitos
                            s = str(val)
                            digits = ''.join([c for c in s if c.isdigit()])
                            if digits:
                                try:
                                    extracted.append(int(digits))
                                except Exception:
                                    pass
                        folios_list = extracted
                elif isinstance(data, dict) and 'folios' in data:
                    folios_list = [int(x) for x in data.get('folios', []) if str(x).strip().isdigit()]
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
    # If backups exist, prefer a backup matching the visit's folio when possible
    try:
        if os.path.exists(backups_dir):
            backup_files = [f for f in os.listdir(backups_dir) if f.lower().endswith('.json')]
            # try to find backups that include the folio_num in their filename
            matching = [os.path.join(backups_dir, f) for f in backup_files if folio_num and folio_num in f]
            if matching:
                # pick the most recent matching backup
                tabla_path = max(matching, key=os.path.getmtime)
                _log_acta(f"Using per-visit backup for tabla_de_relacion: {tabla_path}")
            else:
                # fallback to most recent backup overall if present
                if backup_files:
                    allpaths = [os.path.join(backups_dir, f) for f in backup_files]
                    tabla_path = max(allpaths, key=os.path.getmtime)
                    _log_acta(f"Using latest backup (no per-visit match) for tabla_de_relacion: {tabla_path}")
    except Exception as e:
        _log_acta(f"Error selecting backup: {e}")

    if os.path.exists(tabla_path):
        try:
            with open(tabla_path, 'r', encoding='utf-8') as tf:
                tabla = json.load(tf)

                # Normalizar `tabla` a una lista de registros llamada `records`.
                records = []
                if isinstance(tabla, list):
                    records = tabla
                elif isinstance(tabla, dict):
                    # Buscar el primer valor que sea una lista (común en backups)
                    for v in tabla.values():
                        if isinstance(v, list):
                            records = v
                            break
                    # Intentar claves comunes
                    if not records:
                        for key in ("registros", "tabla", "data", "rows", "records"):
                            if key in tabla and isinstance(tabla[key], list):
                                records = tabla[key]
                                break

                _log_acta(f"Loaded tabla_path={tabla_path}; tabla_type={type(tabla).__name__}; records={len(records)}")

                # tabla puede ser lista de dicts en `records`
                for rec in records:
                    if not isinstance(rec, dict):
                        continue
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
                if not productos and records:
                    first = records[0]
                    if isinstance(first, dict) and first.get('FECHA DE VERIFICACION') and not fecha_verificacion:
                        fecha_verificacion = first.get('FECHA DE VERIFICACION')
        except Exception as e:
            _log_acta(f"Error leyendo tabla_de_relacion: {e}")
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

    # CP: preferir valor en visita, si no, intentar extraer del final de la direccion
    cp = visita.get('cp') or visita.get('CP') or visita.get('codigo_postal','')
    if not cp and visita.get('direccion'):
        last = visita.get('direccion','').split(',')[-1].strip()
        s = ''.join(ch for ch in last if ch.isdigit())
        if s:
            cp = s

    colonia_mostrada = visita.get('colonia','')
    if cp:
        colonia_mostrada = f"{colonia_mostrada} {cp}" if colonia_mostrada else str(cp)

    datos_acta = {
        # Fecha de inicio/termino extraída de tabla_de_relacion (FECHA DE VERIFICACION)
        'fecha_inicio': fecha_formateada or visita.get('fecha_inicio', datetime.now().strftime('%d/%m/%Y')),
        'hora_inicio': '09:00',
        'fecha_termino': fecha_formateada or visita.get('fecha_termino', datetime.now().strftime('%d/%m/%Y')),
        'hora_termino': '18:00',
        'normas': normas_list,
        'empresa_visitada': visita.get('cliente', ''),
        # Preferir campo `calle_numero` si existe, sino tratar de dividir `direccion`
        'calle_numero': visita.get('calle_numero') or (visita.get('direccion','').split(',')[0].strip() if visita.get('direccion') else ''),
        'colonia': colonia_mostrada,
        'municipio': visita.get('municipio', ''),
        'ciudad_estado': visita.get('ciudad_estado', ''),
        'inspectores': [s.strip() for s in (visita.get('supervisores_tabla') or visita.get('nfirma1') or '').split(',') if s.strip()],
        'NOMBRE_DE_INSPECTOR': (visita.get('supervisores_tabla') or visita.get('nfirma1') or '').strip(),
        'observaciones': visita.get('observaciones', ''),
        'acta': visita.get('folio_acta', ''),
        'cp': str(cp) if cp is not None else '',
        'folio_visita': visita.get('folio_visita', ''),
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
        'empresa_visitada': 'ARTICULOS DEPORTIVOS DECATHLON SA DE CV',
        'calle_numero': 'Parque industrial advance II',
        'colonia': 'Capula, 09876',
        'municipio': 'Tepotzotlán',
        'ciudad_estado': 'Estado de México',
        'cp': '09876',
        'firma_inspector': 'Firmas/AFLORES.png',
        'NOMBRE_DE_INSPECTOR': 'Arturo Flores Gómez',


    }
    # Crear carpetas si no existen
    os.makedirs("img", exist_ok=True)
    os.makedirs("Firmas", exist_ok=True)
    os.makedirs("data", exist_ok=True)
    
    # Generar PDF
    generar_acta_pdf(datos, "Plantillas PDF/Acta_inspeccion.pdf")

"""
Ejemplo de cómo debe aparecer la información en los formatos (no modificar nada más):

CALLE Y NO: Parque industrial advance II,
COLONIA O POBLACION: Capula, 54603
MUNICIPIO O ALCADIA: Tepotzotlán,
CIUDAD O ESTADO: Estado de México,

Estos datos se guardan en `data/historial_visitas.json` como:
"calle_numero": "Parque industrial advance II",
"colonia": "Capula",
"municipio": "Tepotzotlán",
"ciudad_estado": "Estado de México",
"cp": "54603",
"""
