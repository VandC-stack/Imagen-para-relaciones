"""Plantilla: Constancia de Conformidad

Esta plantilla genera una constancia en PDF usando img/Fondo.jpeg como fondo
y carga datos desde data/Clientes.json y data/Normas.json. También ofrece una
función para leer tabla_de_relacion.json y actualizar "TIPO DE DOCUMENTO" D->C.
"""
        # Cadena identificadora (cadena del dictamen/constancia) - centrada bajo el título
import os
import json
from datetime import datetime
import re

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
try:
    from plantillaPDF import cargar_firmas
except Exception:
    def cargar_firmas(path="data/Firmas.json"):
        # fallback: intentar cargar JSON directamente
        try:
            p = path
            if not os.path.exists(p):
                p = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')), path)
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
            m = {}
            for item in data:
                codigo = item.get('FIRMA') or item.get('codigo') or ''
                if codigo:
                    m[codigo] = item
            return m
        except Exception:
            return {}

try:
    from plantillaPDF import cargar_clientes as _cargar_clientes_ext, cargar_normas as _cargar_normas_ext
except Exception:
    _cargar_clientes_ext = None
    _cargar_normas_ext = None


class ConstanciaPDFGenerator:
    def __init__(self, datos: dict, base_dir: str | None = None):
        self.datos = datos or {}
        self.width, self.height = letter
        self.base_dir = base_dir or os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        # altura fija del encabezado (misma en todas las páginas)
        self.header_y = self.height - 40
        # posición del cursor (empieza debajo del encabezado)
        self.cursor_y = self.header_y

    def _fondo_path(self) -> str | None:
        p = os.path.join(self.base_dir, 'img', 'Fondo.jpeg')
        if os.path.exists(p):
            return p
        if os.path.exists('img/Fondo.jpeg'):
            return 'img/Fondo.jpeg'
        return None

    def dibujar_fondo(self, c: canvas.Canvas) -> None:
        p = self._fondo_path()
        if not p:
            return
        try:
            img = ImageReader(p)
            c.drawImage(img, 0, 0, width=self.width, height=self.height)
        except Exception:
            pass

    def _dibujar_texto_justificado(self, c: canvas.Canvas, x: float, y: float, texto: str, max_width: float,
                                  font_name: str = 'Helvetica', font_size: int = 10, leading: float = 12) -> None:
        """Dibuja texto justificado en el canvas y actualiza self.cursor_y.

        - `texto` se parte en líneas que caben en `max_width`.
        - Todas las líneas excepto la última se justifican repartiendo el espacio extra entre palabras.
        """
        c.setFont(font_name, font_size)
        lines = _dividir_texto(c, texto, max_width, font_name=font_name, font_size=font_size)
        cur_y = y
        for idx, line in enumerate(lines):
            words = line.split()
            if not words:
                cur_y -= leading
                continue
            # última línea -> alineado a la izquierda normal
            if idx == len(lines) - 1 or len(words) == 1:
                c.drawString(x, cur_y, line)
            else:
                total_words_w = sum(c.stringWidth(w, font_name, font_size) for w in words)
                gaps = len(words) - 1
                extra = max_width - total_words_w
                gap = extra / gaps if gaps > 0 else 0
                cur_x = x
                for w in words:
                    c.drawString(cur_x, cur_y, w)
                    cur_x += c.stringWidth(w, font_name, font_size) + gap
            cur_y -= leading
        # actualizar cursor_y con el valor final
        self.cursor_y = cur_y

    def construir_cadena_identificacion(self) -> str:
        """Construye la cadena identificadora y la guarda en self.datos['cadena'].

        Formato final:
        {year}049UCC{norma}{folio_formateado} Solicitud de Servicio: {year}049UCC{norma}{solicitud_formateado}-{lista}
        """
        # Año a partir de fecha_emision si está, else año actual
        fecha = self.datos.get('fecha_emision', '')
        year = None
        if fecha:
            # intentar formatos dd/mm/YYYY o YYYY-mm-dd
            m = re.search(r"(\d{4})", fecha)
            if m:
                year = m.group(1)
            else:
                m2 = re.search(r"(\d{2})/(\d{2})/(\d{4})", fecha)
                if m2:
                    year = m2.group(3)
        if not year:
            year = datetime.now().strftime('%Y')

        norma = (self.datos.get('norma') or '').replace('-', '')

        # folio_formateado: usar campo si existe, si no extraer dígitos del folio
        folio = str(self.datos.get('folio_constancia',''))
        folio_formateado = self.datos.get('folio_formateado')
        if not folio_formateado:
            nums = re.findall(r"\d+", folio)
            folio_formateado = nums[-1] if nums else folio

        solicitud_formateado = self.datos.get('solicitud_formateado') or self.datos.get('solicitud') or folio_formateado
        lista = str(self.datos.get('lista', '1'))

        cadena = f"{year}049UCC{norma}{folio_formateado} Solicitud de Servicio: {year}049UCC{norma}{solicitud_formateado}-{lista}"
        self.datos['cadena'] = cadena
        return cadena

    def dibujar_encabezado(self, c: canvas.Canvas) -> None:
        # asegurar altura fija del encabezado
        self.cursor_y = self.header_y
        # Logo (if present) at top-left (fallback to background watermark)
        logo_paths = [
            os.path.join(self.base_dir, 'img', 'Logo.png'),
            os.path.join(self.base_dir, 'img', 'VYC.png'),
            'img/Logo.png',
        ]
        for lp in logo_paths:
            if os.path.exists(lp):
                try:
                    c.drawImage(lp, 25 * mm, self.cursor_y - 8 * mm, width=35 * mm, preserveAspectRatio=True, mask='auto')
                    break
                except Exception:
                    pass

        # Title
        c.setFont('Helvetica-Bold', 14)
        c.drawCentredString(self.width / 2, self.cursor_y, 'CONSTANCIA DE CONFORMIDAD')
        self.cursor_y -= 10

        # Cadena identificadora (cadena del dictamen/constancia) - muestra en la parte superior
        c.setFont('Helvetica', 8)
        cadena = self.datos.get('cadena', '')
        if not cadena and self.datos.get('folio_constancia'):
            suffix = self.datos.get('cadena_suffix', '')
            cadena = f"{self.datos.get('folio_constancia')} {suffix}".strip()
        if cadena:
            max_w = self.width - 60 * mm
            y = self.cursor_y - 6
            # Mostrar años con dos dígitos (2026 -> 26) en la cadena identificadora
            try:
                display_cadena = re.sub(r"(\d{4})(?=049)", lambda m: m.group(1)[-2:], cadena)
            except Exception:
                display_cadena = cadena
            # Centrar cada línea de la cadena identificadora bajo el título
            for ln in _dividir_texto(c, display_cadena, max_w, font_name='Helvetica', font_size=8):
                c.drawCentredString(self.width / 2, y, ln)
                y -= 9
            self.cursor_y = y - 6
        else:
            self.cursor_y -= 18

        # small gap after header
        self.cursor_y -= 6

        # dibujar paginación si ya se definió contexto (opcional)

    def dibujar_paginacion(self, c: canvas.Canvas) -> None:
        # legacy: kept for backward-compatibility
        pass

    def dibujar_paginacion(self, c: canvas.Canvas, current: int, total: int) -> None:
        # Right-aligned page counter "Pagina X de Y" at top-right
        c.setFont('Helvetica', 8)
        right_x = self.width - 30
        # formato de código arriba
        c.drawRightString(right_x, self.header_y + 8, self.datos.get('formato_codigo', 'PT-F-208C-00-1'))
        pagina_txt = f"Página {current} de {total}"
        c.drawRightString(right_x, self.header_y - 6, pagina_txt)

    def dibujar_footer(self, c: canvas.Canvas) -> None:
        # Yellow band at bottom with organization info (similar to sample)
        band_height = 18 * mm
        c.saveState()
        c.setFillColor(colors.HexColor('#f6d200'))
        c.rect(0, 0, self.width, band_height, stroke=0, fill=1)
        c.setFillColor(colors.black)
        c.setFont('Helvetica', 8)
        info = self.datos.get('footer_info', 'Verificación y Control UVA, S.C.  Álamos 104, Ofic. 202, Valle de los Pinos 1A, Tlalnepantla, Méx. C.P. 54040.')
        c.drawString(25 * mm, 6 * mm, info)
        # website on right
        website = self.datos.get('website', 'www.vyc.com.mx')
        c.drawRightString(self.width - 25 * mm, 6 * mm, website)
        c.restoreState()

    def dibujar_datos_basicos(self, c: canvas.Canvas) -> None:
        # Mostrar: Norma y nombre de la norma, No. de contrato, Fecha de contrato
        x = 25 * mm
        right_x = self.width - 25 * mm

        # No mostrar número de NOM en la parte superior según solicitud
        # (se mantiene posible referencia a la norma dentro del cuerpo legal)

        # No. de contrato (valor en negritas)
        no_contrato = str(self.datos.get('no_contrato', '') or self.datos.get('no_de_contrato', ''))
        c.setFont('Helvetica', 9)
        c.drawString(x, self.cursor_y, 'No. de contrato:')
        c.setFont('Helvetica-Bold', 9)
        c.drawString(x + 40 * mm, self.cursor_y, no_contrato)
        self.cursor_y -= 12

        # Fecha de contrato (valor en negritas)
        fecha_contrato = str(self.datos.get('fecha_contrato', '') or '')
        c.setFont('Helvetica', 9)
        c.drawString(x, self.cursor_y, 'Fecha de contrato:')
        c.setFont('Helvetica-Bold', 9)
        c.drawString(x + 40 * mm, self.cursor_y, fecha_contrato)
        self.cursor_y -= 12

        # Fecha de emisión (ahora después de la fecha de contrato)
        # Mostrar Fecha de Emisión en UNA sola línea con el label primero
        fecha_emision = str(self.datos.get('fecha_emision', '') or '')
        fecha_larga = _formato_fecha_larga(fecha_emision)
        combined = f"Fecha de Emisión: {fecha_larga}" if fecha_larga else 'Fecha de Emisión:'
        fsize = 9
        min_fsize = 7
        avail_w = right_x - x
        while fsize >= min_fsize and c.stringWidth(combined, 'Helvetica-Bold', fsize) > avail_w:
            fsize -= 1
        c.setFont('Helvetica-Bold', fsize)
        c.drawRightString(right_x, self.cursor_y, combined)
        self.cursor_y -= 12

    def dibujar_cuerpo_legal(self, c: canvas.Canvas) -> None:
        x = 25 * mm
        max_w = 165 * mm
        norma = self.datos.get('norma', '')
        nombre = self.datos.get('nombre_norma', '')

        texto = (
            "De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, 68 y 140 "
            "de la Ley de Infraestructura de la Calidad; 50 del Reglamento de la Ley Federal de Metrología y Normalización; "
            "Punto 2.4.8 Fracción I ACUERDO por el que la Secretaría de Economía emite Reglas y criterios de carácter general "
            "en materia de comercio exterior; publicado en el Diario Oficial de la Federación el 09 de mayo de 2022 y posteriores "
            f"modificaciones; esta Unidad de Inspección, hace constar que la Información Comercial contenida en el producto cuya "
            f"etiqueta muestra aparece en esta Constancia, cumple con la Norma Oficial Mexicana {norma} ({nombre}), modificación del 27 de marzo de 2020, "
            f"ACUERDO por el cual se establecen los Criterios para la implementación, verificación y vigilancia, así como para la evaluación "
            f"de la conformidad de la Modificación a la Norma Oficial Mexicana {norma} ({nombre}), publicada el 27 de marzo de 2020 y la Nota Aclaratoria que emiten "
            f"la Secretaría de Economía y la Secretaría de Salud a través de la Comisión Federal para la Protección contra Riesgos Sanitarios a la Modificación "
            f"a la Norma Oficial Mexicana {norma}, {nombre}.")

        c.setFont('Helvetica', 9)
        self._dibujar_texto_justificado(c, x, self.cursor_y, texto, max_w, font_name='Helvetica', font_size=9, leading=12)
        self.cursor_y -= 20

    def dibujar_condiciones(self, c: canvas.Canvas) -> None:
        # Dibujar sección de condiciones con barras arriba y abajo y título centrado
        left = 25 * mm
        right = self.width - 25 * mm
        line_y = self.cursor_y
        # líneas más gruesas según petición
        c.setLineWidth(1.2)
        c.line(left, line_y, right, line_y)

        # Título centrado
        title_y = line_y - 8
        c.setFont('Helvetica-Bold', 11)
        c.drawCentredString(self.width / 2, title_y, 'Condiciones de la Constancia')

        # Preparar contenido
        self.cursor_y = title_y - 12
        condiciones = [
            '1. Este documento sólo ampara la información contenida en el producto cuya etiqueta muestra se presenta en esta Constancia.',
            '2. Cualquier modificación a la etiqueta debe ser sometida a la consideración de la Unidad de Inspección Acreditada y Aprobada en los términos de la Ley de Infraestructura de la Calidad, para que inspeccione su cumplimiento con la Norma Oficial Mexicana aplicable.',
            f"3. Esta Constancia sólo ampara el cumplimiento con la Norma Oficial Mexicana {self.datos.get('norma','')} ({self.datos.get('nombre_norma','')})."
        ]
        c.setFont('Helvetica', 9)
        for cond in condiciones:
            self._dibujar_texto_justificado(c, left + 4 * mm, self.cursor_y, cond, (right - left) - 8 * mm, font_name='Helvetica', font_size=9, leading=11)
            # _dibujar_texto_justificado atualiza self.cursor_y
            self.cursor_y -= 4

        # Línea inferior que cierra el bloque
        bottom_line_y = self.cursor_y - 6
        c.line(left, bottom_line_y, right, bottom_line_y)
        self.cursor_y = bottom_line_y - 10

    def dibujar_producto(self, c: canvas.Canvas) -> None:
        x = 25 * mm
        c.setFont('Helvetica-Bold', 10)
        # Producto: etiqueta en negritas para el dato
        c.setFont('Helvetica', 10)
        c.drawString(x, self.cursor_y, 'Producto: ')
        prod = str(self.datos.get('producto',''))
        c.setFont('Helvetica-Bold', 10)
        c.drawString(x + 40 * mm, self.cursor_y, prod)
        self.cursor_y -= 20

    def dibujar_tabla_relacion(self, c: canvas.Canvas) -> None:
        # Diseño compacto y con bordes uniformes; ajustar contenido para que quepa en celdas
        margin_x = 28 * mm
        total_w = self.width - 2 * margin_x
        x = margin_x

        # Column widths: hacer la columna 'CONTENIDO NETO' más pequeña
        col1 = 34 * mm
        col3 = 28 * mm
        col2 = total_w - col1 - col3

        title_h = 9 * mm
        header_h = 8 * mm
        row_h = 14 * mm

        table_top = self.cursor_y
        bottom_y = table_top - (title_h + header_h + row_h)

        # Dibujar bandas (rellenos) primero
        c.setFillColor(colors.whitesmoke)
        c.rect(x, table_top - title_h, total_w, title_h, stroke=0, fill=1)
        c.setFillColor(colors.HexColor('#efefef'))
        c.rect(x, table_top - title_h - header_h, total_w, header_h, stroke=0, fill=1)
        c.setFillColor(colors.black)

        # Textos de título y encabezados
        c.setFont('Helvetica-Bold', 10)
        c.drawCentredString(x + total_w / 2, table_top - title_h / 2 - 1, 'RELACIÓN CORRESPONDIENTE')
        c.setFont('Helvetica-Bold', 8)
        c.drawCentredString(x + col1 / 2, table_top - title_h - header_h / 2 - 2, 'CÓDIGO')
        c.drawCentredString(x + col1 + col2 / 2, table_top - title_h - header_h / 2 - 2, 'MEDIDAS')
        c.drawCentredString(x + col1 + col2 + col3 / 2, table_top - title_h - header_h / 2 - 2, 'CONTENIDO NETO')

        c.setLineWidth(0.6)
        c.setStrokeColor(colors.black)
        # Caja externa
        c.rect(x, bottom_y, total_w, title_h + header_h + row_h, stroke=1, fill=0)
        # Líneas verticales (desde la parte superior de la tabla hasta el fondo)
        top_y = table_top
        c.line(x + col1, top_y, x + col1, bottom_y)
        c.line(x + col1 + col2, top_y, x + col1 + col2, bottom_y)
        # Línea separadora entre encabezado y fila
        c.line(x, top_y - header_h, x + total_w, top_y - header_h)

        # Preparar valores
        codigo = str(self.datos.get('codigo', '')).strip() or '7 503049 695501'
        medida = str(self.datos.get('medida', '')).strip() or '17 cm de ancho x 15.35 cm de alto'
        contenido = str(self.datos.get('contenido_neto', '')).strip() or '355 ml'

        def _draw_cell_text(text, cell_x, cell_w, cell_y, cell_h, align='left'):
            fsize = 8
            min_fsize = 6
            while fsize >= min_fsize:
                lines = _dividir_texto(c, text, cell_w - 6 * mm, font_name='Helvetica-Bold', font_size=fsize)
                if len(lines) <= 2:
                    break
                fsize -= 1
            c.setFont('Helvetica-Bold', fsize)
            leading = fsize + 2
            lines = lines[:2]
            y_text = cell_y + cell_h - (cell_h - (len(lines) * leading)) / 2 - leading + 2
            for ln in lines:
                if align == 'left':
                    c.drawString(cell_x + 3 * mm, y_text, ln)
                else:
                    c.drawRightString(cell_x + cell_w - 3 * mm, y_text, ln)
                y_text -= leading

        # Dibujar rectángulos de las celdas (asegurar bordes visibles)
        c.rect(x, bottom_y, col1, row_h, stroke=1, fill=0)
        c.rect(x + col1, bottom_y, col2, row_h, stroke=1, fill=0)
        c.rect(x + col1 + col2, bottom_y, col3, row_h, stroke=1, fill=0)

        # Dibujar contenidos ajustados
        _draw_cell_text(codigo, x, col1, bottom_y, row_h, align='left')
        _draw_cell_text(medida, x + col1, col2, bottom_y, row_h, align='left')
        _draw_cell_text(contenido, x + col1 + col2, col3, bottom_y, row_h, align='right')

        # Actualizar cursor
        self.cursor_y = bottom_y - 8 * mm

    def dibujar_observaciones(self, c: canvas.Canvas) -> None:
        x = 25 * mm
        max_w = 165 * mm
        obs = 'OBSERVACIONES: EN CUMPLIMIENTO CON LOS PUNTOS 4.2.6 Y 4.2.7 DE LA NORMA LOS DATOS DE FECHA DE CONSUMO PREFERENTE Y LOTE SE ENCUENTRAN DECLARADOS EN EL ENVASE DEL PRODUCTO. ESTE PRODUCTO FUE INSPECCIONADO EN CUMPLIMIENTO BAJO LA FASE 2 DE LA NOM CON VIGENCIA AL 31 DE DICIEMBRE DE 2027 Y FASE 3 DE LA NOM CON ENTRADA EN VIGOR A PARTIR DEL 01 DE ENERO DEL 2028.'
        c.setFont('Helvetica', 8)
        self._dibujar_texto_justificado(c, x, self.cursor_y, obs, max_w, font_name='Helvetica', font_size=8, leading=10)
        self.cursor_y -= 30

    def dibujar_firma(self, c: canvas.Canvas) -> None:
        # Imprimir firmas en página(s) final(es) con diseño de dos columnas similar al Dictamen
        # Reservar página nueva y dibujar encabezado en altura fija
        try:
            c.showPage()
        except Exception:
            pass
        self.cursor_y = self.header_y
        try:
            self.dibujar_fondo(c)
        except Exception:
            pass
        try:
            self.dibujar_encabezado(c)
            # paginación será añadida por el llamador con el número correcto
        except Exception:
            pass

        # Cargar mapa de firmas (si existe)
        firmas_map = {}
        try:
            firmas_map = cargar_firmas()
        except Exception:
            firmas_map = {}

        # Preparar datos: intentar obtener dos firmantes
        # Preferir nombres suministrados en self.datos
        nombre1 = self.datos.get('nfirma1') or ''
        nombre2 = self.datos.get('nfirma2') or ''
        img1 = None
        img2 = None

        # Si no hay nombres, intentar sacar de Firmas.json
        if not nombre1 or not nombre2:
            for k, v in (firmas_map or {}).items():
                n = v.get('NOMBRE DE INSPECTOR') or v.get('nombre') or v.get('NOMBRE') or ''
                if not nombre1 and 'Gabriel' in n:
                    nombre1 = nombre1 or n
                    img1 = v.get('IMAGEN') or v.get('imagen')
                if not nombre2 and ('Arturo' in n or 'AFLORES' in (k or '').upper()):
                    nombre2 = nombre2 or n
                    img2 = v.get('IMAGEN') or v.get('imagen')
        # Fallbacks
        if not nombre1:
            nombre1 = 'Nombre del Inspector'
        if not nombre2:
            nombre2 = 'ARTURO FLORES GÓMEZ'
        # localizar rutas de imagen por si no vinieron en Firmas.json
        if not img2:
            candidate = os.path.join(self.base_dir, 'Firmas', 'AFLORES.png')
            if os.path.exists(candidate):
                img2 = candidate
            elif os.path.exists('Firmas/AFLORES.png'):
                img2 = 'Firmas/AFLORES.png'

        # Column coordinates
        left_x = 25 * mm
        right_x = self.width / 2 + 10 * mm

        

        # Draw images if available
        # reducir ligeramente el tamaño de las firmas para que queden más armoniosas
        sig_h = 22 * mm
        sig_w1 = 50 * mm
        sig_w2 = 50 * mm

        # Left signature
        if img1:
            try:
                p1 = img1 if os.path.isabs(img1) or os.path.exists(img1) else os.path.join(self.base_dir, img1)
                if os.path.exists(p1):
                    im1 = ImageReader(p1)
                    iw, ih = im1.getSize()
                    w = iw * (sig_h / ih)
                    sig_w1 = w
                    c.drawImage(im1, left_x, self.cursor_y - sig_h, width=w, height=sig_h, mask='auto')
            except Exception:
                pass

        # Right signature
        if img2:
            try:
                p2 = img2 if os.path.isabs(img2) or os.path.exists(img2) else os.path.join(self.base_dir, img2)
                if os.path.exists(p2):
                    im2 = ImageReader(p2)
                    iw2, ih2 = im2.getSize()
                    w2 = iw2 * (sig_h / ih2)
                    sig_w2 = w2
                    c.drawImage(im2, right_x, self.cursor_y - sig_h, width=w2, height=sig_h, mask='auto')
            except Exception:
                pass

        # Move cursor under signatures (leave extra space so names are lower)
        y_after = self.cursor_y - sig_h - 12 * mm

        # Draw signature lines between image and printed name
        y_line = y_after + 8 * mm
        try:
            c.setLineWidth(0.6)
            c.line(left_x, y_line, left_x + sig_w1, y_line)
        except Exception:
            pass
        try:
            c.setLineWidth(0.6)
            c.line(right_x, y_line, right_x + sig_w2, y_line)
        except Exception:
            pass

        # Left name and role (ligeramente más pequeñas)
        c.setFont('Helvetica-Bold', 10)
        c.drawString(left_x, y_after, nombre1)
        c.setFont('Helvetica', 8)
        c.drawString(left_x, y_after - 12, 'Inspector')

        # Right name and role (ligeramente más pequeñas)
        c.setFont('Helvetica-Bold', 10)
        c.drawString(right_x, y_after, nombre2)
        c.setFont('Helvetica', 8)
        c.drawString(right_x, y_after - 12, 'Responsable de Supervisión UI')
        self.cursor_y = y_after - 30

    def generar(self, salida: str) -> str:
        # Si no se especifica salida, guardar en data/Constancias bajo el base_dir
        if not salida:
            data_dir = os.path.join(self.base_dir, 'data')
            const_dir = os.path.join(data_dir, 'Constancias')
            os.makedirs(const_dir, exist_ok=True)
            fol = str(self.datos.get('folio_constancia') or '')
            safe = fol.replace('/', '_').replace(' ', '_') or datetime.now().strftime('%Y%m%d_%H%M%S')
            salida = os.path.join(const_dir, f'Constancia_{safe}.pdf')

        c = canvas.Canvas(salida, pagesize=letter)
        self.cursor_y = self.header_y
        # calcular páginas totales: página principal + N páginas de evidencia (4 por página) + firmas
        evidencias = (self.datos.get('evidencias_lista') or [])
        import math
        evidence_pages = max(1, math.ceil(len(evidencias) / 4))
        signature_page = 1
        total_pages = 1 + evidence_pages + signature_page
        current_page = 1
        try:
            # Preparar datos: construir cadena identificadora y cargar catálogos
            try:
                self.construir_cadena_identificacion()
            except Exception:
                pass

            # Cargar clientes, normas y firmas desde data/
            data_dir = os.path.join(self.base_dir, 'data')
            clientes_path = os.path.join(data_dir, 'Clientes.json')
            normas_path = os.path.join(data_dir, 'Normas.json')
            firmas_path = os.path.join(data_dir, 'Firmas.json')
            try:
                if _cargar_clientes_ext:
                    clientes_map = _cargar_clientes_ext(clientes_path)
                else:
                    clientes_map = _cargar_clientes(clientes_path)
            except Exception:
                clientes_map = {}
            try:
                if _cargar_normas_ext:
                    normas_map = _cargar_normas_ext(normas_path)
                else:
                    normas_map = _cargar_normas(normas_path)
            except Exception:
                normas_map = {}
            try:
                firmas_map = cargar_firmas(firmas_path)
            except Exception:
                firmas_map = {}

            # Rellenar nombre_norma si está vacío y norma conocida
            try:
                if not self.datos.get('nombre_norma') and self.datos.get('norma'):
                    nn = normas_map.get(self.datos.get('norma')) if isinstance(normas_map, dict) else None
                    if not nn:
                        # intentar buscar por número dentro de la NOM
                        nom = str(self.datos.get('norma'))
                        for k, v in normas_map.items():
                            if k in nom or nom in k:
                                nn = v
                                break
                    if nn:
                        self.datos['nombre_norma'] = nn
            except Exception:
                pass

            # dibujar fondo en la primera página si existe
            self.dibujar_fondo(c)
        except Exception:
            pass
        # asegurar que la cadena esté presente en el encabezado
        try:
            if not self.datos.get('cadena'):
                self.construir_cadena_identificacion()
        except Exception:
            pass
        # Página principal: encabezado y secciones iniciales
        self.dibujar_encabezado(c)
        try:
            self.dibujar_paginacion(c, current_page, total_pages)
        except Exception:
            pass
        self.dibujar_datos_basicos(c)
        self.dibujar_cuerpo_legal(c)
        self.dibujar_condiciones(c)

        # Tabla de relación inmediatamente después de las condiciones
        self.dibujar_tabla_relacion(c)

        
        self.dibujar_observaciones(c)

        # Añadir apartado para pegar evidencia fotográfica (páginas nuevas)
        try:
            # dividir evidencias en páginas de 4
            for p in range(evidence_pages):
                start = p * 4
                page_items = evidencias[start:start + 4]
                current_page += 1
                self.dibujar_evidencia(c, page_items)
                try:
                    self.dibujar_paginacion(c, current_page, total_pages)
                except Exception:
                    pass
        except Exception:
            pass

        # Firmas al final del documento
        try:
            current_page += 1
            self.dibujar_firma(c)
            try:
                self.dibujar_paginacion(c, current_page, total_pages)
            except Exception:
                pass
        except Exception:
            pass
        c.save()
        return salida

    def dibujar_evidencia(self, c: canvas.Canvas, page_items: list | None = None) -> None:
        """Dibuja una página de evidencia con hasta 4 elementos (page_items).

        - `page_items` es lista con hasta 4 elementos. Cada elemento puede ser:
          - '${IMAGEN}' para dejar el placeholder
          - ruta a imagen (string) para dibujarla
          - None para dejar el placeholder
        Si `page_items` es None, se dibujan 4 placeholders.
        """
        try:
            c.showPage()
        except Exception:
            pass
        # Reservar espacio superior para encabezado en la página de evidencia
        self.cursor_y = self.header_y
        try:
            self.dibujar_fondo(c)
        except Exception:
            pass
        try:
            self.dibujar_encabezado(c)
        except Exception:
            pass

        items = page_items or []
        # Título
        c.setFont('Helvetica-Bold', 12)
        c.drawCentredString(self.width / 2, self.cursor_y, 'EVIDENCIA FOTOGRÁFICA')
        self.cursor_y -= 20

        # Márgenes y tamaños de caja (2x2)
        margin_x = 25 * mm
        margin_y = 30 * mm
        gap = 10 * mm
        box_w = (self.width - 2 * margin_x - gap) / 2
        box_h = (self.height - self.cursor_y - margin_y - 40 * mm) / 2
        if box_h <= 40 * mm:
            box_h = 60 * mm

        y_top = self.cursor_y
        num = 1
        for r in range(2):
            y = y_top - r * (box_h + gap)
            for ccol in range(2):
                x = margin_x + ccol * (box_w + gap)
                c.rect(x, y - box_h, box_w, box_h, stroke=1, fill=0)
                idx = (r * 2) + ccol
                val = items[idx] if idx < len(items) else None
                if val == '${IMAGEN}' or val is None:
                    c.setFont('Helvetica-Bold', 14)
                    c.drawCentredString(x + box_w / 2, y - box_h / 2, '${IMAGEN}')
                elif isinstance(val, str) and os.path.exists(val):
                    try:
                        im = ImageReader(val)
                        iw, ih = im.getSize()
                        scale = min((box_w - 6 * mm) / iw, (box_h - 6 * mm) / ih)
                        w = iw * scale
                        h = ih * scale
                        c.drawImage(im, x + (box_w - w) / 2, y - box_h + (box_h - h) / 2, width=w, height=h, mask='auto')
                    except Exception:
                        c.setFont('Helvetica', 8)
                        c.drawCentredString(x + box_w / 2, y - box_h + 6 * mm, f'Evidencia {num}')
                else:
                    c.setFont('Helvetica', 8)
                    c.drawCentredString(x + box_w / 2, y - box_h + 6 * mm, f'Evidencia {num}')
                num += 1

        self.cursor_y = margin_y


def _dividir_texto(c: canvas.Canvas, texto: str, max_width: float, font_name: str = 'Helvetica', font_size: int = 10):
    palabras = texto.split()
    lineas = []
    actual = ''
    for p in palabras:
        prueba = f"{actual} {p}".strip()
        if c.stringWidth(prueba, font_name, font_size) <= max_width:
            actual = prueba
        else:
            if actual:
                lineas.append(actual)
            actual = p
    if actual:
        lineas.append(actual)
    return lineas


def _formato_fecha_larga(fecha_str: str) -> str:
    """Intenta convertir una fecha corta (dd/mm/YYYY, YYYY-mm-dd, etc.)
    a un formato largo en español: 'miércoles 19 de noviembre de 2026'.
    Si no puede parsear, devuelve la cadena original.
    """
    if not fecha_str:
        return ''
    # limpiar
    s = fecha_str.strip()
    meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
    dias = ['lunes','martes','miércoles','jueves','viernes','sábado','domingo']
    fmt_candidates = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y']
    for fmt in fmt_candidates:
        try:
            dt = datetime.strptime(s, fmt)
            dia_nombre = dias[dt.weekday()]
            mes_nombre = meses[dt.month - 1]
            return f"{dia_nombre} {dt.day} de {mes_nombre} de {dt.year}"
        except Exception:
            continue
    # intentar extraer dd/mm/YYYY dentro de la cadena
    m = re.search(r"(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})", s)
    if m:
        try:
            dt = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            dia_nombre = dias[dt.weekday()]
            mes_nombre = meses[dt.month - 1]
            return f"{dia_nombre} {dt.day} de {mes_nombre} de {dt.year}"
        except Exception:
            pass
    return s


def _cargar_clientes(path: str) -> dict:
    clientes = {}
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    for item in data:
                        nombre = item.get('CLIENTE') or item.get('CLIENTE', '')
                        if nombre:
                            clientes[nombre.upper()] = item
                elif isinstance(data, dict):
                    for v in data.values():
                        if isinstance(v, dict) and v.get('CLIENTE'):
                            clientes[v.get('CLIENTE','').upper()] = v
    except Exception:
        pass
    return clientes


def _cargar_normas(path: str) -> dict:
    normas = {}
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    for n in data:
                        nom = n.get('NOM') or n.get('NOM', '')
                        nombre = n.get('NOMBRE') or n.get('NOMBRE', '')
                        if nom:
                            normas[nom] = nombre
                elif isinstance(data, dict):
                    for item in data.values():
                        if isinstance(item, dict) and item.get('NOM'):
                            normas[item.get('NOM')] = item.get('NOMBRE', '')
    except Exception:
        pass
    return normas


def _actualizar_tabla_relacion(path: str) -> None:
    if not os.path.exists(path):
        return
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        changed = False
        if isinstance(data, list):
            for row in data:
                if isinstance(row, dict):
                    key = 'TIPO DE DOCUMENTO' if 'TIPO DE DOCUMENTO' in row else 'TIPO_DE_DOCUMENTO'
                    if row.get(key) == 'D':
                        row[key] = 'C'
                        changed = True
        if changed:
            try:
                with open(path + '.bak', 'w', encoding='utf-8') as b:
                    json.dump(data, b, ensure_ascii=False, indent=2)
            except Exception:
                pass
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def generar_constancia_desde_visita(folio_visita: str | None = None, salida: str | None = None) -> str:
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    data_dir = os.path.join(base, 'data')
    hist = os.path.join(data_dir, 'historial_visitas.json')
    tabla = os.path.join(data_dir, 'tabla_de_relacion.json')
    clientes_p = os.path.join(data_dir, 'Clientes.json')
    normas_p = os.path.join(data_dir, 'Normas.json')

    if not os.path.exists(hist):
        raise FileNotFoundError(hist)
    with open(hist, 'r', encoding='utf-8') as f:
        historial = json.load(f)
    visitas = historial.get('visitas', []) if isinstance(historial, dict) else historial
    if not visitas:
        raise ValueError('No hay visitas en el historial')
    visita = None
    if folio_visita:
        for v in visitas:
            if v.get('folio_visita') == folio_visita or v.get('folio') == folio_visita:
                visita = v
                break
    visita = visita or visitas[-1]

    clientes = _cargar_clientes(clientes_p)
    normas = _cargar_normas(normas_p)

    _actualizar_tabla_relacion(tabla)

    producto = marca = modelo = ''
    if os.path.exists(tabla):
        try:
            with open(tabla, 'r', encoding='utf-8') as f:
                t = json.load(f)
                if isinstance(t, list) and t:
                    r = t[0]
                    producto = r.get('DESCRIPCION','')
                    marca = r.get('MARCA','')
                    modelo = r.get('MODELO','')
        except Exception:
            pass

    norma_str = ''
    if visita.get('norma'):
        norma_str = visita.get('norma').split(',')[0].strip()
    nombre_norma = normas.get(norma_str, '')

    cliente = visita.get('cliente','')
    rfc = (clientes.get(cliente.upper(), {}) or {}).get('RFC','')

    fecha = visita.get('fecha_termino') or visita.get('fecha') or datetime.now().strftime('%d/%m/%Y')

    fol = (visita.get('folio_visita') or visita.get('folio') or '').replace('UDC','UCC')
    datos = {
        'folio_constancia': fol,
        'fecha_emision': fecha,
        'cliente': cliente,
        'rfc': rfc,
        'norma': norma_str,
        'nombre_norma': nombre_norma,
        'producto': producto,
        'marca': marca,
        'modelo': modelo,
    }

    if not salida:
        salida = os.path.join(base, f'Constancia_{fol or "constancia"}.pdf')

    gen = ConstanciaPDFGenerator(datos, base_dir=base)
    return gen.generar(salida)


if __name__ == '__main__':
    # demo rápido
    datos_demo = {
        'folio_constancia': 'UCC-DEMO-0001',
        'fecha_emision': datetime.now().strftime('%d/%m/%Y'),
        'cliente': 'CLIENTE DEMO',
        'rfc': 'XAXX010101000',
        'norma': 'NOM-XXX-XXXX',
        'nombre_norma': 'Nombre de la norma demo',
        'producto': 'PRODUCTO DEMO',
        'marca': 'MARCA DEMO',
        'modelo': 'MODELO DEMO',
    }
    out = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')), 'Plantillas PDF', 'Constancia_demo.pdf')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    ConstanciaPDFGenerator(datos_demo).generar(out)
