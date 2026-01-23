"""Plantilla: Constancia de Conformidad

Esta plantilla genera una constancia en PDF usando img/Fondo.jpg como fondo
y carga datos desde data/Clientes.json y data/Normas.json. También ofrece una
función para leer tabla_de_relacion.json y actualizar "TIPO DE DOCUMENTO" D->C.
"""
        # Cadena identificadora (cadena del dictamen/constancia) - centrada bajo el título
import os
import json
from datetime import datetime
import re
import shutil

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm, inch
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import tempfile


# Canvas personalizado para numerar páginas como "Página X de Y"
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
        # Si se definió un encabezado, dibujarlo inmediatamente en la nueva página
        try:
            # título en la posición alta (misma que usa la plantilla)
            if hasattr(self, 'header_title') and self.header_title:
                try:
                    # título (más alto en la página)
                    self.setFont('Helvetica-Bold', 12)
                    self.drawCentredString(self._pagesize[0] / 2, self._pagesize[1] - 58, self.header_title)
                except Exception:
                    pass
            # cadena identificadora justo debajo del título
            if hasattr(self, 'header_chain') and self.header_chain:
                try:
                    self.setFont('Helvetica', 8)
                    self.drawCentredString(self._pagesize[0] / 2, self._pagesize[1] - 74, self.header_chain)
                except Exception:
                    pass
        except Exception:
            pass

    def save(self):
        # añadir estado de la última página
        self._saved_page_states.append(dict(self.__dict__))
        page_count = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            # Redibujar encabezado (si existe) para cada página al reconstruir
            try:
                if hasattr(self, 'header_title') and self.header_title:
                    try:
                        self.setFont('Helvetica-Bold', 12)
                        self.drawCentredString(self._pagesize[0] / 2, self._pagesize[1] - 58, self.header_title)
                    except Exception:
                        pass
                if hasattr(self, 'header_chain') and self.header_chain:
                    try:
                        self.setFont('Helvetica', 8)
                        self.drawCentredString(self._pagesize[0] / 2, self._pagesize[1] - 74, self.header_chain)
                    except Exception:
                        pass
            except Exception:
                pass
            self.draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count: int) -> None:
        try:
            self.setFont('Helvetica', 8)
            text = f"Página {self._pageNumber} de {page_count}"
            # dibujar en la esquina superior derecha, con un pequeño margen
            x = self._pagesize[0] - 30
            y = self._pagesize[1] - 40
            self.drawRightString(x, y, text)
        except Exception:
            pass

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
try:
    from plantillaPDF import cargar_tabla_relacion as _cargar_tabla_relacion_ext
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_JUSTIFY
except Exception:
    _cargar_tabla_relacion_ext = None

try:
    from plantillaPDF import validar_acreditacion_inspector as _validar_acreditacion_inspector
except Exception:
    _validar_acreditacion_inspector = None


class ConstanciaPDFGenerator:
    def __init__(self, datos: dict, base_dir: str | None = None):
        self.datos = datos or {}
        self.width, self.height = letter
        self.base_dir = base_dir or os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        # bajar la posición inicial para que las firmas queden más abajo en la hoja
        self.cursor_y = self.height - 90

    def _fondo_path(self) -> str | None:
        p = os.path.join(self.base_dir, 'img', 'Fondo.jpg')
        if os.path.exists(p):
            return p
        if os.path.exists('img/Fondo.jpg'):
            return 'img/Fondo.jpg'
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

        # conservar los guiones en la norma (ej. NOM-051-SCFI/SSA1-2010)
        norma = (self.datos.get('norma') or '').strip()

        # folio_formateado: usar campo si existe, si no extraer dígitos del folio
        folio = str(self.datos.get('folio_constancia',''))
        folio_formateado = self.datos.get('folio_formateado')
        if not folio_formateado:
            nums = re.findall(r"\d+", folio)
            folio_formateado = nums[-1] if nums else folio

        # Solicitud: preferir campo ya formateado; si no, intentar dividir "NNN.../YY" -> numero y año
        solicitud_raw = str(
            (self.datos.get('solicitud') or self.datos.get('Solicitud') or self.datos.get('SOLICITUD') or '')
        ).strip()

        # Si no viene en los datos principales, intentar extraer desde tabla_relacion (primera fila)
        if not solicitud_raw:
            try:
                tr = self.datos.get('tabla_relacion') or []
                if isinstance(tr, list) and tr:
                    first = tr[0]
                    solicitud_raw = str(first.get('SOLICITUD') or first.get('Solicitud') or first.get('solicitud') or '').strip()
            except Exception:
                solicitud_raw = solicitud_raw

        solicitud_num = ''
        solicitud_year_two = None
        if solicitud_raw:
            m = re.match(r"^\s*(\d+)(?:[\/-](\d{2,4}))?\s*$", solicitud_raw)
            if m:
                solicitud_num = m.group(1)
                suf = m.group(2)
                if suf:
                    solicitud_year_two = suf[-2:]
            else:
                nums = re.findall(r"\d+", solicitud_raw)
                solicitud_num = nums[0] if nums else solicitud_raw

        # si el usuario pasó un campo de solicitud_formateado preferido, úsalo como número
        if not solicitud_num:
            solicitud_num = str(self.datos.get('solicitud_formateado') or '')

        # preparar la lista
        lista = str(self.datos.get('lista', '1'))

        # determinar año de la parte 'Solicitud de Servicio' (2 dígitos)
        if solicitud_year_two:
            sol_year_two = solicitud_year_two
        else:
            # extraer dos últimos dígitos del `year` principal (puede ser YYYY o YY)
            sol_year_two = str(year)[-2:]

        # asegurar formato de solicitud (6 dígitos cuando sea numérica)
        solicitud_formatted = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num

        cadena = f"{year}049UCC{norma}{folio_formateado} Solicitud de Servicio: {sol_year_two}049UCC{norma}{solicitud_formatted}-{lista}"
        self.datos['cadena'] = cadena
        return cadena

    def dibujar_encabezado(self, c: canvas.Canvas) -> None:
        # Logo (if present) at top-left (fallback to background watermark)
        logo_paths = [
            os.path.join(self.base_dir, 'img', 'Logo.png'),
            os.path.join(self.base_dir, 'img', 'VYC.png'),
            'img/Logo.png',
        ]
        # Dibujar logo en la parte superior izquierda, coordenada fija respecto al tope
        logo_y = self.height - 88
        for lp in logo_paths:
            if os.path.exists(lp):
                try:
                    c.drawImage(lp, 25 * mm, logo_y, width=35 * mm, preserveAspectRatio=True, mask='auto')
                    break
                except Exception:
                    pass

        # # Title
        # c.setFont('Helvetica-Bold', 12)
        # c.drawCentredString(self.width / 2, self.cursor_y, 'CONSTANCIA DE CONFORMIDAD')
        # self.cursor_y -= 10

        # Mostrar fecha de contrato (desde `self.datos` o, como respaldo, desde data/Clientes.json)
        try:
            fecha_contrato = self.datos.get('fecha_contrato') or self.datos.get('fecha_de_contrato') or ''
            if not fecha_contrato:
                try:
                    clientes_path = os.path.join(self.base_dir, 'data', 'Clientes.json')
                    clientes_map = _cargar_clientes(clientes_path)
                    cliente_name = (self.datos.get('cliente') or '').upper()
                    fecha_contrato = (clientes_map.get(cliente_name, {}) or {}).get('FECHA_DE_CONTRATO', '')
                except Exception:
                    fecha_contrato = ''
        except Exception:
            pass
        # El título y la cadena se dibujan en `generar()` y en el canvas final durante el guardado.
        # Solo dejar un pequeño espacio antes del contenido siguiente.
        self.cursor_y -= 8

    def dibujar_paginacion(self, c: canvas.Canvas) -> None:
        # Right-aligned codes and page number similar to sample
        c.setFont('Helvetica', 8)
        right_x = self.width - 30
        # Código de formato en la parte superior derecha (la numeración "Página X de Y" la dibuja NumberedCanvas)
        c.drawRightString(right_x, self.height - 30, self.datos.get('formato_codigo', 'PT-F-208C-00-1'))

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

        # cliente usado para búsquedas en Clientes.json
        cliente = (self.datos.get('cliente') or '').strip()

        # Cliente (valor en negritas)
        cliente_display = str(self.datos.get('cliente') or cliente or '')
        c.setFont('Helvetica', 9)
        c.drawString(x, self.cursor_y, 'Cliente:')
        c.setFont('Helvetica-Bold', 9)
        c.drawString(x + 15 * mm, self.cursor_y, cliente_display)
        self.cursor_y -= 12

        # RFC (valor en negritas). Si no viene en `datos`, intentar leer Clientes.json
        rfc = str(self.datos.get('rfc') or '')
        if not rfc:
            try:
                clientes_path = os.path.join(self.base_dir, 'data', 'Clientes.json')
                clientes_map = _cargar_clientes(clientes_path)
                rfc = (clientes_map.get(cliente.upper().strip(), {}) or {}).get('RFC', '') or ''
            except Exception:
                rfc = ''
        c.setFont('Helvetica', 9)
        c.drawString(x, self.cursor_y, 'RFC:')
        c.setFont('Helvetica-Bold', 9)
        c.drawString(x + 15 * mm, self.cursor_y, rfc)
        self.cursor_y -= 12

        # Fecha de emisión (ahora después de RFC)
        fecha_emision = str(self.datos.get('fecha_emision', '') or '')
        fecha_larga = _formato_fecha_larga(fecha_emision)
        # Construir la línea completa: 'Fecha de Emisión: jueves 8 de enero de 2026'
        combined = f"Fecha de Emisión: {fecha_larga}" if fecha_larga else 'Fecha de Emisión:'
        # Determinar ancho máximo disponible y ajustar tamaño si es necesario
        c.setFont('Helvetica-Bold', 9)
        max_w = right_x - (x + 15 * mm)
        text_w = c.stringWidth(combined, 'Helvetica-Bold', 9)
        if text_w > max_w:
            # reducir ligeramente la fuente para intentar que quepa en una línea
            c.setFont('Helvetica-Bold', 8)
            text_w = c.stringWidth(combined, 'Helvetica-Bold', 8)
        # Dibujar la línea completa alineada a la derecha en `right_x`
        c.drawRightString(right_x, self.cursor_y, combined)
        # dejar un espacio extra entre la fecha de emisión y el siguiente párrafo
        self.cursor_y -= 40

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
        c.setLineWidth(1)
        c.line(left, line_y, right, line_y)

        # Título centrado (más separación respecto a la línea superior)
        title_y = line_y - 14
        c.setFont('Helvetica-Bold', 11)
        c.drawCentredString(self.width / 2, title_y, 'Condiciones de la Constancia')

        # Preparar contenido
        self.cursor_y = title_y - 12
        producto = str(self.datos.get('producto', '') or '').strip()
        condiciones = [
            '1. Este documento sólo ampara la información contenida en el producto cuya etiqueta muestra se presenta en esta Constancia.',
            '2. Cualquier modificación a la etiqueta debe ser sometida a la consideración de la Unidad de Inspección Acreditada y Aprobada en los términos de la Ley de Infraestructura de la Calidad, para que inspeccione su cumplimiento con la Norma Oficial Mexicana aplicable.',
        ]

        # Dibujar las dos primeras condiciones con el método existente
        c.setFont('Helvetica', 9)
        max_cond_w = (right - left) - 8 * mm
        for cond in condiciones:
            self._dibujar_texto_justificado(c, left + 4 * mm, self.cursor_y, cond, max_cond_w, font_name='Helvetica', font_size=9, leading=11)
            self.cursor_y -= 4

        # Tercera condición: norma y nombre de la norma en negritas usando Paragraph (soporta <b>...</b>)
        try:

            norma_text = str(self.datos.get('norma', '')).strip()
            nombre_text = str(self.datos.get('nombre_norma', '')).strip()
            tercero_html = (
            "3. Esta Constancia sólo ampara el cumplimiento con la Norma Oficial Mexicana "
            f"<b>{norma_text}</b> (<b>{nombre_text}</b>) para el producto: <b>{producto}</b>."
            )

            style = ParagraphStyle('cond3', fontName='Helvetica', fontSize=9, leading=11, alignment=TA_JUSTIFY)
            p = Paragraph(tercero_html, style)
            avail_w = max_cond_w
            wrap_w, wrap_h = p.wrap(avail_w, self.cursor_y if self.cursor_y > 0 else 200 * mm)
            # drawOn uses the lower-left corner, por eso restamos wrap_h
            p.drawOn(c, left + 4 * mm, self.cursor_y - wrap_h)
            self.cursor_y = self.cursor_y - wrap_h - 4
        except Exception:
            # Fallback: si Paragraph falla, dibujar la línea sin negritas
            fallback = f"3. Esta Constancia sólo ampara el cumplimiento con la Norma Oficial Mexicana {self.datos.get('norma','')} ({self.datos.get('nombre_norma','')}) para el producto: {producto}."
            self._dibujar_texto_justificado(c, left + 4 * mm, self.cursor_y, fallback, max_cond_w, font_name='Helvetica', font_size=9, leading=11)
            self.cursor_y -= 4
        # ya se dibujaron las dos primeras condiciones arriba; no volver a imprimir.

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
        # Nuevo diseño: tabla autoajustable a contenido (columnas y alturas dinámicas)
        margin_x = 48 * mm
        left = margin_x
        right = self.width - margin_x
        total_w = right - left
        x = left

        # Título en caja completa
        c.setLineWidth(0.6)
        title_box_h = 8 * mm
        top_y = self.cursor_y
        c.rect(x, top_y - title_box_h, total_w, title_box_h, stroke=1, fill=0)
        c.setFont('Helvetica-Bold', 9)
        c.drawCentredString(self.width / 2, top_y - title_box_h / 2 + 2, 'RELACIÓN CORRESPONDIENTE')

        # Avanzar cursor debajo del título con pequeño espacio
        self.cursor_y = top_y - title_box_h - 4 * mm

        # Preparar datos
        filas = self.datos.get('tabla_relacion') or []
        headers = ['CODIGO', 'MEDIDAS', 'CONTENIDO NETO']

        # Medir ancho requerido por columna (sin wrapping)
        font_name = 'Helvetica'
        font_size = 8
        padding = 4 * mm
        col_need = [0, 0, 0]
        # incluir encabezados en la medición
        for idx, h in enumerate(headers):
            w = c.stringWidth(h, font_name, font_size)
            col_need[idx] = max(col_need[idx], w + padding)

        # medir celdas
        for row in filas:
            try:
                codigo = str(row.get('CODIGO') or row.get('codigo') or row.get('Codigo') or '')
                medidas = str(row.get('MEDIDAS') or row.get('medidas') or row.get('Medidas') or '')
                contenido = str(row.get('CONTENIDO') or row.get('CONTENIDO NETO') or row.get('CONTENIDO_NETO') or row.get('contenido') or '')
            except Exception:
                codigo = medidas = contenido = ''
            col_need[0] = max(col_need[0], c.stringWidth(codigo, font_name, font_size) + padding)
            col_need[1] = max(col_need[1], c.stringWidth(medidas, font_name, font_size) + padding)
            col_need[2] = max(col_need[2], c.stringWidth(contenido, font_name, font_size) + padding)

        sum_need = sum(col_need)
        if sum_need == 0:
            # fallback: dividir en tres columnas iguales
            col_w = [total_w / 3.0] * 3
        else:
            if sum_need <= total_w:
                # usar los anchos necesarios
                col_w = col_need
            else:
                # escalar proporcionalmente
                factor = total_w / sum_need
                col_w = [w * factor for w in col_need]

        # ahora calcular la altura por fila según wrapping
        header_h = 7 * mm
        leading = 9  # points
        row_heights = []
        rows_cells = []
        for row in filas:
            try:
                codigo = str(row.get('CODIGO') or row.get('codigo') or row.get('Codigo') or '')
                medidas = str(row.get('MEDIDAS') or row.get('medidas') or row.get('Medidas') or '')
                contenido = str(row.get('CONTENIDO') or row.get('CONTENIDO NETO') or row.get('CONTENIDO_NETO') or row.get('contenido') or '')
            except Exception:
                codigo = medidas = contenido = ''

            # dividir texto por columna con el ancho disponible
            lines0 = _dividir_texto(c, codigo, col_w[0], font_name=font_name, font_size=font_size)
            lines1 = _dividir_texto(c, medidas, col_w[1], font_name=font_name, font_size=font_size)
            lines2 = _dividir_texto(c, contenido, col_w[2], font_name=font_name, font_size=font_size)
            max_lines = max(len(lines0), len(lines1), len(lines2), 1)
            h = max_lines * leading + (4 * mm)
            row_heights.append(h)
            rows_cells.append((lines0, lines1, lines2))

        table_h = header_h + sum(row_heights)
        table_top = self.cursor_y

        # Dibujar caja externa
        c.rect(x, table_top - table_h, total_w, table_h, stroke=1, fill=0)

        # Dibujar líneas verticales
        cur_x = x
        acc = x
        for w in col_w[:-1]:
            acc += w
            c.line(acc, table_top, acc, table_top - table_h)

        # Línea separadora entre header y contenido
        c.setLineWidth(0.8)
        c.line(x, table_top - header_h, x + total_w, table_top - header_h)
        c.setLineWidth(0.6)


        # Encabezados
        c.setFont('Helvetica-Bold', font_size)
        cur_x = x
        for i, h in enumerate(headers):
            w = col_w[i]
            cx = cur_x + w / 2
            cy = table_top - (header_h / 2) - (font_size / 2)
            c.drawCentredString(cx, cy, h)
            cur_x += w


        # Dibujar filas de contenido
        y = table_top - header_h
        c.setFont(font_name, font_size)
        for ri, (lines0, lines1, lines2) in enumerate(rows_cells):
            y -= row_heights[ri]
            # dibujar línea horizontal que cierra la fila
            c.line(x, y, x + total_w, y)
            # dibujar cada celda
            cell_x = x
            for ci, lines in enumerate((lines0, lines1, lines2)):
                tx = cell_x + (3 * mm)
                ty = y + row_heights[ri] - leading - (3 * mm)
                for ln in lines:
                    try:
                        c.drawString(tx, ty, ln)
                    except Exception:
                        pass
                    ty += -leading
                cell_x += col_w[ci]

        # actualizar cursor_y al final de la tabla (agregar espacio extra antes de observaciones)
        self.cursor_y = table_top - table_h - (8 * mm)

    def dibujar_observaciones(self, c: canvas.Canvas) -> None:
        x = 25 * mm
        max_w = 165 * mm
        obs = 'OBSERVACIONES: EN CUMPLIMIENTO CON LOS PUNTOS 4.2.6 Y 4.2.7 DE LA NORMA LOS DATOS DE FECHA DE CONSUMO PREFERENTE Y LOTE SE ENCUENTRAN DECLARADOS EN EL ENVASE DEL PRODUCTO. ESTE PRODUCTO FUE INSPECCIONADO EN CUMPLIMIENTO BAJO LA FASE 2 DE LA NOM CON VIGENCIA AL 31 DE DICIEMBRE DE 2027 Y FASE 3 DE LA NOM CON ENTRADA EN VIGOR A PARTIR DEL 01 DE ENERO DEL 2028.'
        c.setFont('Helvetica', 8)
        self._dibujar_texto_justificado(c, x, self.cursor_y, obs, max_w, font_name='Helvetica', font_size=8, leading=10)
        self.cursor_y -= 30

    def dibujar_evidencia(self, c: canvas.Canvas) -> None:
            """Crea una página para pegar evidencia fotográfica (2x2 cajas).

            Esta implementación añade una única página con cuatro recuadros.
            """
            try:
                c.showPage()
            except Exception:
                pass
            self.cursor_y = self.height - 40
            try:
                self.dibujar_fondo(c)
            except Exception:
                pass

            evidencias = self.datos.get('evidencias_lista', []) or []

            # Si no hay evidencias, añadir placeholder ${IMAGEN} centrado
            if not evidencias:
                c.setFont('Helvetica-Bold', 14)
                c.drawCentredString(self.width / 2, self.cursor_y - 70, '${IMAGEN}')
                # dejar espacio y regresar para que las firmas se dibujen después
                self.cursor_y -= 80
                return

            # Mostrar cada evidencia en su propia hoja, centrada y a mayor tamaño
            margin_x = 25 * mm
            margin_y_top = 40 * mm
            margin_y_bottom = 40 * mm
            total = len(evidencias)
            for idx, path in enumerate(evidencias, start=1):
                if idx > 1:
                    try:
                        c.showPage()
                    except Exception:
                        pass
                    try:
                        self.dibujar_fondo(c)
                    except Exception:
                        pass

                # Dibujar imagen centrada y lo más grande posible dentro de márgenes
                try:
                    if path and os.path.exists(path):
                        im = ImageReader(path)
                        iw, ih = im.getSize()
                        max_w = self.width - 2 * margin_x
                        max_h = self.height - (margin_y_top + margin_y_bottom + 30 * mm)
                        scale = min(max_w / iw, max_h / ih, 1)
                        draw_w = iw * scale
                        draw_h = ih * scale
                        draw_x = (self.width - draw_w) / 2
                        # subir la imagen ligeramente para que no quede demasiado baja
                        draw_y = (self.height - margin_y_bottom - draw_h) / 2 + 10 * mm
                        c.drawImage(im, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask='auto')
                except Exception:
                    pass

            # ajustar cursor_y para que las firmas se dibujen en la siguiente página
            self.cursor_y = margin_y_bottom






    def dibujar_firma(self, c: canvas.Canvas) -> None:
        # Imprimir firmas en página(s) final(es) con diseño de dos columnas similar al Dictamen
        try:
            c.showPage()
        except Exception:
            pass
        self.cursor_y = self.height - 150
        try:
            self.dibujar_fondo(c)
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
        code1 = None
        code2 = None

        # Si la tabla_relacion especifica un código de firma en su primera fila,
        # usar esa entrada preferente para la firma izquierda (como hace el dictamen).
        try:
            tr_tmp = list(self.datos.get('tabla_relacion') or [])
            if tr_tmp:
                first_row_pref = tr_tmp[0]
                pref_code = (first_row_pref.get('FIRMA') or first_row_pref.get('firma') or '').strip()
                if pref_code:
                    pref_entry = firmas_map.get(pref_code) or firmas_map.get(pref_code.upper())
                    if pref_entry:
                        nombre1 = pref_entry.get('NOMBRE DE INSPECTOR') or pref_entry.get('nombre') or nombre1
                        img1 = pref_entry.get('IMAGEN') or pref_entry.get('imagen') or img1
                        code1 = pref_code or code1
        except Exception:
            pass

        # Si no hay nombres, intentar sacar de Firmas.json
        if not nombre1 or not nombre2:
            for k, v in (firmas_map or {}).items():
                n = v.get('NOMBRE DE INSPECTOR') or v.get('nombre') or v.get('NOMBRE') or ''
                if not nombre1 and 'Gabriel' in n:
                    nombre1 = nombre1 or n
                    img1 = v.get('IMAGEN') or v.get('imagen')
                    code1 = k
                if not nombre2 and ('Arturo' in n or 'AFLORES' in (k or '').upper()):
                    nombre2 = nombre2 or n
                    img2 = v.get('IMAGEN') or v.get('imagen')
                    code2 = k
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

        # Validar acreditación de firmas: si no están acreditadas para la norma requerida, no imprimir
        norma_req = str(self.datos.get('norma') or '').strip()
        def _find_code_by_name(name):
            if not name:
                return None
            for kk, vv in (firmas_map or {}).items():
                n = vv.get('NOMBRE DE INSPECTOR') or vv.get('nombre') or vv.get('NOMBRE') or ''
                if n and n.strip().upper() == name.strip().upper():
                    return kk
            return None

        # Try to resolve codes if missing
        if nombre1 and not code1:
            code1 = _find_code_by_name(nombre1)
        if nombre2 and not code2:
            code2 = _find_code_by_name(nombre2)

        # Use plantillaPDF validator if available
        try:
            if code1:
                if _validar_acreditacion_inspector:
                    _, img_from_map, ok = _validar_acreditacion_inspector(code1, norma_req, firmas_map)
                    if not ok:
                        nombre1 = ''
                        img1 = None
                    else:
                        img1 = img1 or img_from_map
                else:
                    # fallback: check 'normas_acreditadas' field manually
                    inspector = firmas_map.get(code1, {})
                    normas_ac = inspector.get('Normas acreditadas') or inspector.get('normas_acreditadas') or inspector.get('Normas') or []
                    if norma_req and normas_ac and (norma_req in normas_ac or any(norma_req in na for na in normas_ac)):
                        img1 = img1 or inspector.get('IMAGEN') or inspector.get('imagen')
                    else:
                        nombre1 = ''
                        img1 = None
        except Exception:
            pass
        try:
            if code2:
                if _validar_acreditacion_inspector:
                    _, img_from_map2, ok2 = _validar_acreditacion_inspector(code2, norma_req, firmas_map)
                    if not ok2:
                        nombre2 = ''
                        img2 = None
                    else:
                        img2 = img2 or img_from_map2
                else:
                    inspector2 = firmas_map.get(code2, {})
                    normas_ac2 = inspector2.get('Normas acreditadas') or inspector2.get('normas_acreditadas') or inspector2.get('Normas') or []
                    if norma_req and normas_ac2 and (norma_req in normas_ac2 or any(norma_req in na for na in normas_ac2)):
                        img2 = img2 or inspector2.get('IMAGEN') or inspector2.get('imagen')
                    else:
                        nombre2 = ''
                        img2 = None
        except Exception:
            pass

        # Column coordinates
        left_x = 25 * mm
        right_x = self.width / 2 + 10 * mm

        

        # Draw images if available
        # reducir ligeramente el tamaño de las firmas para que queden más armoniosas
        sig_h = 22 * mm
        sig_w1 = 50 * mm
        sig_w2 = 50 * mm

        # Left signature
        # Si en la tabla_relacion hay un código de firma preferido, usarlo (coincide con el armado de dictamen)
        try:
            tr = list(self.datos.get('tabla_relacion') or [])
            if tr:
                first_row = tr[0]
                pref_code = (first_row.get('FIRMA') or first_row.get('firma') or '').strip()
                if pref_code:
                    # buscar en firmas_map
                    pref_entry = firmas_map.get(pref_code) or firmas_map.get(pref_code.upper())
                    if pref_entry:
                        nombre1 = pref_entry.get('NOMBRE DE INSPECTOR') or pref_entry.get('nombre') or nombre1
                        img1 = img1 or pref_entry.get('IMAGEN') or pref_entry.get('imagen')
                        code1 = pref_code
        except Exception:
            pass

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
        # Si no se especifica salida, no guardar PDFs en data/Constancias por defecto.
        # Usar un archivo temporal para evitar congestionar la carpeta `data/Constancias`.
        if not salida:
            fol = str(self.datos.get('folio_constancia') or '')
            safe = fol.replace('/', '_').replace(' ', '_') or datetime.now().strftime('%Y%m%d_%H%M%S')
            tmp_dir = tempfile.gettempdir()
            salida = os.path.join(tmp_dir, f'Constancia_{safe}.pdf')

        # Use NumberedCanvas so we can print "Página X de Y" after all pages are created
        try:
            c = NumberedCanvas(salida, pagesize=letter)
        except Exception:
            c = canvas.Canvas(salida, pagesize=letter)
        # Asegurar que el canvas tenga título/cadena para que se impriman en todas las páginas
        try:
            # construir cadena identificadora y su versión para mostrar (2 dígitos)
            cadena = self.datos.get('cadena') or ''
            if not cadena:
                try:
                    cadena = self.construir_cadena_identificacion() or ''
                except Exception:
                    cadena = ''
            # Preparar la cadena para mostrar en encabezado: forzar año en 2 dígitos
            try:
                display_cadena = re.sub(r"(\d{4})(?=049)", lambda m: m.group(1)[-2:], cadena)
            except Exception:
                display_cadena = cadena
            # asignar valores al canvas para que NumberedCanvas los dibuje en showPage
            try:
                c.header_title = 'CONSTANCIA DE CONFORMIDAD'
                c.header_chain = display_cadena
                # Dibujar encabezado inmediatamente en la primera página (titulo arriba, cadena debajo)
                try:
                    c.setFont('Helvetica-Bold', 20)
                    c.drawCentredString(self.width / 2, self.height - 58, c.header_title)
                except Exception:
                    pass
                try:
                    c.setFont('Helvetica', 8)
                    c.drawCentredString(self.width / 2, self.height - 74, c.header_chain)
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass
        # Reservar más espacio vertical para contenido, dejando margen bajo el encabezado
        self.cursor_y = self.height - 120
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
        self.dibujar_datos_basicos(c)
        self.dibujar_cuerpo_legal(c)
        self.dibujar_condiciones(c)

        # Tabla de relación inmediatamente después de las condiciones
        self.dibujar_tabla_relacion(c)

        
        self.dibujar_observaciones(c)

        # Añadir apartado para pegar evidencia fotográfica (páginas nuevas)
        try:
            # Intentar localizar evidencias automáticamente usando data/evidence_paths.json
            try:
                evidencia_cfg = {}
                cfg_path = os.path.join(self.base_dir, 'data', 'evidence_paths.json')
                if os.path.exists(cfg_path):
                    with open(cfg_path, 'r', encoding='utf-8') as ef:
                        evidencia_cfg = json.load(ef) or {}

                IMG_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}
                def _normalizar(s):
                    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

                # construir índice simple: clave_normalizada -> [paths]
                indice = {}
                for grp, lst in (evidencia_cfg or {}).items():
                    if not isinstance(lst, list):
                        continue
                    for carpeta in lst:
                        try:
                            # permitir rutas relativas guardadas en la configuración
                            carpeta_path = carpeta
                            try:
                                if not os.path.isabs(carpeta_path):
                                    carpeta_path = os.path.join(self.base_dir, carpeta_path)
                            except Exception:
                                carpeta_path = carpeta
                            if not os.path.exists(carpeta_path):
                                continue
                            for root, _, files in os.walk(carpeta_path):
                                for nombre in files:
                                    base, ext = os.path.splitext(nombre)
                                    if ext.lower() not in IMG_EXTS:
                                        continue
                                    path = os.path.join(root, nombre)
                                    # extraer core y normalizar
                                    try:
                                        core = re.sub(r"[\s\-_]*\(\s*\d+\s*\)$", "", base)
                                        core = re.sub(r"[\s\-_]+\d+$", "", core)
                                    except Exception:
                                        core = base
                                    key = _normalizar(core)
                                    if not key:
                                        continue
                                    indice.setdefault(key, []).append(path)
                                    # indexar también por nombre de carpeta padre
                                    try:
                                        parent = os.path.basename(root or "")
                                        parent_core = re.sub(r"[\s\-_]*\(\s*\d+\s*\)$", "", parent)
                                        parent_core = re.sub(r"[\s\-_]+\d+$", "", parent_core)
                                        parent_key = _normalizar(parent_core)
                                        if parent_key and parent_key != key:
                                            indice.setdefault(parent_key, []).append(path)
                                    except Exception:
                                        pass
                        except Exception:
                            continue

                # claves a buscar: folio, solicitud, cliente y claves de tabla_relacion (CODIGO, DESCRIPCION, MARCA)
                buscar = []
                fol = str(self.datos.get('folio_constancia') or '')
                if fol:
                    buscar.append(_normalizar(fol))
                sol = str(self.datos.get('solicitud_formateado') or self.datos.get('solicitud') or '')
                if sol:
                    buscar.append(_normalizar(sol))
                cliente = str(self.datos.get('cliente') or '')
                if cliente:
                    buscar.append(_normalizar(cliente))
                # añadir claves desde la tabla de relación para mejorar matching (como en dictamen)
                try:
                    tr = list(self.datos.get('tabla_relacion') or [])
                    for row in tr:
                        try:
                            codigo = str(row.get('CODIGO') or row.get('codigo') or '')
                            desc = str(row.get('DESCRIPCION') or row.get('descripcion') or row.get('Contenido') or row.get('CONTENIDO') or '')
                            marca = str(row.get('MARCA') or row.get('marca') or '')
                            if codigo:
                                buscar.append(_normalizar(codigo))
                            if desc:
                                buscar.append(_normalizar(desc))
                            if marca:
                                buscar.append(_normalizar(marca))
                        except Exception:
                            continue
                except Exception:
                    pass

                encontrados = []
                for k in buscar:
                    if not k:
                        continue
                    for ik, paths in indice.items():
                        if k in ik or ik in k:
                            for p in paths:
                                if p not in encontrados and os.path.exists(p):
                                    encontrados.append(p)

                # conservar evidencias ya provistas en datos, y anteponer las encontradas
                prov = list(self.datos.get('evidencias_lista') or [])
                final_list = encontrados + [p for p in prov if p not in encontrados]
                if final_list:
                    self.datos['evidencias_lista'] = final_list
            except Exception:
                pass
            self.dibujar_evidencia(c)
        except Exception:
            pass

        # Firmas al final del documento
        try:
            self.dibujar_firma(c)
        except Exception:
            pass
        c.save()
        return salida

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
        # siempre crear un respaldo timestamped en data/tabla_relacion_backups
        try:
            backups_dir = os.path.join(os.path.dirname(path), 'tabla_relacion_backups')
            os.makedirs(backups_dir, exist_ok=True)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_name = f"tabla_de_relacion_before_update_{ts}.json"
            shutil.copy(path, os.path.join(backups_dir, backup_name))
        except Exception:
            pass

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
                # Además crear un respaldo con timestamp en data/tabla_relacion_backups
                try:
                    backups_dir = os.path.join(os.path.dirname(path), 'tabla_relacion_backups')
                    os.makedirs(backups_dir, exist_ok=True)
                    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                    backup_name = f"tabla_de_relacion_{ts}.json"
                    with open(os.path.join(backups_dir, backup_name), 'w', encoding='utf-8') as bf:
                        json.dump(data, bf, ensure_ascii=False, indent=2)
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
            # Intentar cargar por medio de plantillaPDF.cargar_tabla_relacion (retorna DataFrame)
            if _cargar_tabla_relacion_ext:
                try:
                    df = _cargar_tabla_relacion_ext(tabla)
                    if not df.empty:
                        first = df.iloc[0].to_dict()
                        producto = first.get('DESCRIPCION','')
                        marca = first.get('MARCA','')
                        modelo = first.get('MODELO','')
                except Exception:
                    pass
            # Fallback: leer JSON bruto
            if not producto:
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
    nombre_norma = ''
    if visita.get('norma'):
        norma_raw = str(visita.get('norma')).split(',')[0].strip()
        # si viene numérico como '4', intentar mapear a NOM-004-... en Normas.json
        if norma_raw.isdigit():
            target = f"{int(norma_raw):03d}"
            try:
                # buscar entry en el archivo Normas.json que contenga el número
                if os.path.exists(normas_p):
                    with open(normas_p, 'r', encoding='utf-8') as nf:
                        ndata = json.load(nf)
                        if isinstance(ndata, list):
                            for item in ndata:
                                nom = str(item.get('NOM') or '')
                                if f"-{target}-" in nom or f"{target}" in nom:
                                    norma_str = nom
                                    nombre_norma = item.get('NOMBRE') or ''
                                    break
            except Exception:
                pass
        # si no se resolvió como número, tomar como código directo
        if not norma_str:
            norma_str = norma_raw
            nombre_norma = normas.get(norma_str, '')

    cliente = visita.get('cliente','')
    rfc = (clientes.get(cliente.upper(), {}) or {}).get('RFC','')
    # Número de contrato (desde Clientes.json campo "NÚMERO_DE_CONTRATO")
    no_contrato = (clientes.get(cliente.upper(), {}) or {}).get('NÚMERO_DE_CONTRATO', '')
    # Fecha de contrato (desde Clientes.json campo "FECHA_DE_CONTRATO")
    fecha_contrato = (clientes.get(cliente.upper(), {}) or {}).get('FECHA_DE_CONTRATO', '')

    fecha = visita.get('fecha_termino') or visita.get('fecha') or datetime.now().strftime('%d/%m/%Y')

    # Preferir folio de dictamen/familia si está presente; si no, usar folio_visita
    fol = (visita.get('folio') or visita.get('folio_visita') or '').replace('UDC', 'UCC')
    # Capturar solicitud de la visita en múltiples formatos posibles (ej. "000333/25")
    solicitud_raw = str(visita.get('solicitud') or visita.get('Solicitud') or visita.get('SOLICITUD') or '').strip()
    solicitud_num = ''
    solicitud_year_full = None
    if solicitud_raw:
        m = re.match(r"^\s*(\d+)(?:[\/-](\d{2,4}))?\s*$", solicitud_raw)
        if m:
            solicitud_num = m.group(1)
            suf = m.group(2)
            if suf:
                solicitud_year_full = ("20" + suf) if len(suf) == 2 else suf
        else:
            nums = re.findall(r"\d+", solicitud_raw)
            solicitud_num = nums[0] if nums else solicitud_raw
    # preparar solicitud_formateado (6 dígitos cuando es numérica)
    solicitud_formateado = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num
    # Cargar tabla_de_relacion y seleccionar filas relacionadas (CODIGO, MEDIDAS, CONTENIDO)
    tabla_rows = []
    if os.path.exists(tabla):
        try:
            with open(tabla, 'r', encoding='utf-8') as f:
                t = json.load(f)
                if isinstance(t, list) and t:
                    # Normalizar búsqueda
                    fol_clean = str(fol).strip()
                    solicitud_vis = str(visita.get('solicitud','')).strip()
                    cliente_upper = cliente.upper().strip()
                    for row in t:
                        if not isinstance(row, dict):
                            continue
                        if str(row.get('FOLIO','')).strip() == fol_clean:
                            tabla_rows.append(row)
                            continue
                        if solicitud_vis and str(row.get('SOLICITUD','')).strip() == solicitud_vis:
                            tabla_rows.append(row)
                            continue
                        # tratar coincidencias por cliente/marca
                        if str(row.get('CLIENTE','')).strip().upper() == cliente_upper or str(row.get('MARCA','')).strip().upper() == cliente_upper:
                            tabla_rows.append(row)
                    # si no encontramos coincidencias, tomar los primeros 4 como ejemplo
                    if not tabla_rows:
                        tabla_rows = t[:4]
        except Exception:
            tabla_rows = []

    # Crear copia de respaldo completa de la tabla de relación en data/tabla_relacion_backups
    try:
        if os.path.exists(tabla):
            backups_dir2 = os.path.join(data_dir, 'tabla_relacion_backups')
            os.makedirs(backups_dir2, exist_ok=True)
            ts2 = datetime.now().strftime('%Y%m%d_%H%M%S')
            shutil.copy(tabla, os.path.join(backups_dir2, f"tabla_de_relacion_{ts2}.json"))
    except Exception:
        pass

    # Si no existía el archivo fuente o no se pudo copiar, guardar al menos
    # el extracto `tabla_rows` en el mismo directorio de backups para trazabilidad.
    try:
        if tabla_rows:
            backups_dir3 = os.path.join(data_dir, 'tabla_relacion_backups')
            os.makedirs(backups_dir3, exist_ok=True)
            ts3 = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_fol = str(fol).replace('/', '_') or 'nofolio'
            out_name = f"tabla_relacion_extract_{safe_fol}_{ts3}.json"
            with open(os.path.join(backups_dir3, out_name), 'w', encoding='utf-8') as bf:
                json.dump(tabla_rows, bf, ensure_ascii=False, indent=2)
    except Exception:
        pass

    # Si la visita no incluye campo 'solicitud', intentar obtenerlo desde la primera fila
    # de la `tabla_relacion` (muchas veces la solicitud viene en esa columna).
    if (not solicitud_num or not solicitud_raw) and tabla_rows:
        try:
            first = tabla_rows[0]
            s = str(first.get('SOLICITUD') or first.get('Solicitud') or first.get('solicitud') or '').strip()
            if s:
                solicitud_raw = s
                m2 = re.match(r"^\s*(\d+)(?:[\/-](\d{2,4}))?\s*$", solicitud_raw)
                if m2:
                    solicitud_num = m2.group(1)
                    suf2 = m2.group(2)
                    if suf2:
                        solicitud_year_full = ("20" + suf2) if len(suf2) == 2 else suf2
                else:
                    nums2 = re.findall(r"\d+", solicitud_raw)
                    solicitud_num = nums2[0] if nums2 else solicitud_raw
                solicitud_formateado = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num
        except Exception:
            pass

    datos = {
        'folio_constancia': fol,
        'fecha_emision': fecha,
        'cliente': cliente,
        'rfc': rfc,
        'no_contrato': no_contrato,
        'fecha_contrato': fecha_contrato,
        'solicitud': solicitud_raw,
        'solicitud_formateado': solicitud_formateado,
        'norma': norma_str,
        'normades': nombre_norma,
        'nombre_norma': nombre_norma,
        'producto': producto,
        'marca': marca,
        'modelo': modelo,
        'tabla_relacion': tabla_rows,
    }

    if not salida:
        tmp_dir = tempfile.gettempdir()
        safe = str(fol or 'constancia').replace('/', '_').replace(' ', '_')
        salida = os.path.join(tmp_dir, f'Constancia_{safe}.pdf')

    gen = ConstanciaPDFGenerator(datos, base_dir=base)
    return gen.generar(salida)

def generar_json_constancias_desde_historial(salida_dir: str | None = None, max_items: int | None = None) -> list:
    """Lee `data/historial_visitas.json` y genera un JSON con los datos de constancia
    para cada visita encontrada. Guarda los JSON en `data/Constancias` o en `salida_dir`.

    Devuelve la lista de rutas de los JSON creados.
    """
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
        return []

    clientes = _cargar_clientes(clientes_p)
    normas = _cargar_normas(normas_p)

    # actualizar tabla de relacion si es necesario
    _actualizar_tabla_relacion(tabla)

    # cargar primer registro de tabla_de_relacion si existe (usado como ejemplo)
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

    out_dir = salida_dir or os.path.join(data_dir, 'Constancias')
    os.makedirs(out_dir, exist_ok=True)
    created = []
    count = 0
    for v in visitas:
        if max_items is not None and count >= max_items:
            break
        try:
            norma_str = ''
            if v.get('norma'):
                norma_str = v.get('norma').split(',')[0].strip()
            nombre_norma = normas.get(norma_str, '')

            cliente = v.get('cliente','')
            rfc = (clientes.get(cliente.upper(), {}) or {}).get('RFC','')

            fecha = v.get('fecha_termino') or v.get('fecha') or datetime.now().strftime('%d/%m/%Y')

            # Número y fecha de contrato desde Clientes.json
            no_contrato = (clientes.get(cliente.upper(), {}) or {}).get('NÚMERO_DE_CONTRATO', '')
            fecha_contrato = (clientes.get(cliente.upper(), {}) or {}).get('FECHA_DE_CONTRATO', '')

            fol = (v.get('folio_visita') or v.get('folio') or '')
            safe_fol = str(fol).replace('/','_').replace(' ', '_') or f'visita_{count+1}'

            datos = {
                'folio_constancia': fol,
                'fecha_emision': fecha,
                'cliente': cliente,
                'rfc': rfc,
                'no_contrato': no_contrato,
                'fecha_contrato': fecha_contrato,
                'norma': norma_str,
                'nombre_norma': nombre_norma,
                'producto': producto,
                'marca': marca,
                'modelo': modelo,
                'origen_visita': v,
            }

            out_path = os.path.join(out_dir, f'Constancia_{safe_fol}.json')
            with open(out_path, 'w', encoding='utf-8') as jf:
                json.dump(datos, jf, ensure_ascii=False, indent=2)
            created.append(out_path)
            count += 1
        except Exception:
            continue

    return created

if __name__ == '__main__':
    # Generador integrado de ejemplos: crea JSON en data/Constancias y genera PDFs
    def generar_ejemplos_integrados(count: int = 3) -> None:
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        data_dir = os.path.join(base, 'data')
        const_dir = os.path.join(data_dir, 'Constancias')
        os.makedirs(const_dir, exist_ok=True)

        for i in range(1, count + 1):
            fecha = (datetime.now()).strftime('%d/%m/%Y')
            folio = f'UCC-DEMO-{i:04d}'
            datos = {
                'folio_constancia': folio,
                'fecha_emision': fecha,
                'cliente': f'CLIENTE DEMO {i}',
                'rfc': f'XAXX01010{i:03d}',
                'norma': f'NOM-00{i}-XXXX',
                'nombre_norma': f'Nombre de la norma demo {i}',
                'producto': f'PRODUCTO DEMO {i}',
                'marca': f'MARCA DEMO {i}',
                'modelo': f'MODELO DEMO {i}',
            }
            # Guardar JSON
            json_path = os.path.join(const_dir, f'constancia_demo_{i}.json')
            try:
                with open(json_path, 'w', encoding='utf-8') as jf:
                    json.dump(datos, jf, ensure_ascii=False, indent=2)
                print('Wrote', json_path)
            except Exception as e:
                print('Error writing JSON', json_path, e)

            # Generar PDF (se guarda en data/Constancias por defecto)
            try:
                gen = ConstanciaPDFGenerator(datos, base_dir=base)
                out = gen.generar(None)
                print('Generated PDF:', out)
            except Exception as e:
                print('Error generating PDF for', json_path, e)

    # Ejecutar generador integrado al invocar este script
    generar_ejemplos_integrados(3)
