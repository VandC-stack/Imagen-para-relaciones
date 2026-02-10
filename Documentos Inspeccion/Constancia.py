"""Plantilla: Constancia de Conformidad

Esta plantilla genera una constancia en PDF usando img/Fondo.jpg como fondo
y carga datos desde data/Clientes.json y data/Normas.json. También ofrece una
función para leer tabla_de_relacion.json y actualizar "TIPO DE DOCUMENTO" D->C.
"""
        # Cadena identificadora (cadena del dictamen/constancia) - centrada bajo el título
import os
import sys
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
import time


# Determinar rutas base y data de manera consistente entre ejecución python y .exe
try:
    if getattr(sys, 'frozen', False):
        PACKAGE_BASE = getattr(sys, '_MEIPASS', os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    else:
        PACKAGE_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
except Exception:
    PACKAGE_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Ruta de datos persistentes (la misma que `app.py` exporta en FOLIO_DATA_DIR)
def get_data_dir():
    """Retorna la carpeta de datos persistente en tiempo de ejecución.

    Prioriza la variable de entorno `FOLIO_DATA_DIR` establecida por `app.py`
    cuando la aplicación está empacada (.exe). Si no existe, cae al
    `PACKAGE_BASE/data` junto al bundle.
    """
    try:
        return os.getenv('FOLIO_DATA_DIR') or os.path.join(PACKAGE_BASE, 'data')
    except Exception:
        return os.path.join(PACKAGE_BASE, 'data')

DATA_DIR = get_data_dir()


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
        # Año a partir de fecha_emision si está, else año actual (AAAA)
        fecha = self.datos.get('fecha_emision', '')
        year_full = None
        if fecha:
            m = re.search(r"(\d{4})", fecha)
            if m:
                year_full = m.group(1)
            else:
                m2 = re.search(r"(\d{2})/(\d{2})/(\d{4})", fecha)
                if m2:
                    year_full = m2.group(3)
        if not year_full:
            year_full = datetime.now().strftime('%Y')

        # conservar los guiones en la norma (ej. NOM-051-SCFI/SSA1-2010)
        norma = (self.datos.get('norma') or '').strip()
        # Si no viene la norma completa, pero tenemos 'clasif_uva' intentar mapearla usando data/Normas.json
        try:
            if (not norma) and self.datos.get('clasif_uva'):
                clas = str(self.datos.get('clasif_uva'))
                mnum = re.search(r"(\d{1,3})", clas)
                if mnum:
                    code3 = f"{int(mnum.group(1)):03d}"
                    normas_path = os.path.join(DATA_DIR, 'Normas.json')
                    if os.path.exists(normas_path):
                        with open(normas_path, 'r', encoding='utf-8') as nf:
                            ndata = json.load(nf)
                            if isinstance(ndata, list):
                                for item in ndata:
                                    nom = str(item.get('NOM') or '')
                                    if f"-{code3}-" in nom or code3 in nom:
                                        norma = nom
                                        # also set nombre_norma in datos if available
                                        nombre = item.get('NOMBRE') or item.get('NOMBRE', '')
                                        if nombre:
                                            self.datos['nombre_norma'] = nombre
                                        break
        except Exception:
            pass

        # Prioridad para determinar folio_formateado (más robusta):
        # 1) archivo per-visit en data/folios_visitas/folios_{folio_visita}.json (si existe)
        # 2) `data/folio_counter.json` -> campo 'last' (preferir si es >= valor encontrado)
        # 3) `self.datos['folio_formateado']` si ya fue establecido
        # 4) extraer dígitos de `folio_constancia` o `folio`
        folio_formateado = None

        # 1) per-visit saved folio (re-check at generation time to catch recent writes)
        try:
            fid = str(self.datos.get('folio_visita') or self.datos.get('folio') or self.datos.get('folio_constancia') or '').strip()
            if fid:
                archivo_f = os.path.join(DATA_DIR, 'folios_visitas', f"folios_{fid}.json")
                if os.path.exists(archivo_f):
                    with open(archivo_f, 'r', encoding='utf-8') as ff:
                        obj = json.load(ff) or {}
                    fols_list = obj.get('folios') if isinstance(obj, dict) else obj
                    if isinstance(fols_list, list) and fols_list:
                        # Try to find the folio that matches the current 'lista' in datos.
                        preferred = None
                        lista_pref = str(self.datos.get('lista') or '').strip()
                        for ent in fols_list:
                            try:
                                if not ent:
                                    continue
                                ent_lista = str(ent.get('LISTA') or ent.get('lista') or '').strip()
                                if lista_pref and ent_lista and ent_lista == lista_pref:
                                    preferred = ent
                                    break
                            except Exception:
                                continue
                        if not preferred:
                            # fallback: try to match by CODIGO or MARCA if available
                            codigo_pref = ''
                            try:
                                tr = self.datos.get('tabla_relacion') or []
                                if isinstance(tr, list) and tr:
                                    codigo_pref = str((tr[0] or {}).get('CODIGO') or '').strip()
                            except Exception:
                                codigo_pref = ''
                            if codigo_pref:
                                for ent in fols_list:
                                    try:
                                        if str(ent.get('CODIGO') or '').strip() == codigo_pref:
                                            preferred = ent
                                            break
                                    except Exception:
                                        continue
                        if not preferred:
                            preferred = fols_list[0]
                        first = preferred or fols_list[0] or {}
                        fval = first.get('FOLIOS') or first.get('FOLIO') or ''
                        fd = ''.join([c for c in str(fval) if c.isdigit()])
                        if fd:
                            folio_formateado = fd.zfill(6)
                            self.datos['folio_formateado'] = folio_formateado
        except Exception:
            pass

        # 2) folio_counter.json: preferir si es >= al valor encontrado
        try:
            fc_path = os.path.join(DATA_DIR, 'folio_counter.json')
            if os.path.exists(fc_path):
                with open(fc_path, 'r', encoding='utf-8') as fcf:
                    j = json.load(fcf) or {}
                last = j.get('last')
                if last is not None:
                    try:
                        last_int = int(last)
                        if folio_formateado:
                            try:
                                curr_int = int(str(folio_formateado).lstrip('0') or folio_formateado)
                            except Exception:
                                curr_int = None
                            if curr_int is None or last_int >= curr_int:
                                folio_formateado = str(last_int).zfill(6)
                                self.datos['folio_formateado'] = folio_formateado
                        else:
                            folio_formateado = str(last_int).zfill(6)
                            self.datos['folio_formateado'] = folio_formateado
                    except Exception:
                        pass
        except Exception:
            pass

        # 3) existing self.datos value (if still not set)
        if not folio_formateado:
            try:
                f_existing = str(self.datos.get('folio_formateado') or '').strip()
                if f_existing:
                    folio_formateado = f_existing
            except Exception:
                folio_formateado = None

        # 4) fallback: extraer de folio_constancia o folio
        if not folio_formateado:
            folio_const = str(self.datos.get('folio_constancia','') or '').strip()
            if folio_const:
                nums = re.findall(r"\d+", folio_const)
                digits = nums[-1] if nums else ''
                if digits:
                    folio_formateado = digits.zfill(6)
                else:
                    folio_formateado = folio_const
            else:
                folio = str(self.datos.get('folio') or '')
                nums = re.findall(r"\d+", folio)
                digits = nums[-1] if nums else ''
                if digits:
                    folio_formateado = digits.zfill(6)
                else:
                    folio_formateado = folio
            if folio_formateado:
                self.datos['folio_formateado'] = folio_formateado

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

        # determinar año de emisión corto (2 dígitos) para la parte izquierda: usar año actual
        year_left_two = datetime.now().strftime('%y')

        # determinar año de la parte 'Solicitud de Servicio' (2 dígitos): preferir el año extraído de la solicitud
        if solicitud_year_two:
            sol_year_two = solicitud_year_two
        else:
            # extraer dos últimos dígitos del `solicitud` si posible, si no usar el año actual
            if solicitud_raw:
                msol = re.search(r"/(\d{2,4})\s*$", solicitud_raw)
                if msol:
                    sfull = msol.group(1)
                    sol_year_two = sfull[-2:]
                else:
                    sol_year_two = datetime.now().strftime('%y')
            else:
                sol_year_two = datetime.now().strftime('%y')

        # asegurar formato de solicitud (6 dígitos cuando sea numérica)
        solicitud_formatted = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num

        cadena = f"{year_left_two}049UCC{norma}{folio_formateado} Solicitud de Servicio: {sol_year_two}049UCC{norma}{solicitud_formatted}-{lista}"
        self.datos['cadena'] = cadena
        return cadena

    def _refresh_folio_from_disk(self) -> None:
        """Re-lee `data/folio_counter.json` y `data/folios_visitas/folios_{id}.json`
        para asegurar que `self.datos['folio_formateado']` refleja lo que hay en disco.
        Busca la entrada que coincida con `self.datos['lista']`, `folio_tabla` o `codigo`.
        """
        try:
            # Leer desde DATA_DIR persistente (AppData cuando está empaquetado)
            data_dir = DATA_DIR
            # Primero, si hay archivo per-visit, intentar usarlo (y buscar por lista)
            fid = str(self.datos.get('folio_visita') or self.datos.get('folio') or self.datos.get('folio_constancia') or '').strip()
            chosen = None
            if fid:
                archivo_f = os.path.join(data_dir, 'folios_visitas', f"folios_{fid}.json")
                if os.path.exists(archivo_f):
                    try:
                        with open(archivo_f, 'r', encoding='utf-8') as ff:
                            obj = json.load(ff) or {}
                        fols_list = obj.get('folios') if isinstance(obj, dict) else obj
                        if isinstance(fols_list, list) and fols_list:
                            lista_pref = str(self.datos.get('lista') or '').strip()
                            for ent in fols_list:
                                try:
                                    if not ent:
                                        continue
                                    ent_lista = str(ent.get('LISTA') or ent.get('lista') or '').strip()
                                    if lista_pref and ent_lista and ent_lista == lista_pref:
                                        chosen = ent
                                        break
                                except Exception:
                                    continue
                            if not chosen:
                                # fallback matches
                                folio_tabla_pref = str(self.datos.get('folio_tabla') or '').strip()
                                codigo_pref = str(self.datos.get('codigo') or '').strip()
                                if folio_tabla_pref:
                                    for ent in fols_list:
                                        try:
                                            if str(ent.get('FOLIO') or ent.get('FOLIOS') or '').strip() == folio_tabla_pref:
                                                chosen = ent
                                                break
                                        except Exception:
                                            continue
                                if not chosen and codigo_pref:
                                    for ent in fols_list:
                                        try:
                                            if str(ent.get('CODIGO') or '').strip() == codigo_pref:
                                                chosen = ent
                                                break
                                        except Exception:
                                            continue
                            if not chosen:
                                chosen = fols_list[0]
                            try:
                                fval = chosen.get('FOLIOS') or chosen.get('FOLIO') or ''
                                fd = ''.join([c for c in str(fval) if c.isdigit()])
                                if fd:
                                    self.datos['folio_formateado'] = fd.zfill(6)
                                    return
                            except Exception:
                                pass
                    except Exception:
                        pass

            # Si no hallamos por visita, leer folio_counter.json
            try:
                fc_path = os.path.join(data_dir, 'folio_counter.json')
                if os.path.exists(fc_path):
                    with open(fc_path, 'r', encoding='utf-8') as fcf:
                        j = json.load(fcf) or {}
                    last = j.get('last')
                    if last is not None:
                        try:
                            last_int = int(last)
                            self.datos['folio_formateado'] = str(last_int).zfill(6)
                            return
                        except Exception:
                            pass
            except Exception:
                pass
        except Exception:
            pass

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
                    clientes_path = os.path.join(get_data_dir(), 'Clientes.json')
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
                clientes_path = os.path.join(get_data_dir(), 'Clientes.json')
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
        marca = str(self.datos.get('marca', '') or '').strip()
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
            f"<b>{norma_text}</b> (<b>{nombre_text}</b>) para el producto: <b>{producto}</b> - Marca: <b>{marca}</b>."
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
        # tabla autoajustable a contenido (columnas y alturas dinámicas)
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
        headers = ['MARCA','CODIGO', 'MEDIDAS', 'CONTENIDO NETO']

        # Medir ancho requerido por columna (sin wrapping)
        font_name = 'Helvetica'
        font_size = 8
        padding = 4 * mm
        col_need = [0, 0, 0, 0]
        # incluir encabezados en la medición
        for idx, h in enumerate(headers):
            w = c.stringWidth(h, font_name, font_size)
            col_need[idx] = max(col_need[idx], w + padding)

        # medir celdas (ahora incluyendo MARCA)
        for row in filas:
            try:
                marca = str(row.get('MARCA') or row.get('marca') or row.get('Marca') or '')
                codigo = str(row.get('CODIGO') or row.get('codigo') or row.get('Codigo') or '')
                medidas = str(row.get('MEDIDAS') or row.get('medidas') or row.get('Medidas') or '')
                contenido = str(row.get('CONTENIDO') or row.get('CONTENIDO NETO') or row.get('CONTENIDO_NETO') or row.get('contenido') or '')
            except Exception:
                marca = codigo = medidas = contenido = ''
            col_need[0] = max(col_need[0], c.stringWidth(marca, font_name, font_size) + padding)
            col_need[1] = max(col_need[1], c.stringWidth(codigo, font_name, font_size) + padding)
            col_need[2] = max(col_need[2], c.stringWidth(medidas, font_name, font_size) + padding)
            col_need[3] = max(col_need[3], c.stringWidth(contenido, font_name, font_size) + padding)

        sum_need = sum(col_need)
        if sum_need == 0:
            # fallback: dividir en cuatro columnas iguales
            col_w = [total_w / 4.0] * 4
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
                marca = str(row.get('MARCA') or row.get('marca') or row.get('Marca') or '')
                codigo = str(row.get('CODIGO') or row.get('codigo') or row.get('Codigo') or '')
                medidas = str(row.get('MEDIDAS') or row.get('medidas') or row.get('Medidas') or '')
                contenido = str(row.get('CONTENIDO') or row.get('CONTENIDO NETO') or row.get('CONTENIDO_NETO') or row.get('contenido') or '')
            except Exception:
                marca = codigo = medidas = contenido = ''

            # dividir texto por columna con el ancho disponible
            lines0 = _dividir_texto(c, marca, col_w[0], font_name=font_name, font_size=font_size)
            lines1 = _dividir_texto(c, codigo, col_w[1], font_name=font_name, font_size=font_size)
            lines2 = _dividir_texto(c, medidas, col_w[2], font_name=font_name, font_size=font_size)
            lines3 = _dividir_texto(c, contenido, col_w[3], font_name=font_name, font_size=font_size)
            max_lines = max(len(lines0), len(lines1), len(lines2), len(lines3), 1)
            h = max_lines * leading + (4 * mm)
            row_heights.append(h)
            rows_cells.append((lines0, lines1, lines2, lines3))

        table_h = header_h + sum(row_heights)
        table_top = self.cursor_y

        # paginación: si la tabla no cabe en el espacio restante, dividir en varias páginas
        bottom_margin = 25 * mm
        rows_count = len(rows_cells)
        start_idx = 0
        cur_table_top = table_top

        while start_idx < rows_count:
            # espacio disponible en esta "página" para la tabla
            available = cur_table_top - bottom_margin
            # reservar espacio para header
            acc_h = header_h
            end_idx = start_idx
            while end_idx < rows_count and (acc_h + row_heights[end_idx]) <= available:
                acc_h += row_heights[end_idx]
                end_idx += 1

            # si ninguna fila cabe (fila demasiado grande), forzar al menos una fila para evitar bucle infinito
            if end_idx == start_idx and start_idx < rows_count:
                acc_h = header_h + row_heights[start_idx]
                end_idx = start_idx + 1

            # dibujar caja externa para esta porción
            c.rect(x, cur_table_top - acc_h, total_w, acc_h, stroke=1, fill=0)

            # dibujar líneas verticales
            accx = x
            for w in col_w[:-1]:
                accx += w
                c.line(accx, cur_table_top, accx, cur_table_top - acc_h)

            # línea separadora entre header y contenido
            c.setLineWidth(0.8)
            c.line(x, cur_table_top - header_h, x + total_w, cur_table_top - header_h)
            c.setLineWidth(0.6)

            # Encabezados
            c.setFont('Helvetica-Bold', font_size)
            cur_x = x
            for i, h in enumerate(headers):
                w = col_w[i]
                cx = cur_x + w / 2
                cy = cur_table_top - (header_h / 2) - (font_size / 2)
                c.drawCentredString(cx, cy, h)
                cur_x += w

            # Dibujar filas de contenido para esta página
            y = cur_table_top - header_h
            c.setFont(font_name, font_size)
            for ri in range(start_idx, end_idx):
                y -= row_heights[ri]
                # dibujar línea horizontal que cierra la fila
                c.line(x, y, x + total_w, y)
                # dibujar cada celda
                cell_x = x
                cells = rows_cells[ri]
                for ci, lines in enumerate(cells):
                    tx = cell_x + (3 * mm)
                    ty = y + row_heights[ri] - leading - (3 * mm)
                    for ln in lines:
                        try:
                            c.drawString(tx, ty, ln)
                        except Exception:
                            pass
                        ty += -leading
                    cell_x += col_w[ci]

            # preparar para la siguiente página o finalizar
            start_idx = end_idx
            # si quedan más filas, crear nueva página y dibujar fondo
            if start_idx < rows_count:
                try:
                    c.showPage()
                except Exception:
                    pass
                try:
                    self.dibujar_fondo(c)
                except Exception:
                    pass
                # resetear cursor en la nueva página: dejar espacio para encabezado
                cur_table_top = self.height - 120
                self.cursor_y = cur_table_top
            else:
                # última porción: actualizar cursor_y y salir
                self.cursor_y = cur_table_top - acc_h - (8 * mm)
                break

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

            # DEBUG: listar evidencias encontradas
            try:
                dbg_e = [f"[DEBUG evidencia] total={len(evidencias)}"]
                for p in (evidencias or [])[:10]:
                    dbg_e.append(f"[DEBUG evidencia] path={p}")
                log_path = os.path.join(DATA_DIR, 'constancia_debug.log')
                with open(log_path, 'a', encoding='utf-8') as lf:
                    for L in dbg_e:
                        lf.write(L + '\n')
                    lf.write('\n')
                for L in dbg_e:
                    print(L)
            except Exception:
                pass
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

        # --- DEBUG: información sobre firmantes que se imprimirán ---
        try:
            dbg = []
            dbg.append(f"[DEBUG firma] nombre1={nombre1!r} code1={code1!r} img1={img1!r}")
            dbg.append(f"[DEBUG firma] nombre2={nombre2!r} code2={code2!r} img2={img2!r}")
            # escribir también al log de constancias
            try:
                log_path = os.path.join(DATA_DIR, 'constancia_debug.log')
                with open(log_path, 'a', encoding='utf-8') as lf:
                    for L in dbg:
                        lf.write(L + '\n')
                    lf.write('\n')
            except Exception:
                pass
            for L in dbg:
                print(L)
        except Exception:
            pass
        

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
        # Si no se especifica salida, guardar en data/Constancias dentro del proyecto
        if not salida:
            fol = str(self.datos.get('folio_formateado') or self.datos.get('folio_constancia') or '')
            safe = fol.replace('/', '_').replace(' ', '_') or datetime.now().strftime('%Y%m%d_%H%M%S')
            # Usar DATA_DIR persistente (AppData cuando está en .exe) para salidas
            const_dir = os.path.join(DATA_DIR, 'Constancias')
            os.makedirs(const_dir, exist_ok=True)
            salida = os.path.join(const_dir, f'Constancia_{safe}.pdf')

        # Use NumberedCanvas so we can print "Página X de Y" after all pages are created
        try:
            c = NumberedCanvas(salida, pagesize=letter)
        except Exception:
            c = canvas.Canvas(salida, pagesize=letter)
        # Asegurar que el canvas tenga título/cadena para que se impriman en todas las páginas
        try:
            # Forzar re-lectura final de folios desde disco: esto evita que se use
            # un valor cacheado en `self.datos` que pueda estar desincronizado.
            try:
                self._refresh_folio_from_disk()
            except Exception:
                pass
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
                # --- DEBUG: registrar información de generación ---
                try:
                    debug_lines = []
                    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    debug_lines.append(f"[{ts}] Generando constancia: salida={salida}")
                    # Folio valores
                    ff = str(self.datos.get('folio_formateado') or '')
                    fcandidate = str(self.datos.get('folio_constancia') or '')
                    fvis = str(self.datos.get('folio_visita') or '')
                    debug_lines.append(f"folio_formateado={ff}  (folio_constancia={fcandidate})  folio_visita={fvis}")
                    # folio_counter.json
                    try:
                        # leer desde DATA_DIR persistente
                        fc_path = os.path.join(DATA_DIR, 'folio_counter.json')
                        if os.path.exists(fc_path):
                            with open(fc_path, 'r', encoding='utf-8') as fcf:
                                j = json.load(fcf) or {}
                            debug_lines.append(f"folio_counter.last={j.get('last')}")
                    except Exception as e:
                        debug_lines.append(f"folio_counter.read_error={e}")
                    # per-visit file contents (if any)
                    try:
                        fid = fvis or (self.datos.get('folio') or '')
                        if fid:
                            archivo_f = os.path.join(DATA_DIR, 'folios_visitas', f"folios_{fid}.json")
                            if os.path.exists(archivo_f):
                                with open(archivo_f, 'r', encoding='utf-8') as ffp:
                                    obj = json.load(ffc:=ffp) if False else json.load(ffp)
                                # include meta and first folios entries
                                if isinstance(obj, dict):
                                    meta = obj.get('_meta') or {}
                                    debug_lines.append(f"per_visit._meta={meta}")
                                    fols = obj.get('folios') or []
                                else:
                                    fols = obj if isinstance(obj, list) else []
                                debug_lines.append(f"per_visit.total_folios={len(fols)}")
                                if fols:
                                    # show first two entries summary
                                    for i, e in enumerate(fols[:2], start=1):
                                        try:
                                            debug_lines.append(f"per_visit.folio[{i}]={e.get('FOLIOS') or e.get('FOLIO') or e}")
                                        except Exception:
                                            debug_lines.append(f"per_visit.folio[{i}]={str(e)}")
                    except Exception as e:
                        debug_lines.append(f"per_visit.read_error={e}")
                    # evidencias
                    try:
                        evids = list(self.datos.get('evidencias_lista') or [])
                        debug_lines.append(f"evidencias.count={len(evids)}")
                        for p in evids[:5]:
                            debug_lines.append(f"evidencia.path={p}")
                    except Exception:
                        pass
                    # firmantes previstos
                    try:
                        firmas_map = cargar_firmas()
                    except Exception:
                        firmas_map = {}
                    f1 = self.datos.get('nfirma1') or ''
                    f2 = self.datos.get('nfirma2') or ''
                    # intentar resolver desde tabla_relacion FIRMA
                    try:
                        tr0 = list(self.datos.get('tabla_relacion') or [])
                        pref = ''
                        if tr0:
                            pref = (tr0[0].get('FIRMA') or tr0[0].get('firma') or '').strip()
                        if pref and pref in firmas_map:
                            entry = firmas_map.get(pref)
                            f1 = f1 or (entry.get('NOMBRE DE INSPECTOR') or entry.get('nombre') or entry.get('NOMBRE') or '')
                    except Exception:
                        pass
                    debug_lines.append(f"firmante1={f1}  firmante2={f2}  firma_map_keys={len(firmas_map)}")
                    # escribir log
                    try:
                        log_path = os.path.join(DATA_DIR, 'constancia_debug.log')
                        with open(log_path, 'a', encoding='utf-8') as lf:
                            for L in debug_lines:
                                lf.write(L + '\n')
                            lf.write('\n')
                    except Exception:
                        pass
                    # también imprimir en consola para retroalimentación rápida
                    for L in debug_lines:
                        print(L)
                except Exception:
                    pass
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
            data_dir = DATA_DIR
            clientes_path = os.path.join(DATA_DIR, 'Clientes.json')
            normas_path = os.path.join(DATA_DIR, 'Normas.json')
            firmas_path = os.path.join(DATA_DIR, 'Firmas.json')
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


def _reserve_next_folio(data_dir: str) -> str:
    """Reserva y retorna el siguiente folio numérico como cadena (sin ceros a la izquierda).

    Archivo: data/folio_counter.json con formato {"last": N}
    """
    # Preferir delegar a folio_manager (reserva atómica). Si no está disponible,
    # usar un fallback local que adquiere un lock por archivo para evitar
    # condiciones de carrera entre procesos.
    try:
        try:
            # import dinámico para no forzar dependencias en entornos limitados
            from folio_manager import reserve_next
            nxt = reserve_next()
            return str(int(nxt))
        except Exception:
            # Fallback: archivo counter + lock en el mismo directorio `data_dir`
            p = os.path.join(data_dir, 'folio_counter.json')
            lock = os.path.join(data_dir, 'folio_counter.lock')
            timeout = 5.0
            start = time.time()
            # intentar adquirir lock simple mediante os.open O_EXCL
            acquired = False
            while True:
                try:
                    fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    acquired = True
                    break
                except FileExistsError:
                    if (time.time() - start) >= timeout:
                        break
                    time.sleep(0.05)
            try:
                if not os.path.exists(p):
                    with open(p, 'w', encoding='utf-8') as f:
                        json.dump({'last': 0}, f)
                with open(p, 'r', encoding='utf-8') as f:
                    data = json.load(f) or {}
                last = int(data.get('last') or 0)
                nxt = last + 1
                data['last'] = nxt
                # escribir de forma atómica
                tmp = p + '.tmp'
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                try:
                    os.replace(tmp, p)
                except Exception:
                    try:
                        if os.path.exists(p):
                            os.remove(p)
                    except Exception:
                        pass
                    os.replace(tmp, p)
                return str(nxt)
            finally:
                if acquired:
                    try:
                        if os.path.exists(lock):
                            os.remove(lock)
                    except Exception:
                        pass
    except Exception:
        return '1'

def _get_last_historial_fecha(data_dir: str) -> str:
    """Leer `data/historial_visitas.json` y devolver la fecha de la última visita.

    Retorna cadena vacía si no existe o no contiene fechas.
    """
    try:
        p = os.path.join(data_dir, 'historial_visitas.json')
        if not os.path.exists(p):
            return ''
        with open(p, 'r', encoding='utf-8') as f:
            hist = json.load(f) or {}
        visitas = hist.get('visitas', []) if isinstance(hist, dict) else hist
        if not visitas:
            return ''
        last = visitas[-1]
        return last.get('fecha_termino') or last.get('fecha') or last.get('fecha_emision') or ''
    except Exception:
        return ''

def generar_constancia_desde_visita(folio_visita: str | None = None, salida: str | None = None) -> str:
    base = PACKAGE_BASE
    data_dir = DATA_DIR
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
    fol = (visita.get('folio') or visita.get('folio_visita') or '').replace('UDC', f'UCC')
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

    # Extraer campos relevantes desde la primera fila de tabla_relacion cuando exista
    extra = {}
    try:
        if tabla_rows:
            first = tabla_rows[0] if isinstance(tabla_rows, list) else tabla_rows
            # helpers para claves con/ sin mayúsculas y guiones
            def _get(k):
                return first.get(k) or first.get(k.upper()) or first.get(k.replace(' ', '_')) or first.get(k.replace(' ', '')) or ''

            extra['lista'] = _get('LISTA') or _get('Lista') or ''
            extra['fecha_verificacion'] = _get('FECHA DE VERIFICACION') or _get('FECHA_DE_VERIFICACION') or _get('Fecha de Verificacion') or ''
            extra['marca'] = _get('MARCA') or extra.get('marca') or marca
            extra['codigo'] = _get('CODIGO') or _get('Codigo') or ''
            extra['pais_origen'] = _get('PAIS DE ORIGEN') or _get('PAIS_DE_ORIGEN') or _get('PAIS DE PROCEDENCIA') or _get('PAIS_DE_PROCEDENCIA') or ''
            extra['descripcion'] = _get('DESCRIPCION') or _get('Descripcion') or _get('descripcion') or ''
            extra['contenido'] = _get('CONTENIDO') or _get('CONTENIDO NETO') or _get('CONTENIDO_NETO') or ''
            extra['clasif_uva'] = _get('CLASIF UVA') or _get('CLASIF_UVA') or _get('CLASIF') or ''
            extra['norma_uva'] = _get('NORMA UVA') or _get('NORMA_UVA') or ''
            extra['obs_dictamen'] = _get('OBSERVACIONES DICTAMEN') or _get('OBSERVACIONES_DICTAMEN') or _get('OBSERVACIONES') or ''
            extra['tipo_documento'] = _get('TIPO DE DOCUMENTO') or _get('TIPO_DE_DOCUMENTO') or ''
            extra['folio_tabla'] = _get('FOLIO') or ''
            extra['medidas'] = _get('MEDIDAS') or _get('Medidas') or ''
            extra['pais_procedencia'] = _get('PAIS DE PROCEDENCIA') or _get('PAIS_DE_PROCEDENCIA') or ''

            # Si la fila trae solicitud más fiable, actualizar solicitud y formato
            sol_row = str(_get('SOLICITUD') or _get('Solicitud') or _get('solicitud') or '').strip()
            if sol_row:
                solicitud_raw = sol_row
                m2 = re.match(r"^\s*(\d+)(?:[\/\-](\d{2,4}))?\s*$", solicitud_raw)
                if m2:
                    solicitud_num = m2.group(1)
                    suf2 = m2.group(2)
                    if suf2:
                        solicitud_year_full = ("20" + suf2) if len(suf2) == 2 else suf2
                else:
                    nums2 = re.findall(r"\d+", solicitud_raw)
                    solicitud_num = nums2[0] if nums2 else solicitud_raw
                solicitud_formateado = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num

            # Mapear CLASIF UVA a NOM en Normas.json cuando sea numérico
            try:
                clas = extra.get('clasif_uva')
                if clas:
                    clas_s = str(clas).strip()
                    # extraer número si viene con texto
                    mnum = re.search(r"(\d{1,3})", clas_s)
                    if mnum:
                        num = int(mnum.group(1))
                        code3 = f"{num:03d}"
                        mapped_nom = ''
                        mapped_name = ''
                        for nom_code, nom_name in normas.items():
                            if f"-{code3}-" in nom_code or code3 in nom_code:
                                mapped_nom = nom_code
                                mapped_name = nom_name
                                break
                        if mapped_nom:
                            norma_str = mapped_nom
                            nombre_norma = mapped_name
            except Exception:
                pass
    except Exception:
        extra = {}

    # Construir el diccionario final de datos incluyendo los campos extraídos
    # Asegurar folio: si viene vacío o no contiene dígitos, reservar uno nuevo
    try:
        fol_digits = ''.join([c for c in str(fol) if c.isdigit()])
    except Exception:
        fol_digits = ''
    # Preferir folio reservado por el sistema (data/folio_counter.json) cuando exista
    try:
        fc_path = os.path.join(data_dir, 'folio_counter.json')
        if os.path.exists(fc_path):
            with open(fc_path, 'r', encoding='utf-8') as fcf:
                j = json.load(fcf) or {}
                last = j.get('last')
                if last is not None:
                    try:
                        last_int = int(last)
                        if fol_digits:
                            if last_int >= int(fol_digits):
                                fol_digits = str(last_int)
                        else:
                            fol_digits = str(last_int)
                    except Exception:
                        pass
    except Exception:
        pass
    if not fol_digits:
        try:
            fol_digits = _reserve_next_folio(data_dir)
        except Exception:
            fol_digits = '1'
    # formatear a 6 dígitos
    fol_display = fol_digits.zfill(6)

    datos = {
        'folio_constancia': fol,
        'fecha_emision': fecha,
        'folio_visita': visita.get('folio_visita') or folio_visita or '',
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
        'marca': extra.get('marca') or marca,
        'modelo': modelo,
        'tabla_relacion': tabla_rows,
        'folio_formateado': fol_display,
        'lista': extra.get('lista',''),
        'fecha_verificacion': extra.get('fecha_verificacion',''),
        'codigo': extra.get('codigo',''),
        'pais_origen': extra.get('pais_origen',''),
        'descripcion': extra.get('descripcion',''),
        'contenido': extra.get('contenido',''),
        'clasif_uva': extra.get('clasif_uva',''),
        'norma_uva': extra.get('norma_uva',''),
        'obs_dictamen': extra.get('obs_dictamen',''),
        'tipo_documento': extra.get('tipo_documento',''),
        'folio_tabla': extra.get('folio_tabla',''),
        'medidas': extra.get('medidas',''),
        'pais_procedencia': extra.get('pais_procedencia',''),
    }

    # Si existe un archivo con folios guardados para esta visita, preferir el folio guardado allí
    try:
        fid = str(visita.get('folio_visita') or folio_visita or '').strip()
        if fid:
            archivo_folios_visit = os.path.join(data_dir, 'folios_visitas', f"folios_{fid}.json")
            if os.path.exists(archivo_folios_visit):
                try:
                    with open(archivo_folios_visit, 'r', encoding='utf-8') as ff:
                        obj = json.load(ff) or {}
                    fols_list = obj.get('folios') if isinstance(obj, dict) else obj
                    if isinstance(fols_list, list) and fols_list:
                        # Prefer the folio entry that matches the requested 'lista'
                        preferred = None
                        lista_pref = str(datos.get('lista') or '').strip()
                        for ent in fols_list:
                            try:
                                if not ent:
                                    continue
                                ent_lista = str(ent.get('LISTA') or ent.get('lista') or '').strip()
                                if lista_pref and ent_lista and ent_lista == lista_pref:
                                    preferred = ent
                                    break
                            except Exception:
                                continue
                        if not preferred:
                            # fallback: try to match by FOLIO_TABLA or CODE if available
                            folio_tabla_pref = str(datos.get('folio_tabla') or '').strip()
                            codigo_pref = str(datos.get('codigo') or '').strip()
                            if folio_tabla_pref:
                                for ent in fols_list:
                                    try:
                                        if str(ent.get('FOLIO') or ent.get('FOLIOS') or '').strip() == folio_tabla_pref:
                                            preferred = ent
                                            break
                                    except Exception:
                                        continue
                            if not preferred and codigo_pref:
                                for ent in fols_list:
                                    try:
                                        if str(ent.get('CODIGO') or '').strip() == codigo_pref:
                                            preferred = ent
                                            break
                                    except Exception:
                                        continue
                        if not preferred:
                            preferred = fols_list[0]
                        first = preferred or fols_list[0] or {}
                        fval = first.get('FOLIOS') or first.get('FOLIO') or ''
                        fdigits = ''.join([c for c in str(fval) if c.isdigit()])
                        if fdigits:
                            fol_display = fdigits.zfill(6)
                            datos['folio_formateado'] = fol_display
                except Exception:
                    pass
    except Exception:
        pass

    if not salida:
        const_dir = os.path.join(data_dir, 'Constancias')
        os.makedirs(const_dir, exist_ok=True)
        safe = str(fol or 'constancia').replace('/', '_').replace(' ', '_')
        salida = os.path.join(const_dir, f'Constancia_{safe}.pdf')

    gen = ConstanciaPDFGenerator(datos, base_dir=base)
    return gen.generar(salida)

def generar_json_constancias_desde_historial(salida_dir: str | None = None, max_items: int | None = None) -> list:
    """Lee `data/historial_visitas.json` y genera un JSON con los datos de constancia
    para cada visita encontrada. Guarda los JSON en `data/Constancias` o en `salida_dir`.

    Devuelve la lista de rutas de los JSON creados.
    """
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    data_dir = get_data_dir()
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


def convertir_constancia_a_json(datos: dict) -> dict:
    """Convierte una estructura de constancia a un JSON con el mismo acomodo que el dictamen.

    Devuelve un dict serializable; intenta normalizar identificador, norma, fechas,
    cliente, producto, tabla_productos, cantidad_total, observaciones y firmas.
    """
    try:
        # Intentar reconstruir cadena identificadora
        cadena = datos.get('cadena') or ''
        try:
            gen = ConstanciaPDFGenerator(datos)
            cadena = gen.construir_cadena_identificacion() or cadena
        except Exception:
            pass

        # Extraer year/folio/solicitud/lista
        solicitud_raw = str(datos.get('solicitud') or '')
        year = ''
        if '/' in solicitud_raw:
            parts = solicitud_raw.split('/')
            if parts[-1].strip().isdigit():
                year = parts[-1].strip()[-2:]
        if not year:
            fem = str(datos.get('fecha_emision') or '')
            m = re.search(r"(\d{4})", fem)
            if m:
                year = m.group(1)[-2:]

        folio_raw = str(datos.get('folio_constancia') or datos.get('folio_formateado') or '')
        fol_digits = ''.join([c for c in folio_raw if c.isdigit()])
        folio = fol_digits.zfill(6) if fol_digits else folio_raw

        solicitud_num = ''
        if solicitud_raw:
            if '/' in solicitud_raw:
                solicitud_num = solicitud_raw.split('/')[0].strip()
            else:
                solicitud_num = ''.join([c for c in solicitud_raw if c.isdigit()])
        solicitud = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num

        lista = str(datos.get('lista') or '')

        norma_codigo = str(datos.get('norma') or '')
        nombre_norma = datos.get('nombre_norma') or datos.get('normades') or ''

        verificacion = str(datos.get('fecha_verificacion') or '')
        emision = str(datos.get('fecha_emision') or '')

        cliente_nombre = datos.get('cliente','')
        cliente_rfc = datos.get('rfc','')
        producto_desc = datos.get('producto','')

        tabla_productos = []
        total_cantidad = 0
        for row in (datos.get('tabla_relacion') or []):
            marca = row.get('MARCA') or row.get('marca') or ''
            codigo = row.get('CODIGO') or row.get('codigo') or ''
            factura = row.get('FACTURA') or row.get('factura') or ''
            cantidad = row.get('CANTIDAD') or row.get('CANTIDAD ') or 0
            try:
                cant_num = int(cantidad) if isinstance(cantidad, (int, float)) or str(cantidad).isdigit() else 0
            except Exception:
                cant_num = 0
            total_cantidad += cant_num
            tabla_productos.append({'marca': marca, 'codigo': codigo, 'factura': factura, 'cantidad': cant_num})

        cantidad_texto = f"{total_cantidad} unidades" if total_cantidad else '0 unidades'

        # Observaciones y firmas
        observaciones = datos.get('obs_dictamen') or datos.get('obs') or datos.get('OBSERVACIONES') or ''
        firmas_map = cargar_firmas()
        firma1 = {'nombre': datos.get('nfirma1',''), 'valida': False, 'codigo_solicitado': '', 'razon_sin_firma': ''}
        firma2 = {'nombre': datos.get('nfirma2','') or 'Arturo Flores Gomez'}
        try:
            tr0 = list(datos.get('tabla_relacion') or [])
            pref = ''
            if tr0:
                pref = (tr0[0].get('FIRMA') or tr0[0].get('firma') or '').strip()
            if pref and pref in firmas_map:
                entry = firmas_map.get(pref)
                nombre = entry.get('NOMBRE DE INSPECTOR') or entry.get('nombre') or entry.get('NOMBRE') or ''
                firma1['nombre'] = nombre
                firma1['codigo_solicitado'] = pref
                normas_ac = entry.get('Normas acreditadas') or entry.get('normas_acreditadas') or entry.get('Normas') or []
                if norma_codigo and normas_ac and (norma_codigo in normas_ac or any(norma_codigo in na for na in normas_ac)):
                    firma1['valida'] = True
                else:
                    firma1['valida'] = False
                    firma1['razon_sin_firma'] = f"Inspector {nombre} no acreditado para {norma_codigo}"
        except Exception:
            pass

        json_data = {
            'identificacion': {'cadena_identificacion': cadena, 'year': year, 'folio': folio, 'solicitud': solicitud, 'lista': lista},
            'norma': {'codigo': norma_codigo, 'descripcion': nombre_norma, 'capitulo': ''},
            'fechas': {'verificacion': verificacion, 'verificacion_larga': '', 'emision': emision},
            'cliente': {'nombre': cliente_nombre, 'rfc': cliente_rfc},
            'producto': {'descripcion': producto_desc, 'pedimento': ''},
            'tabla_productos': tabla_productos,
            'cantidad_total': {'valor': total_cantidad, 'texto': cantidad_texto},
            'observaciones': observaciones,
            'firmas': {'firma1': firma1, 'firma2': firma2},
            'modo_insertado': datos.get('modo_insertado', 'etiqueta'),
            'etiquetas': {'cantidad': len(datos.get('evidencias_lista', []) or [])}
        }
        return json_data
    except Exception:
        return {}


def generar_constancias_desde_tabla(salida_dir: str | None = None) -> list:
    """Genera una constancia PDF por cada fila en data/tabla_de_relacion.json.

    Devuelve la lista de rutas generadas.
    """
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    data_dir = get_data_dir()
    tabla = os.path.join(data_dir, 'tabla_de_relacion.json')
    normas_p = os.path.join(data_dir, 'Normas.json')

    if not os.path.exists(tabla):
        raise FileNotFoundError(tabla)

    # cargar normas para mapeos
    normas_map = _cargar_normas(normas_p)

    with open(tabla, 'r', encoding='utf-8') as f:
        rows = json.load(f) or []

    out_dir = salida_dir or os.path.join(data_dir, 'Constancias')
    os.makedirs(out_dir, exist_ok=True)
    created = []

    for i, row in enumerate(rows, start=1):
        try:
            # preparar campos básicos
            solicitud_raw = str(row.get('SOLICITUD') or row.get('Solicitud') or row.get('solicitud') or '').strip()
            producto = str(row.get('DESCRIPCION') or '')
            marca = str(row.get('MARCA') or '')

            # mapear CLASIF UVA a NOM y nombre de norma si es posible
            norma_str = ''
            nombre_norma = ''
            clas = row.get('CLASIF UVA') or row.get('CLASIF_UVA') or row.get('CLASIF') or ''
            try:
                if clas:
                    mnum = re.search(r"(\d{1,3})", str(clas))
                    if mnum:
                        code3 = f"{int(mnum.group(1)):03d}"
                        for nom_code, nom_name in normas_map.items():
                            if f"-{code3}-" in nom_code or code3 in nom_code:
                                norma_str = nom_code
                                nombre_norma = nom_name
                                break
            except Exception:
                pass

            # fecha de emisión desde historial o hoy
            fecha_emision = _get_last_historial_fecha(data_dir) or datetime.now().strftime('%d/%m/%Y')

            # reservar folio
            fol_digits = _reserve_next_folio(data_dir)
            fol_display = str(fol_digits).zfill(6)

            datos = {
                'folio_constancia': '',
                'fecha_emision': fecha_emision,
                'cliente': str(row.get('CLIENTE') or ''),
                'rfc': '',
                'no_contrato': '',
                'fecha_contrato': '',
                'solicitud': solicitud_raw,
                'solicitud_formateado': '',
                'norma': norma_str,
                'nombre_norma': nombre_norma,
                'producto': producto,
                'marca': marca,
                'modelo': str(row.get('MODELO') or ''),
                'tabla_relacion': [row],
                'lista': row.get('LISTA',''),
                'clasif_uva': row.get('CLASIF UVA') or row.get('CLASIF_UVA') or row.get('CLASIF') or '',
                'folio_formateado': fol_display,
            }

            out_path = os.path.join(out_dir, f'Constancia_{fol_display}.pdf')
            gen = ConstanciaPDFGenerator(datos, base_dir=base)
            out = gen.generar(out_path)
            created.append(out)
            # Guardar JSON en formato similar al dictamen usando el convertidor
            try:
                # construir nombre similar al que usa la app: incluir lista/folio/solicitud
                sol_no = ''
                sol_year = ''
                if '/' in solicitud_raw:
                    parts = solicitud_raw.split('/')
                    sol_no = parts[0].strip()
                    sol_year = parts[1].strip() if len(parts) > 1 else ''
                else:
                    sol_no = solicitud_raw

                json_name = f"Constancia_Lista_{datos.get('lista','')}_{fol_display}_{sol_no}_{sol_year}.json"
                json_name = json_name.replace('/', '_').replace(' ', '_')
                json_path = os.path.join(out_dir, json_name)
                try:
                    json_data = convertir_constancia_a_json(datos)
                except Exception:
                    json_data = datos
                # añadir metadata
                try:
                    json_data.setdefault('metadata', {})
                    json_data['metadata']['pdf_generado'] = True
                    json_data['metadata']['pdf_path'] = out
                except Exception:
                    pass
                with open(json_path, 'w', encoding='utf-8') as jf:
                    json.dump(json_data, jf, ensure_ascii=False, indent=2)
            except Exception:
                pass
            print('Generated:', out)
        except Exception as e:
            print('Error generating for row', i, e)
            continue

    return created

if __name__ == '__main__':
    import sys

    def _preview_cadena_from_tabla():
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        data_dir = get_data_dir()
        tabla = os.path.join(data_dir, 'tabla_de_relacion.json')
        try:
            with open(tabla, 'r', encoding='utf-8') as f:
                t = json.load(f)
                first = t[0] if isinstance(t, list) and t else {}
        except Exception as e:
            print('Error loading tabla_de_relacion.json:', e)
            return

        datos = {
            'folio_constancia': '',
            'fecha_emision': '',
            'cliente': '',
            'rfc': '',
            'no_contrato': '',
            'fecha_contrato': '',
            'solicitud': str(first.get('SOLICITUD') or '').strip(),
            'solicitud_formateado': '',
            'norma': '',
            'normades': '',
            'nombre_norma': '',
            'producto': str(first.get('DESCRIPCION') or ''),
            'marca': str(first.get('MARCA') or ''),
            'modelo': str(first.get('MODELO') or ''),
            'tabla_relacion': [first],
            'lista': first.get('LISTA',''),
            'clasif_uva': first.get('CLASIF UVA') or first.get('CLASIF_UVA') or first.get('CLASIF') or '',
        }

        gen = ConstanciaPDFGenerator(datos, base_dir=base)
        cadena = gen.construir_cadena_identificacion()
        print('CADENA IDENTIFICACION:')
        print(cadena)

    def generar_ejemplos_integrados(count: int = 3) -> None:
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        data_dir = get_data_dir()
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

    if '--preview-cadena' in sys.argv:
        _preview_cadena_from_tabla()
    else:
        # Si se solicita, generar una constancia a partir de la primera fila de tabla_de_relacion
        if '--generar-constancia-tabla' in sys.argv:
            base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            data_dir = os.path.join(base, 'data')
            tabla = os.path.join(data_dir, 'tabla_de_relacion.json')
            try:
                with open(tabla, 'r', encoding='utf-8') as f:
                    t = json.load(f)
                    first = t[0] if isinstance(t, list) and t else {}
            except Exception as e:
                print('Error loading tabla_de_relacion.json:', e)
                sys.exit(1)

            datos = {
                'folio_constancia': '',
                'fecha_emision': _get_last_historial_fecha(data_dir) or datetime.now().strftime('%d/%m/%Y'),
                'cliente': '',
                'rfc': '',
                'no_contrato': '',
                'fecha_contrato': '',
                'solicitud': str(first.get('SOLICITUD') or '').strip(),
                'solicitud_formateado': '',
                'norma': '',
                'normades': '',
                'nombre_norma': '',
                'producto': str(first.get('DESCRIPCION') or ''),
                'marca': str(first.get('MARCA') or ''),
                'modelo': str(first.get('MODELO') or ''),
                'tabla_relacion': [first],
                'lista': first.get('LISTA',''),
                'clasif_uva': first.get('CLASIF UVA') or first.get('CLASIF_UVA') or first.get('CLASIF') or '',
            }
            # Reservar folio y asignarlo
            fol_digits = _reserve_next_folio(data_dir)
            datos['folio_formateado'] = fol_digits.zfill(6)
            gen = ConstanciaPDFGenerator(datos, base_dir=base)
            out = gen.generar(None)
            print('Generated (from tabla) PDF:', out)
        elif '--generar-constancias-tabla' in sys.argv:
            # Generar una constancia por cada fila de tabla_de_relacion.json
            base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            created = generar_constancias_desde_tabla()
            print('Created', len(created), 'constancias')
        else:
            generar_ejemplos_integrados(3)
