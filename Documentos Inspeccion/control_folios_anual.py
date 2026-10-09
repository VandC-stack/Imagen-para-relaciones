import hashlib
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from datetime import datetime, timedelta
from typing import Callable, Dict, List, Optional, Tuple
import unicodedata
import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Claves del JSON de dictamen que realmente usa el reporte. Se descarta el resto
# para que el índice en memoria/caché siga siendo manejable con decenas de miles
# de archivos en data/Dictamenes.
_CLAVES_DICTAMEN = ('identificacion', 'norma', 'cliente', 'cliente_info', 'clienteDatos',
                    'firmas', 'tabla_productos', 'producto', 'fechas')
_VERSION_CACHE_INDICE = 1

class ControlFoliosAnual:
    """Clase para generar el control de folios anual desde archivos JSON"""
    
    def __init__(self, data_dir: str = "data"):
        """
        Inicializar el generador de control de folios
        
        Args:
            data_dir: Directorio donde se encuentran los archivos JSON
        """
        self.data_dir = data_dir
        self.clientes = []
        self.firmas = []
        self.tabla_relacion = []
        self.historial_visitas = []
        self.folio_to_cliente = {}  # Mapeo de folio a cliente
        self.normas = []
        # Índice de data/Dictamenes + data/Constancias (se construye una sola vez)
        self._indice: Optional[List[Dict]] = None
        self._indice_por_folio: Dict[int, List[Dict]] = {}
        self._find_memo: Dict[Tuple[str, str], Optional[Dict]] = {}
        # Mapeo (SOLICITUD, FOLIO) -> FECHA DE ENTRADA extraída de backups
        self.backup_fecha_entrada: Dict[Tuple[str, str], str] = {}
        # Índices secundarios para _lookup_backup_fecha (primer valor encontrado)
        self._backup_por_sol: Dict[str, str] = {}
        self._backup_por_folio: Dict[str, str] = {}
        self._backup_memo: Dict[str, Optional[str]] = {}
        # Callback opcional de progreso: progress_cb(hechos, total, mensaje)
        self.progress_cb: Optional[Callable[[int, int, str], None]] = None

    def _reportar_progreso(self, hechos: int, total: int, mensaje: str):
        if self.progress_cb:
            try:
                self.progress_cb(hechos, total, mensaje)
            except Exception:
                pass

    def cargar_datos(self) -> Tuple[bool, str]:
        """
        Cargar todos los archivos JSON necesarios
        
        Returns:
            Tuple[bool, str]: (éxito, mensaje)
        """
        try:
            # Cargar Clientes.json
            clientes_path = os.path.join(self.data_dir, "Clientes.json")
            if os.path.exists(clientes_path):
                try:
                    with open(clientes_path, 'r', encoding='utf-8') as f:
                        self.clientes = json.load(f)
                    print(f"✅ Clientes cargados: {len(self.clientes)} registros")
                except json.JSONDecodeError as e:
                    return False, f"Error al decodificar JSON en Clientes.json: {e}"
                except Exception as e:
                    return False, f"Error al cargar Clientes.json: {e}"
            else:
                print(f"⚠️ Advertencia: No se encontró {clientes_path}. Continuando con clientes vacíos.")
                self.clientes = []

            # Cargar Firmas.json
            firmas_path = os.path.join(self.data_dir, "Firmas.json")
            if os.path.exists(firmas_path):
                try:
                    with open(firmas_path, 'r', encoding='utf-8') as f:
                        self.firmas = json.load(f)
                    print(f"✅ Firmas cargadas: {len(self.firmas)} registros")
                except json.JSONDecodeError as e:
                    return False, f"Error al decodificar JSON en Firmas.json: {e}"
                except Exception as e:
                    return False, f"Error al cargar Firmas.json: {e}"
            else:
                print(f"⚠️ Advertencia: No se encontró {firmas_path}. Continuando con firmas vacías.")
                self.firmas = []

            # Cargar tabla_de_relacion.json
            tabla_path = os.path.join(self.data_dir, "tabla_de_relacion.json")
            if os.path.exists(tabla_path):
                try:
                    with open(tabla_path, 'r', encoding='utf-8') as f:
                        self.tabla_relacion = json.load(f)
                    print(f"✅ Tabla de relación cargada: {len(self.tabla_relacion)} registros")
                except json.JSONDecodeError as e:
                    return False, f"Error al decodificar JSON en tabla_de_relacion.json: {e}"
                except Exception as e:
                    return False, f"Error al cargar tabla_de_relacion.json: {e}"
            else:
                print(f"⚠️ Advertencia: No se encontró {tabla_path}. Continuando con tabla_de_relacion vacía.")
                self.tabla_relacion = []

            # Cargar backups de tabla_relacion (cada visita genera una copia aquí)
            # y construir un mapeo (SOLICITUD, FOLIO) -> FECHA DE ENTRADA
            try:
                backups_dir = os.path.join(self.data_dir, "tabla_relacion_backups")
                if os.path.exists(backups_dir):
                    for bf in os.listdir(backups_dir):
                        if not bf.lower().endswith('.json'):
                            continue
                        bfp = os.path.join(backups_dir, bf)
                        try:
                            with open(bfp, 'r', encoding='utf-8') as bf_f:
                                bdata = json.load(bf_f)
                        except Exception:
                            continue

                        # normalizar lista de registros si es dict o lista
                        records = []
                        if isinstance(bdata, dict):
                            # buscar primer value que sea lista
                            for v in bdata.values():
                                if isinstance(v, list):
                                    records = v
                                    break
                            if not records:
                                # fallback: si el dict parece ya una lista por claves numéricas
                                try:
                                    # intentar interpretar como lista de diccionarios
                                    possible = list(bdata.values())
                                    if all(isinstance(x, dict) for x in possible):
                                        records = possible
                                except Exception:
                                    records = []
                        elif isinstance(bdata, list):
                            records = bdata

                        for r in records:
                            sol = str(r.get('SOLICITUD', '')).strip()
                            fol = str(r.get('FOLIO', '')).strip()
                            fecha_ent = r.get('FECHA DE ENTRADA') or r.get('FECHA_DE_ENTRADA')
                            if sol and fol and fecha_ent:
                                key = (sol, fol)
                                # preferir la primera encontrada (no sobrescribir)
                                if key not in self.backup_fecha_entrada:
                                    self.backup_fecha_entrada[key] = str(fecha_ent)
                                self._backup_por_sol.setdefault(sol, str(fecha_ent))
                                self._backup_por_folio.setdefault(fol, str(fecha_ent))
                else:
                    # si no existe carpeta de backups, seguir silenciosamente
                    pass
            except Exception:
                # no bloquear la carga por errores en backups
                pass

            # Cargar Normas.json (opcional, para mostrar nombres completos de NOM)
            normas_path = os.path.join(self.data_dir, "Normas.json")
            if os.path.exists(normas_path):
                with open(normas_path, 'r', encoding='utf-8') as f:
                    try:
                        self.normas = json.load(f)
                        print(f"✅ Normas cargadas: {len(self.normas)}")
                    except Exception:
                        self.normas = []
            
            # Cargar historial_visitas.json (opcional, para mapeo de clientes)
            historial_path = os.path.join(self.data_dir, "historial_visitas.json")
            if os.path.exists(historial_path):
                try:
                    with open(historial_path, 'r', encoding='utf-8') as f:
                        hist_data = json.load(f)
                    if isinstance(hist_data, dict) and 'visitas' in hist_data:
                        self.historial_visitas = hist_data['visitas']
                        self._crear_mapeo_folio_cliente()
                        print(f"✅ Historial de visitas cargado: {len(self.historial_visitas)} registros")
                except Exception:
                    pass
            
            return True, "Datos cargados correctamente"
            
        except json.JSONDecodeError as e:
            return False, f"Error al decodificar JSON: {e}"
        except Exception as e:
            return False, f"Error al cargar datos: {e}"
    
    def _crear_mapeo_folio_cliente(self):
        """
        Crear un mapeo entre folios y clientes desde el historial de visitas
        """
        for visita in self.historial_visitas:
            cliente_nombre = visita.get('cliente', '')
            folios_str = visita.get('folios_utilizados', '')
            
            if not cliente_nombre or not folios_str:
                continue
            
            # Parse folio range (e.g., "075339 - 075552")
            if ' - ' in folios_str:
                parts = folios_str.split(' - ')
                if len(parts) == 2:
                    try:
                        inicio = int(parts[0].strip())
                        fin = int(parts[1].strip())

                        # Map all folios in range to this client
                        for folio_num in range(inicio, fin + 1):
                            self.folio_to_cliente[folio_num] = cliente_nombre
                    except ValueError:
                        pass
    
    def buscar_cliente_por_solicitud(self, solicitud: str, folio: int) -> Optional[Dict]:
        """
        Buscar información del cliente basándose en el folio
        
        Args:
            solicitud: Número de solicitud (ej: "006916/25") - no usado directamente
            folio: Número de folio para buscar el cliente
            
        Returns:
            Diccionario con información del cliente o None
        """
        # Primero intentar buscar por folio en el historial
        cliente_nombre = self.folio_to_cliente.get(folio)
        
        if cliente_nombre:
            # Buscar información completa del cliente por nombre
            for cliente in self.clientes:
                if cliente.get('CLIENTE', '').strip().upper() == cliente_nombre.strip().upper():
                    return cliente
        
        # Si no se encontró, retornar información genérica con el nombre del historial
        if cliente_nombre:
            return {
                'CLIENTE': cliente_nombre,
                'NÚMERO_DE_CONTRATO': 'N/A',
                'RFC': 'N/A',
                'CURP': 'N/A'
            }
        
        # Como último recurso, retornar N/A
        return {
            'CLIENTE': 'N/A',
            'NÚMERO_DE_CONTRATO': 'N/A',
            'RFC': 'N/A',
            'CURP': 'N/A'
        }
    
    def buscar_inspector_por_firma(self, firma: str) -> str:
        """
        Buscar el nombre completo del inspector por su firma
        
        Args:
            firma: Código de firma (ej: "GRAMIREZ")
            
        Returns:
            Nombre completo del inspector o "N/A"
        """
        for inspector in self.firmas:
            if inspector.get("FIRMA") == firma:
                # Intentar extraer nombre por varias claves posibles
                nombre = (
                    inspector.get("NOMBRE") or inspector.get("NOMBRE_COMPLETO")
                    or inspector.get("NOMBRE DE INSPECTOR") or inspector.get("nombre")
                    or inspector.get("NOMBRE INSPECTOR") or inspector.get("NOMBRE_COMPLETO_INSPECTOR")
                    or inspector.get("FIRMA")
                )
                return self._normalize_name(nombre)
        return "N/A"

    def _format_date(self, value) -> str:
        """
        Normaliza una fecha a formato DD/MM/YYYY. Acepta strings en
        YYYY-MM-DD, DD/MM/YYYY, datetime.date/datetime.datetime.
        Si no puede parsear, devuelve el valor original o 'N/A'.
        """
        if not value:
            return "N/A"
        try:
            # Already a date/datetime
            if hasattr(value, 'strftime'):
                return value.strftime("%d/%m/%Y")
            s = str(value).strip()
            # Número de serie de Excel (ej. 46094 -> 10/03/2026) que quedó así
            # al importar la tabla. Rango acotado a 2000-2064 para no confundir
            # otros números con fechas.
            if re.fullmatch(r"\d{5}(\.0+)?", s) and 36526 <= int(float(s)) <= 60000:
                return (datetime(1899, 12, 30) + timedelta(days=int(float(s)))).strftime("%d/%m/%Y")
            # If already in DD/MM/YYYY
            if '/' in s and len(s.split('/')[0]) <= 2:
                parts = s.split('/')
                if len(parts[0]) == 1:
                    parts[0] = parts[0].zfill(2)
                return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}/{parts[2]}"
            # Try ISO YYYY-MM-DD
            try:
                dt = datetime.strptime(s, "%Y-%m-%d")
                return dt.strftime("%d/%m/%Y")
            except Exception:
                pass
            # Try other common formats
            for fmt in ("%Y/%m/%d", "%d-%m-%Y", "%d.%m.%Y"):
                try:
                    dt = datetime.strptime(s, fmt)
                    return dt.strftime("%d/%m/%Y")
                except Exception:
                    continue
        except Exception:
            pass
        return str(value)

    def _normalize_name(self, name: Optional[str]) -> str:
        """
        Normalizar nombre: convertir a mayúsculas y quitar acentos.
        """
        if not name:
            return "N/A"
        nfkd = unicodedata.normalize('NFKD', str(name))
        only_ascii = ''.join([c for c in nfkd if not unicodedata.combining(c)])
        return only_ascii.upper()
    
    def formatear_folio_ema(self, folio) -> str:
        """
        Formatear el folio EMA a 6 dígitos
        
        Args:
            folio: Número de folio
            
        Returns:
            Folio formateado a 6 dígitos
        """
        try:
            folio_str = str(int(folio))
            return folio_str.zfill(6)
        except (ValueError, TypeError):
            return "000000"
    
    def extraer_sol_ema(self, numero_solicitud: str) -> str:
        """
        Extraer los últimos valores del número de solicitud
        
        Args:
            numero_solicitud: Número de solicitud completo (ej: "006916/25")
            
        Returns:
            Últimos valores separados por guión
        """
        if not numero_solicitud:
            return "N/A"
        # Extraer los últimos componentes separados por '/'
        partes = numero_solicitud.split('/')
        if len(partes) >= 2:
            # construir sol-xxx, pero quitar sufijo de año si el último componente es año corto (ej '25')
            last = partes[-1]
            penult = partes[-2]
            if last.isdigit() and len(last) == 2:
                return penult
            return f"{penult}-{last}"
        return numero_solicitud

    # ------------------------------------------------------------------
    # Índice de dictámenes
    # ------------------------------------------------------------------
    # Con decenas de miles de archivos en data/Dictamenes, abrir cada JSON en
    # cada búsqueda vuelve el reporte inutilizable (y en carpetas sincronizadas
    # con Dropbox/OneDrive cada lectura puede forzar una descarga). Se recorre
    # la carpeta una sola vez, se guarda sólo lo necesario de cada JSON y se
    # persiste en una caché local (fuera de la carpeta sincronizada) que se
    # invalida por archivo usando tamaño + fecha de modificación.

    def _ruta_cache_indice(self) -> str:
        base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), '.cache')
        clave = hashlib.md5(os.path.abspath(self.data_dir).lower().encode('utf-8')).hexdigest()[:12]
        return os.path.join(base, 'GeneradorDictamenes', 'cache', f'indice_dictamenes_{clave}.json')

    @staticmethod
    def _reducir_dictamen(d) -> Optional[Dict]:
        """Conservar sólo las claves del JSON que usa el reporte."""
        if not isinstance(d, dict):
            return None
        r = {k: d[k] for k in _CLAVES_DICTAMEN if k in d}
        tp = r.get('tabla_productos')
        if isinstance(tp, list):
            r['tabla_productos'] = [
                {'codigo': p.get('codigo'), 'marca': p.get('marca')}
                for p in tp if isinstance(p, dict)
            ]
        firmas = r.get('firmas')
        if isinstance(firmas, dict):
            reducidas = {}
            for k in ('firma1', 'firma2'):
                f = firmas.get(k)
                if isinstance(f, dict):
                    reducidas[k] = {c: f.get(c) for c in ('nombre', 'codigo_solicitado', 'codigo') if c in f}
            r['firmas'] = reducidas
        return r

    def _cargar_indice_dictamenes(self) -> List[Dict]:
        """Devuelve [{'sub', 'name', 'tipo', 'ok', 'd'}] para Dictamenes y Constancias."""
        if self._indice is not None:
            return self._indice

        cache_path = self._ruta_cache_indice()
        cache_previa: Dict[str, Dict] = {}
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                c = json.load(f)
            if isinstance(c, dict) and c.get('version') == _VERSION_CACHE_INDICE:
                cache_previa = c.get('archivos') or {}
        except Exception:
            cache_previa = {}

        # 1) Listar archivos (scandir trae tamaño/fecha sin abrir los archivos)
        archivos = []
        for sub in ('Dictamenes', 'Constancias'):
            dpath = os.path.join(self.data_dir, sub)
            try:
                with os.scandir(dpath) as it:
                    for e in it:
                        low = e.name.lower()
                        if not (low.endswith('.json') or low.endswith('.pdf')):
                            continue
                        try:
                            st = e.stat()
                            archivos.append((sub, e.name, e.path, st.st_mtime, st.st_size))
                        except OSError:
                            continue
            except OSError:
                continue

        # 2) Leer sólo los archivos nuevos o modificados. Cada apertura tiene
        # una latencia alta (antivirus / sincronización), así que se leen en
        # paralelo: el costo es de espera de E/S, no de CPU.
        def _leer(item):
            sub, name, path, _mtime, _size = item
            if not name.lower().endswith('.json'):
                return {'sub': sub, 'name': name, 'tipo': 'pdf', 'ok': True, 'd': None}
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    d = self._reducir_dictamen(json.load(f))
                return {'sub': sub, 'name': name, 'tipo': 'json', 'ok': d is not None, 'd': d}
            except Exception:
                return {'sub': sub, 'name': name, 'tipo': 'json', 'ok': False, 'd': None}

        total = len(archivos)
        entradas: List[Optional[Dict]] = [None] * total
        pendientes = []
        for i, item in enumerate(archivos):
            prev = cache_previa.get(f"{item[0]}/{item[1]}")
            if prev and prev.get('mtime') == item[3] and prev.get('size') == item[4]:
                entradas[i] = prev['rec']
            else:
                pendientes.append(i)

        leidos = len(pendientes)
        hechos = total - leidos
        self._reportar_progreso(hechos, total, f"Indexando dictámenes ({hechos:,}/{total:,})")
        if pendientes:
            with ThreadPoolExecutor(max_workers=16) as ex:
                for n, (i, entrada) in enumerate(zip(pendientes, ex.map(lambda i: _leer(archivos[i]), pendientes)), 1):
                    entradas[i] = entrada
                    if n % 500 == 0 or n == leidos:
                        self._reportar_progreso(hechos + n, total, f"Indexando dictámenes ({hechos + n:,}/{total:,})")

        indice: List[Dict] = entradas  # type: ignore[assignment]
        cache_nueva: Dict[str, Dict] = {
            f"{sub}/{name}": {'mtime': mtime, 'size': size, 'rec': entrada}
            for (sub, name, _path, mtime, size), entrada in zip(archivos, indice)
        }

        # 3) Persistir caché si algo cambió
        if leidos or len(cache_nueva) != len(cache_previa):
            try:
                os.makedirs(os.path.dirname(cache_path), exist_ok=True)
                tmp = cache_path + '.tmp'
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump({'version': _VERSION_CACHE_INDICE, 'data_dir': os.path.abspath(self.data_dir),
                               'archivos': cache_nueva}, f, ensure_ascii=False, separators=(',', ':'))
                os.replace(tmp, cache_path)
            except Exception as e:
                print(f"⚠️ No se pudo guardar la caché del índice de dictámenes: {e}")

        print(f"✅ Índice de dictámenes: {total} archivos ({leidos} leídos de disco, {total - leidos} desde caché)")

        # Índice auxiliar por folio numérico (identificacion.folio) para búsquedas O(1)
        por_folio: Dict[int, List[Dict]] = {}
        for entrada in indice:
            d = entrada.get('d')
            if not d:
                continue
            fol = str((d.get('identificacion') or {}).get('folio', '')).strip()
            digits = ''.join(ch for ch in fol if ch.isdigit())
            if digits:
                por_folio.setdefault(int(digits), []).append(entrada)

        self._indice = indice
        self._indice_por_folio = por_folio
        return indice

    @staticmethod
    def _dictamen_desde_entrada(entrada: Dict) -> Optional[Dict]:
        """JSON reducido de la entrada, o un dict mínimo construido del nombre si es PDF."""
        if entrada.get('tipo') == 'json':
            return entrada.get('d') if entrada.get('ok') else None
        name_base = os.path.splitext(entrada.get('name', ''))[0]
        digit_sequences = re.findall(r"(\d+)", name_base)
        fol_guess = max(digit_sequences, key=lambda s: len(s)) if digit_sequences else ''
        sol_match = re.search(r"([A-Za-z0-9\-]{4,})", name_base)
        sol_guess = sol_match.group(1) if sol_match else ''
        return {
            'identificacion': {
                'solicitud': sol_guess,
                'folio': fol_guess,
                'cadena_identificacion': ''
            }
        }

    def _find_dictamen(self, solicitud: str, folio) -> Optional[Dict]:
        """Buscar dictamen JSON en data/Dictamenes que coincida con solicitud y folio."""
        try:
            # Normalizar folio y solicitud para mejorar coincidencias
            folio_s = str(folio).strip() if folio is not None else ''
            sol_search = str(solicitud).strip() if solicitud is not None else ''
            # Si viene con formato 'XXXX/25', usar la parte antes de la barra
            sol_base = sol_search.split('/')[0] if '/' in sol_search else sol_search

            memo_key = (folio_s, sol_base)
            if memo_key not in self._find_memo:
                self._find_memo[memo_key] = self._find_dictamen_en_indice(folio_s, sol_base)
            return self._find_memo[memo_key]
        except Exception:
            return None

    def _find_dictamen_en_indice(self, folio_s: str, sol_base: str) -> Optional[Dict]:
        indice = self._cargar_indice_dictamenes()
        if not indice:
            return None

        # Ruta rápida: mismo folio numérico y solicitud compatible
        digits_b = ''.join(ch for ch in folio_s if ch.isdigit())
        if digits_b:
            for entrada in self._indice_por_folio.get(int(digits_b), []):
                d = entrada['d']
                ident = d.get('identificacion', {})
                sol_file = str(ident.get('solicitud', '')).strip()
                cadena = (ident.get('cadena_identificacion') or '')
                if not sol_base or (sol_file and (sol_file == sol_base or sol_file.endswith(sol_base))) \
                        or (cadena and sol_base in cadena):
                    return d

        # Búsqueda completa (en memoria) con las mismas reglas de coincidencia de siempre
        for entrada in indice:
            fname = entrada.get('name', '')
            d = self._dictamen_desde_entrada(entrada)
            if d is None:
                continue

            ident = d.get('identificacion', {})
            sol_file = str(ident.get('solicitud', '')).strip()
            fol_file = str(ident.get('folio', '')).strip()
            cadena = (ident.get('cadena_identificacion') or '')

            # 1) Coincidencia exacta en folio y solicitud (cuando ambos estén presentes)
            if fol_file and folio_s:
                equal_folio = False
                digits_a = ''
                try:
                    # Comparar como texto (posible padded) y como dígitos (ignorar ceros)
                    if fol_file == folio_s:
                        equal_folio = True
                    else:
                        digits_a = ''.join(ch for ch in fol_file if ch.isdigit())
                        if digits_a and digits_b and int(digits_a) == int(digits_b):
                            equal_folio = True
                except Exception:
                    equal_folio = False
                if equal_folio:
                    if sol_base:
                        if sol_file and (sol_file == sol_base or sol_file.endswith(sol_base)):
                            return d
                        # También checar si la solicitud completa aparece en la cadena_identificacion
                        if cadena and sol_base in cadena:
                            return d
                        # también intentar por filename (considerar variantes sin padding)
                        fname_digits = ''.join(ch for ch in fname if ch.isdigit())
                        if (f"_{folio_s}_" in fname and f"_{sol_base}_" in fname) or (digits_a and str(int(digits_a)) in fname_digits):
                            return d
                    else:
                        return d

            # 2) Coincidencia exacta en solicitud cuando folio no disponible
            if sol_base and sol_file and (sol_file == sol_base or sol_file.endswith(sol_base)):
                return d

            # 3) Intentar por filename patrones (folio y solicitud dentro del nombre)
            if folio_s and f"_{folio_s}_" in fname:
                return d

            # 4) Por último, coincidencia en cadena_identificacion
            if cadena and sol_base and sol_base in cadena:
                return d
        # Si no encontramos ninguna coincidencia estricta, devolvemos None
        return None

    def _lookup_backup_fecha(self, solicitud: str, folio) -> Optional[str]:
        """
        Buscar una FECHA DE ENTRADA en el mapeo de backups usando varias heurísticas.
        """
        try:
            sol = str(solicitud).strip() if solicitud is not None else ''
            fol_s = str(folio).strip() if folio is not None else ''

            # intento 1: búsqueda exacta
            key = (sol, fol_s)
            if key in self.backup_fecha_entrada:
                return self.backup_fecha_entrada[key]

            # intento 2: buscar con la parte antes de '/' en la solicitud
            sol_base = sol.split('/')[0] if '/' in sol else sol
            if fol_s:
                fecha = self.backup_fecha_entrada.get((sol_base, fol_s))
            else:
                fecha = self._backup_por_sol.get(sol_base)
            if fecha:
                return fecha

            # intento 3: buscar por folio solamente
            if fol_s and fol_s in self._backup_por_folio:
                return self._backup_por_folio[fol_s]

            # intento 4: buscar por coincidencia parcial en solicitud
            if not sol:
                return None
            # (sólo depende de la solicitud: memorizar por solicitud)
            if sol not in self._backup_memo:
                encontrada = None
                for (k_sol, _k_fol), fecha in self.backup_fecha_entrada.items():
                    if sol in k_sol:
                        encontrada = fecha
                        break
                self._backup_memo[sol] = encontrada
            return self._backup_memo[sol]

        except Exception:
            return None

    def agrupar_por_dictamen(self) -> List[Dict]:
        """
        Agrupar los registros de tabla_relacion por dictamen (SOLICITUD + FOLIO)
        
        Returns:
            Lista de dictámenes agrupados con su información
        """
        dictamenes = {}
        
        for registro in self.tabla_relacion:
            solicitud = registro.get("SOLICITUD", "")
            folio = registro.get("FOLIO", "")
            
            # Crear clave única por dictamen
            clave_dictamen = f"{solicitud}_{folio}"
            
            if clave_dictamen not in dictamenes:
                dictamenes[clave_dictamen] = {
                    "solicitud": solicitud,
                    "folio": folio,
                    "registros": []
                }
            
            dictamenes[clave_dictamen]["registros"].append(registro)
        
        return list(dictamenes.values())
    
    def generar_fila_excel(self, dictamen: Dict) -> Dict:
        """
        Generar una fila del Excel a partir de un dictamen
        
        Args:
            dictamen: Diccionario con información del dictamen
            
        Returns:
            Diccionario con los datos de la fila
        """
        registros = dictamen["registros"]
        primer_registro = registros[0] if registros else {}
        
        solicitud = dictamen["solicitud"]
        folio = dictamen["folio"]
        
        # Buscar cliente usando el folio
        try:
            folio_num = int(folio) if folio else 0
        except (ValueError, TypeError):
            folio_num = 0

        cliente = self.buscar_cliente_por_solicitud(solicitud, folio_num)

        # asegurar variable usada en heurísticas aunque se asigne más abajo
        cadena_ident = None

        # Si la búsqueda por historial no arrojó datos, intentar heurísticas
        if cliente is None or cliente.get('CLIENTE', 'N/A') in (None, '', 'N/A'):
            # 1) intentar encontrar número de contrato en el primer registro o en los registros
            contrato = None
            # campos comunes donde podría aparecer
            posible_keys = ['NÚMERO_DE_CONTRATO', 'NUMERO_DE_CONTRATO', 'NÚMERO DE CONTRATO', 'NUMERO DE CONTRATO', 'CONTRATO']
            for k in posible_keys:
                val = primer_registro.get(k) or primer_registro.get(k.upper()) or primer_registro.get(k.lower())
                if val and isinstance(val, str) and val.strip() and val.strip() != 'N/A':
                    contrato = str(val).strip()
                    break

            # 2) buscar en todos los registros si no se encontró en el primero
            if not contrato:
                for r in registros:
                    for k in posible_keys:
                        val = r.get(k) or r.get(k.upper()) or r.get(k.lower())
                        if val and isinstance(val, str) and val.strip() and val.strip() != 'N/A':
                            contrato = str(val).strip()
                            break
                    if contrato:
                        break

            # 3) intentar extraer un token parecido a contrato desde cadena_ident o desde otros campos
            if not contrato:
                # buscar tokens tipo 25049U... en cadena_ident o en solicitud
                search_sources = [cadena_ident or '', solicitud or '']
                for src in search_sources:
                    if not src:
                        continue
                    m = re.search(r"([0-9]{4,}[A-Z0-9\-]{0,30})", str(src))
                    if m:
                        contrato = m.group(1)
                        break

            # 4) si tenemos un contrato, buscar en self.clientes por NÚMERO_DE_CONTRATO o RFC
            if contrato:
                found = None
                for c in self.clientes:
                    try:
                        if str(c.get('NÚMERO_DE_CONTRATO') or c.get('NUMERO_DE_CONTRATO') or '').strip() == contrato:
                            found = c
                            break
                        if str(c.get('RFC') or '').strip() == contrato:
                            found = c
                            break
                    except Exception:
                        continue
                if found:
                    cliente = found

        # Asegurar que cliente sea un dict con claves esperadas
        if not cliente or not isinstance(cliente, dict):
            cliente = {'CLIENTE': 'N/A', 'NÚMERO_DE_CONTRATO': 'N/A', 'RFC': 'N/A', 'CURP': 'N/A'}

        # Intentar localizar dictamen JSON para extraer cadena_identificacion y norma completa
        cadena_ident = None
        norma_codigo = None
        # Si el dictamen ya lleva el JSON embebido (creado en crear_excel), úsalo
        if dictamen.get('dictamen_json'):
            dictamen_json = dictamen.get('dictamen_json')
            ident = dictamen_json.get('identificacion', {})
            cadena_ident = ident.get('cadena_identificacion')
            norma = dictamen_json.get('norma', {})
            norma_codigo = norma.get('codigo') or norma.get('NOM')
        else:
            dictamen_json = self._find_dictamen(solicitud, folio)
            if dictamen_json:
                ident = dictamen_json.get('identificacion', {})
                cadena_ident = ident.get('cadena_identificacion')
                norma = dictamen_json.get('norma', {})
                norma_codigo = norma.get('codigo') or norma.get('NOM')

        # Si el dictamen JSON contiene información del cliente, usarla (sobrescribe heurísticas)
        try:
            if dictamen_json and isinstance(dictamen_json, dict):
                cinfo = dictamen_json.get('cliente') or dictamen_json.get('cliente_info') or dictamen_json.get('clienteDatos')
                if cinfo and isinstance(cinfo, dict):
                    # Normalizar posibles claves
                    nombre = cinfo.get('nombre') or cinfo.get('CLIENTE') or cinfo.get('razon_social') or cinfo.get('nombre_cliente')
                    contrato = cinfo.get('numero_de_contrato') or cinfo.get('NÚMERO_DE_CONTRATO') or cinfo.get('numero_contrato') or cinfo.get('N_CONTRATO')
                    rfc_val = cinfo.get('rfc') or cinfo.get('RFC') or cinfo.get('rfc_cliente')
                    curp = cinfo.get('curp') or cinfo.get('CURP')
                    if nombre:
                        cliente = {
                            'CLIENTE': nombre,
                            'NÚMERO_DE_CONTRATO': contrato or (cliente.get('NÚMERO_DE_CONTRATO') if cliente and isinstance(cliente, dict) else 'N/A'),
                            'RFC': rfc_val or (cliente.get('RFC') if cliente and isinstance(cliente, dict) else 'N/A'),
                            'CURP': curp or (cliente.get('CURP') if cliente and isinstance(cliente, dict) else 'N/A')
                        }
        except Exception:
            pass
        
        # Obtener información del inspector
        firma = primer_registro.get("FIRMA", "")
        nombre_inspector = self.buscar_inspector_por_firma(firma)
        # Fallbacks si no se encontró por código de firma
        if not nombre_inspector or nombre_inspector == 'N/A':
            # Revisar si el registro contiene el nombre explícito
            posibles_nombres = [
                primer_registro.get('NOMBRE INSPECTOR'), primer_registro.get('Inspector'),
                primer_registro.get('INSPECTOR'), primer_registro.get('NOMBRE DE INSPECTOR')
            ]
            for pn in posibles_nombres:
                if pn:
                    nombre_inspector = self._normalize_name(pn)
                    break
        # Intentar obtener desde el JSON del dictamen si aún no hay nombre
        if (not nombre_inspector or nombre_inspector == 'N/A') and dictamen_json and isinstance(dictamen_json, dict):
            firmas_block = dictamen_json.get('firmas') or {}
            f1 = firmas_block.get('firma1') or {}
            f2 = firmas_block.get('firma2') or {}
            n1 = f1.get('nombre') if isinstance(f1, dict) else None
            n2 = f2.get('nombre') if isinstance(f2, dict) else None
            if n1:
                nombre_inspector = self._normalize_name(n1)
            elif n2:
                nombre_inspector = self._normalize_name(n2)
        
        # Extraer descripciones, marcas, NOMs y modelos de todos los registros
        descripciones = set()
        marcas = set()
        noms = set()

        # Preferir los códigos que vienen en el JSON del dictamen (tabla_productos)
        modelos = []
        if dictamen.get('dictamen_json'):
            try:
                tp = dictamen['dictamen_json'].get('tabla_productos', [])
                for p in tp:
                    c = p.get('codigo')
                    if c is not None:
                        modelos.append(str(c))
                # También extraer marcas y descripciones desde el JSON si no vienen en registros
                prod = dictamen['dictamen_json'].get('producto', {})
                if prod.get('descripcion'):
                    descripciones.add(prod.get('descripcion'))
                if tp and isinstance(tp, list):
                    for p in tp:
                        m = p.get('marca')
                        if m:
                            marcas.add(m)
            except Exception:
                modelos = []

        # Si no hay dictamen_json, o además, recorrer registros (tabla_relacion) para completar datos
        for reg in registros:
            if reg.get("DESCRIPCION"):
                descripciones.add(reg.get("DESCRIPCION"))
            if reg.get("MARCA"):
                marcas.add(reg.get("MARCA"))
            if reg.get("CLASIF UVA"):
                noms.add(str(reg.get("CLASIF UVA")))

            # Aceptar tanto 'CODIGO' simple como 'CODIGOS' lista en registros de tabla_relacion
            if reg.get("CODIGO"):
                val = reg.get("CODIGO")
                if isinstance(val, (list, tuple)):
                    for v in val:
                        modelos.append(str(v))
                else:
                    if isinstance(val, str) and ',' in val:
                        for v in [x.strip() for x in val.split(',') if x.strip()]:
                            modelos.append(str(v))
                    else:
                        modelos.append(str(val))
            if reg.get("CODIGOS"):
                try:
                    for c in reg.get("CODIGOS"):
                        modelos.append(str(c))
                except Exception:
                    pass
        
        # Preparar valores derivados
        numero_solicitud_display = None
        if cadena_ident:
            # Intentar extraer token después de 'Solicitud de Servicio:'
            m = re.search(r"Solicitud de Servicio:\s*([A-Za-z0-9\-]+)", cadena_ident)
            if m:
                numero_solicitud_display = m.group(1)
            else:
                m2 = re.search(r"([A-Za-z0-9\-]+-[0-9]+)$", cadena_ident)
                if m2:
                    numero_solicitud_display = m2.group(1)
                else:
                    numero_solicitud_display = cadena_ident

        numero_solicitud_display = numero_solicitud_display or solicitud or "N/A"

        # Tipo de documento (mapear letra a texto)
        tipo_raw = primer_registro.get("TIPO DE DOCUMENTO") or primer_registro.get("TIPO DE DOCUMENTO OFICIAL EMITIDO", "D")
        _t = str(tipo_raw).strip().upper() if tipo_raw is not None else ''
        if _t == 'D':
            tipo_display = "Dictamen"
        elif _t == 'C':
            tipo_display = "Constancia"
        else:
            tipo_display = str(tipo_raw)

        # NOM: preferir norma del dictamen, si no mapear CLASIF UVA usando Normas.json
        if norma_codigo:
            nom_display = norma_codigo
        else:
            mapped = []
            for c in noms:
                mapped_nom = None
                try:
                    ci = int(c)
                    padded = f"{ci:03d}"
                except Exception:
                    padded = str(c)
                for n in self.normas:
                    nom_field = n.get('NOM', '')
                    if padded and padded in nom_field:
                        mapped_nom = nom_field
                        break
                    if str(c) and str(c) in nom_field:
                        mapped_nom = nom_field
                        break
                if mapped_nom:
                    mapped.append(mapped_nom)
                else:
                    mapped.append(str(c))
            nom_display = ", ".join(sorted(set(mapped))) if mapped else "N/A"

        # Para la columna DOCUMENTO EMITIDO preferimos usar un token informativo
        # extraído de `cadena_identificacion`. Normalmente la cadena contiene
        # un bloque "Solicitud de Servicio: <token>" que es el más completo. Si
        # no, tomamos el primer token (antes de '/') y, cuando disponemos de
        # `solicitud` y `folio`, anexamos la parte '<solicitud>-<folio>' para
        # obtener formatos como '25049USDNOM-141-SSA1007045-126'.
        if cadena_ident:
            try:
                token = None
                # Preferir token después de 'Solicitud de Servicio:' si existe
                msol = re.search(r"Solicitud de Servicio:\s*([A-Za-z0-9\-/]+)", str(cadena_ident))
                if msol:
                    token = msol.group(1).split('/')[0]
                else:
                    # Fallback: primer token antes de espacios, y quitar cualquier parte tras '/'
                    token = str(cadena_ident).strip().split()[0].split('/')[0]

                sol_part = self.extraer_sol_ema(solicitud) or ''
                # Build a sensible documento_emitido without duplicating fragments.
                if sol_part and folio:
                    suffix = f"{sol_part}-{folio}".lstrip('-')
                    # If token already contains the suffix or the sol_part, avoid appending again
                    if token and (token.endswith(suffix) or sol_part in token):
                        documento_emitido = token
                    else:
                        # ensure separator
                        sep = ''
                        if token and not token.endswith('-'):
                            sep = '-'
                        documento_emitido = f"{token}{sep}{suffix}" if token else f"{suffix}"
                else:
                    documento_emitido = token
            except Exception:
                documento_emitido = numero_solicitud_display
        else:
            documento_emitido = numero_solicitud_display

        # Normalizar modelos: mantener orden y eliminar duplicados vacíos
        modelos_norm = []
        for m in modelos:
            if not m:
                continue
            if m not in modelos_norm:
                modelos_norm.append(m)

        # Preferir FECHA DE ENTRADA desde backups (tabla_relacion_backups) si existe
        fecha_desaduanamiento = None
        try:
            fecha_desaduanamiento = self._lookup_backup_fecha(solicitud, folio)
            if not fecha_desaduanamiento:
                # buscar en registros cualquiera con FECHA DE ENTRADA válida
                for r in registros:
                    # Priorizar campos relacionados con pago de pedimento
                    candidates = [
                        r.get('FECHA DE PAGO PEDIMENTO'), r.get('FECHA PAGO PEDIMENTO'),
                        r.get('FECHA_PAGO_PEDIMENTO'), r.get('FECHA PAGO'),
                        r.get('FECHA DE ENTRADA'), r.get('FECHA_DE_ENTRADA'), r.get('FECHA_ENTRADA')
                    ]
                    for fe in candidates:
                        if fe:
                            fecha_desaduanamiento = fe
                            break
                    if fecha_desaduanamiento:
                        break
            # Si aún no hay fecha, intentar extraerla del JSON del dictamen si existe
            if not fecha_desaduanamiento and dictamen_json and isinstance(dictamen_json, dict):
                fjs = dictamen_json.get('fechas') or {}
                fecha_desaduanamiento = fjs.get('emision') or fjs.get('verificacion') or fjs.get('verificacion_larga')
        except Exception:
            fecha_desaduanamiento = None

        fila = {
            "NÚMERO DE SOLICITUD": numero_solicitud_display,
            "CLIENTE": cliente.get("CLIENTE", "N/A") if cliente else "N/A",
            "NÚMERO DE CONTRATO": cliente.get("NÚMERO_DE_CONTRATO", "N/A") if cliente else "N/A",
            "RFC": cliente.get("RFC", "N/A") if cliente else "N/A",
            "CURP": "N/A",
            "PRODUCTO VERIFICADO": ", ".join(descripciones) if descripciones else "N/A",
            "MARCAS": ", ".join(marcas) if marcas else "N/A",
            "NOM": nom_display,
            "TIPO DE DOCUMENTO OFICIAL EMITIDO": tipo_display,
            "DOCUMENTO EMITIDO": documento_emitido or "N/A",
            "FECHA DE DOCUMENTO EMITIDO": self._format_date(primer_registro.get("FECHA DE EMISION DE SOLICITUD", "N/A")),
            "VERIFICADOR": self._normalize_name(nombre_inspector),
            "PEDIMENTO DE IMPORTACION": primer_registro.get("PEDIMENTO", "N/A"),
            "FECHA DE DESADUANAMIENTO (CUANDO APLIQUE)": self._format_date(fecha_desaduanamiento or primer_registro.get("FECHA DE ENTRADA", "N/A")),
            "FECHA DE VISITA (CUANDO APLIQUE)": self._format_date(primer_registro.get("FECHA DE VERIFICACION", "N/A")),
            "MODELOS": ", ".join(modelos_norm) if modelos_norm else "N/A",
            "SOL EMA": self.extraer_sol_ema(solicitud),
            "FOLIO EMA": self.formatear_folio_ema(folio),
            "INSP EMA": self._normalize_name(nombre_inspector)
        }
        
        return fila
    
    def filtrar_por_fechas(self, fila: Dict, fecha_inicio: Optional[str] = None, 
                          fecha_fin: Optional[str] = None) -> bool:
        """
        Filtrar una fila por rango de fechas
        
        Args:
            fila: Fila de datos
            fecha_inicio: Fecha de inicio en formato YYYY-MM-DD
            fecha_fin: Fecha de fin en formato YYYY-MM-DD
            
        Returns:
            True si la fila está en el rango, False si no
        """
        if not fecha_inicio and not fecha_fin:
            return True

        # El rango se aplica sobre la fecha de documento emitido (emisión del dictamen)
        fecha_str = fila.get("FECHA DE DOCUMENTO EMITIDO", "")

        # Sin fecha legible no se puede ubicar en el rango: excluir el registro
        if not fecha_str or fecha_str == "N/A":
            return False

        try:
            # Intentar parsear la fecha en diferentes formatos
            for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"]:
                try:
                    fecha = datetime.strptime(str(fecha_str).strip(), fmt)
                    break
                except ValueError:
                    continue
            else:
                # Fecha mal capturada (ej. 22/05/20126): excluir el registro
                return False
            
            if fecha_inicio:
                inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d")
                if fecha < inicio:
                    return False
            
            if fecha_fin:
                fin = datetime.strptime(fecha_fin, "%Y-%m-%d")
                if fecha > fin:
                    return False
            
            return True

        except Exception:
            return False

    def crear_excel(self, nombre_archivo: str, fecha_inicio: Optional[str] = None,
                   fecha_fin: Optional[str] = None) -> Tuple[bool, str]:
        """
        Crear el archivo Excel con el control de folios
        
        Args:
            nombre_archivo: Nombre del archivo Excel a crear
            fecha_inicio: Fecha de inicio para filtrar (YYYY-MM-DD)
            fecha_fin: Fecha de fin para filtrar (YYYY-MM-DD)
            
        Returns:
            Tuple[bool, str]: (éxito, mensaje)
        """
        try:
            print("\n🚀 Generando archivo Excel...")

            # Libro en modo write_only: con decenas de miles de filas el modo
            # normal de openpyxl consume cientos de MB y es mucho más lento.
            wb = openpyxl.Workbook(write_only=True)
            ws = wb.create_sheet("Control de Folios")

            # Definir encabezados
            encabezados = [
                "NÚMERO DE SOLICITUD",
                "CLIENTE",
                "NÚMERO DE CONTRATO",
                "RFC",
                "CURP",
                "PRODUCTO VERIFICADO",
                "MARCAS",
                "NOM",
                "TIPO DE DOCUMENTO OFICIAL EMITIDO",
                "ESTATUS",
                "DOCUMENTO EMITIDO",
                "FECHA DE DOCUMENTO EMITIDO",
                "VERIFICADOR",
                "PEDIMENTO DE IMPORTACION",
                "FECHA DE DESADUANAMIENTO (CUANDO APLIQUE)",
                "FECHA DE VISITA (CUANDO APLIQUE)",
                "MODELOS",
                "SOL EMA",
                "FOLIO EMA",
                "INSP EMA"
            ]

            # En write_only, anchos y paneles deben definirse antes de escribir filas
            for col in range(1, len(encabezados) + 1):
                columna_letra = get_column_letter(col)
                # Anchos específicos según el contenido
                if col == 1:  # NÚMERO DE SOLICITUD
                    ws.column_dimensions[columna_letra].width = 18
                elif col in [2, 6, 16]:  # CLIENTE, PRODUCTO, MODELOS
                    ws.column_dimensions[columna_letra].width = 30
                elif col in [11, 14, 15]:  # FECHAS
                    ws.column_dimensions[columna_letra].width = 15
                else:
                    ws.column_dimensions[columna_letra].width = 20
            # Congelar primera fila
            ws.freeze_panes = 'A2'

            # Estilos (se crean una sola vez y se reutilizan en todas las celdas)
            borde = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            alin_encabezado = Alignment(horizontal="center", vertical="center", wrap_text=True)
            alin_celda = Alignment(horizontal="left", vertical="center", wrap_text=True)
            font_encabezado = Font(bold=True, color="FFFFFF")
            fill_encabezado = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            fill_cancelado = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            fill_pendiente = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            fill_completado = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font_negrita = Font(bold=True)
            font_normal = Font(bold=False)

            # Asignar fill/font/borde por celda obliga a openpyxl a hashear cada
            # estilo (lentísimo con millones de celdas). Se arma una celda
            # plantilla por estatus y se copia su arreglo de estilo ya indexado.
            def _estilo_plantilla(fill=None, font=None):
                plantilla = WriteOnlyCell(ws)
                plantilla.alignment = alin_celda
                plantilla.border = borde
                if fill is not None:
                    plantilla.fill = fill
                    plantilla.font = font
                return plantilla._style
            estilo_base = _estilo_plantilla()
            estilo_cancelado = _estilo_plantilla(fill_cancelado, font_negrita)
            estilo_pendiente = _estilo_plantilla(fill_pendiente, font_negrita)
            estilo_completado = _estilo_plantilla(fill_completado, font_normal)

            # Escribir encabezados
            fila_encabezados = []
            for encabezado in encabezados:
                celda = WriteOnlyCell(ws, value=encabezado)
                celda.font = font_encabezado
                celda.fill = fill_encabezado
                celda.alignment = alin_encabezado
                celda.border = borde
                fila_encabezados.append(celda)
            ws.append(fila_encabezados)

            # Reescribir la construcción de dictámenes para asegurar que
            # cada folio entero presente en `data/Dictamenes` o en
            # `tabla_de_relacion` se incluya una sola vez y en orden
            # numérico ascendente (a partir de folio 849).
            folio_map_int: Dict[int, Dict] = {}
            try:
                for entrada in self._cargar_indice_dictamenes():
                    if entrada.get('sub') != 'Dictamenes':
                        continue
                    fname = entrada.get('name', '')

                    # Default registro
                    registro = {
                        'FECHA DE EMISION DE SOLICITUD': None,
                        'FECHA DE VERIFICACION': None,
                        'PEDIMENTO': None,
                        'FIRMA': None,
                        'DESCRIPCION': None,
                        'MARCA': None,
                        'CODIGOS': [],
                    }

                    sol = ''
                    fol_int = None
                    d = None

                    if entrada.get('tipo') == 'json':
                        try:
                            d = entrada.get('d') if entrada.get('ok') else None
                            if d is None:
                                raise ValueError("JSON ilegible")
                            ident = d.get('identificacion', {})
                            sol = str(ident.get('solicitud') or '').strip()
                            fol_raw = ident.get('folio')
                            # intentar parsear folio desde identificacion.folio
                            try:
                                fol_int = int(float(str(fol_raw))) if fol_raw not in (None, '') else None
                            except Exception:
                                fol_int = None
                            # intentar extraer folio desde el nombre de archivo con patrón conocido
                            # muchos archivos tienen formato: Dictamen_Lista_<lista>_<folio>_<solicitud>_...
                            m_f = re.search(r"Dictamen_Lista_[^_]+_([0-9]{3,})_", fname)
                            if m_f:
                                try:
                                    fol_from_name = int(m_f.group(1))
                                    # priorizar el folio extraído del nombre de archivo cuando exista
                                    fol_int = fol_from_name
                                except Exception:
                                    pass

                            fechas = d.get('fechas', {})
                            producto = d.get('producto', {})
                            tabla_prod = d.get('tabla_productos', [])
                            firmas = d.get('firmas', {})

                            registro['FECHA DE EMISION DE SOLICITUD'] = fechas.get('emision')
                            registro['FECHA DE VERIFICACION'] = fechas.get('verificacion')
                            registro['PEDIMENTO'] = producto.get('pedimento')
                            if isinstance(firmas, dict):
                                f1 = firmas.get('firma1')
                                if f1:
                                    registro['FIRMA'] = f1.get('codigo_solicitado') or f1.get('codigo') or f1.get('nombre')
                            registro['DESCRIPCION'] = producto.get('descripcion')
                            if isinstance(tabla_prod, list) and tabla_prod:
                                registro['MARCA'] = tabla_prod[0].get('marca')
                                cods = []
                                for p in tabla_prod:
                                    c = p.get('codigo')
                                    if c is not None:
                                        cods.append(str(c))
                                registro['CODIGOS'] = cods
                        except Exception:
                            # Si falla la carga JSON, intentar extraer folio desde el nombre
                            fol_int = None
                    else:
                        # PDF: extraer la mayor secuencia de dígitos como folio probable
                        name_base = os.path.splitext(fname)[0]
                        # Preferir patrón de nombre similar a JSON files
                        m_f = re.search(r"Dictamen_Lista_[^_]+_([0-9]{3,})_", fname)
                        if m_f:
                            try:
                                fol_int = int(m_f.group(1))
                            except Exception:
                                fol_int = None
                        else:
                            seqs = re.findall(r"(\d+)", name_base)
                            if seqs:
                                try:
                                    fol_int = int(max(seqs, key=lambda s: len(s)))
                                except Exception:
                                    fol_int = None
                        m = re.search(r"([A-Za-z0-9\-]{4,})", name_base)
                        if m:
                            sol = m.group(1)

                    if fol_int is None:
                        # si no pudimos extraer folio numérico, saltar
                        continue

                    # Insertar o fusionar en folio_map_int
                    if fol_int not in folio_map_int:
                        folio_map_int[fol_int] = {
                            'solicitud': sol,
                            'folio': fol_int,
                            'registros': [registro],
                            'dictamen_json': d
                        }
                    else:
                        # fusionar registros y preferir dictamen_json no nulo
                        existing = folio_map_int[fol_int]
                        if registro not in existing.get('registros', []):
                            existing.setdefault('registros', []).append(registro)
                        if not existing.get('dictamen_json') and d:
                            existing['dictamen_json'] = d
            except Exception as e:
                print(f"⚠️ Error leyendo data/Dictamenes: {e}")

            # Añadir folios presentes en tabla_de_relacion que no estén en Dictamenes
            try:
                for registro in self.tabla_relacion:
                    fol = registro.get('FOLIO') or registro.get('folio')
                    try:
                        fol_i = int(float(str(fol)))
                    except Exception:
                        continue
                    if fol_i not in folio_map_int:
                        folio_map_int[fol_i] = {
                            'solicitud': registro.get('SOLICITUD') or registro.get('solicitud') or '',
                            'folio': fol_i,
                            'registros': [registro],
                            'dictamen_json': None
                        }
                    else:
                        # anexamos registro si no está
                        existing = folio_map_int[fol_i]
                        if registro not in existing.get('registros', []):
                            existing.setdefault('registros', []).append(registro)
            except Exception:
                pass

            candidates = [k for k in folio_map_int.keys() if k >= 849]
            if not candidates:
                candidates = list(folio_map_int.keys())
            all_folios_sorted = sorted(candidates)
            print(f"📊 Folios detectados para exportar: {len(all_folios_sorted)} (threshold applied: {any(k>=849 for k in candidates)})")

            # Generar filas: iterar por folio entero ordenado para evitar saltos
            filas_procesadas = 0

            # Construir mapa de estatus por folio a partir del historial (pendientes/cancelados)
            folio_status_map = {}
            try:
                for visita in self.historial_visitas:
                    est = str(visita.get('estatus', '')).strip().lower()
                    folios_str = str(visita.get('folios_utilizados', '')).strip()
                    if not folios_str:
                        continue
                    if ' - ' in folios_str:
                        parts = folios_str.split(' - ')
                        if len(parts) == 2:
                            try:
                                inicio = int(parts[0].strip())
                                fin = int(parts[1].strip())
                                for f in range(inicio, fin + 1):
                                    folio_status_map[f] = est
                            except Exception:
                                pass
                    else:
                        for part in [p.strip() for p in folios_str.split(',') if p.strip()]:
                            try:
                                folio_status_map[int(part)] = est
                            except Exception:
                                digits = ''.join(ch for ch in part if ch.isdigit())
                                if digits:
                                    try:
                                        folio_status_map[int(digits)] = est
                                    except Exception:
                                        pass
            except Exception:
                folio_status_map = {}

            total_folios = len(all_folios_sorted)
            for i, fol in enumerate(all_folios_sorted, 1):
                if i % 200 == 0 or i == total_folios:
                    self._reportar_progreso(i, total_folios, f"Generando filas ({i:,}/{total_folios:,})")

                dictamen = folio_map_int.get(fol)
                if not dictamen:
                    continue

                fila_datos = self.generar_fila_excel(dictamen)

                # Filtrar por fechas si se especificaron
                if not self.filtrar_por_fechas(fila_datos, fecha_inicio, fecha_fin):
                    continue

                # Determinar estatus (si existe en historial)
                est = folio_status_map.get(int(fol), '') if fol is not None else ''
                est_display = est.upper() if est else 'COMPLETADO'

                # Color según estatus (más visible)
                if est and est in ('cancelado', 'cancelada'):
                    estilo = estilo_cancelado
                elif est and 'pend' in est:
                    estilo = estilo_pendiente
                elif not est:
                    # considerar completado como verde
                    estilo = estilo_completado
                else:
                    estilo = estilo_base

                # Escribir datos (incluir ESTATUS)
                fila_celdas = []
                for encabezado in encabezados:
                    if encabezado == 'ESTATUS':
                        valor = est_display
                    else:
                        valor = fila_datos.get(encabezado, "N/A")
                    celda = WriteOnlyCell(ws, value=valor)
                    celda._style = copy(estilo)
                    fila_celdas.append(celda)
                ws.append(fila_celdas)

                filas_procesadas += 1

            # Guardar archivo
            self._reportar_progreso(total_folios, total_folios, "Guardando archivo Excel...")
            wb.save(nombre_archivo)

            mensaje = f"✅ Archivo Excel generado exitosamente: {nombre_archivo}\n"
            mensaje += f"   📊 Total de registros: {filas_procesadas}"
            
            if fecha_inicio or fecha_fin:
                mensaje += f"\n   📅 Rango de fechas aplicado: "
                mensaje += f"{fecha_inicio or 'inicio'} a {fecha_fin or 'fin'}"
            
            print(mensaje)
            return True, mensaje
            
        except Exception as e:
            mensaje = f"❌ Error al crear archivo Excel: {e}"
            print(mensaje)
            return False, mensaje

def main():
    """Función principal para ejecutar el script"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Generar Control de Folios Anual desde archivos JSON"
    )
    parser.add_argument(
        "--output",
        "-o",
        default="Control_Folios_Anual.xlsx",
        help="Nombre del archivo Excel de salida (default: Control_Folios_Anual.xlsx)"
    )
    parser.add_argument(
        "--fecha-inicio",
        "-fi",
        help="Fecha de inicio para filtrar (formato: YYYY-MM-DD)"
    )
    parser.add_argument(
        "--fecha-fin",
        "-ff",
        help="Fecha de fin para filtrar (formato: YYYY-MM-DD)"
    )
    parser.add_argument(
        "--data-dir",
        "-d",
        default="data",
        help="Directorio donde se encuentran los archivos JSON (default: data)"
    )
    
    args = parser.parse_args()
    
    print("=" * 70)
    print("📊 GENERADOR DE CONTROL DE FOLIOS ANUAL")
    print("=" * 70)
    print()
    
    # Crear instancia del generador
    generador = ControlFoliosAnual(data_dir=args.data_dir)
    
    # Cargar datos
    print("📂 Cargando datos desde archivos JSON...")
    exito, mensaje = generador.cargar_datos()
    
    if not exito:
        print(f"\n❌ Error: {mensaje}")
        return 1
    
    print()
    
    # Generar Excel
    exito, mensaje = generador.crear_excel(
        args.output,
        fecha_inicio=args.fecha_inicio,
        fecha_fin=args.fecha_fin
    )
    
    if not exito:
        print(f"\n❌ Error: {mensaje}")
        return 1
    
    print()
    print("=" * 70)
    print("✅ PROCESO COMPLETADO")
    print("=" * 70)
    
    return 0

def generar_control_folios_anual(
    historial_path,
    tabla_backups_dir,
    output_path,
    year,
    start_date=None,
    end_date=None,
    export_cache=None,
    historial_list: Optional[List[Dict]] = None,
    data_dir: Optional[str] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None
):
    from datetime import datetime
    import os

    def normalizar(fecha):
        if not fecha:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(fecha.strip(), fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        # Antes se ignoraba en silencio y el reporte salía sin filtrar
        raise Exception(f"Fecha inválida en el rango: '{fecha}'. Use el formato dd/mm/aaaa.")

    fecha_inicio = normalizar(start_date)
    fecha_fin = normalizar(end_date)

    # DEBUG: imprimir parámetros recibidos
    try:
        print(f"[DEBUG] generar_control_folios_anual called with:\n  historial_path={historial_path}\n  tabla_backups_dir={tabla_backups_dir}\n  output_path={output_path}\n  year={year}\n  start_date={start_date} -> {fecha_inicio}\n  end_date={end_date} -> {fecha_fin}\n  export_cache={export_cache}\n")
    except Exception:
        pass

    if not data_dir:
        base_dir = os.path.dirname(os.path.dirname(historial_path))
        data_dir = os.path.join(base_dir, "data")

    print(f"[DEBUG] Resolved data_dir: {data_dir}")

    generador = ControlFoliosAnual(data_dir=data_dir)
    generador.progress_cb = progress_cb

    exito, mensaje = generador.cargar_datos()
    if not exito:
        raise Exception(mensaje)

    if historial_list is not None:
        # historial_list puede venir como {'visitas': [...]} o directamente como lista
        if isinstance(historial_list, dict) and 'visitas' in historial_list:
            generador.historial_visitas = historial_list['visitas']
        elif isinstance(historial_list, list):
            generador.historial_visitas = historial_list
        else:
            # ignorar si no tiene formato esperado
            pass
        # recrear mapeo
        try:
            generador._crear_mapeo_folio_cliente()
        except Exception:
            pass

    exito, mensaje = generador.crear_excel(
        output_path,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin
    )

    if not exito:
        raise Exception(mensaje)

    return True


def generar_reporte_ema(tabla_de_relacion_path, historial_path, output_path, export_cache=None):
    """Genera el reporte EMA a partir de un archivo de tabla_de_relacion (o lista JSON).

    Args:
        tabla_de_relacion_path: Ruta al JSON de tabla_de_relacion o a un JSON temporal con registros.
        historial_path: Ruta al historial (se usa para resolver data_dir y clientes si es necesario).
        output_path: Ruta de salida .xlsx
        export_cache: (opcional) ruta a cache de export
    """
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    # Resolver data_dir igual que en generar_control_folios_anual
    base_dir = os.path.dirname(os.path.dirname(historial_path))
    data_dir = os.path.join(base_dir, "data")

    generador = ControlFoliosAnual(data_dir=data_dir)
    exito, mensaje = generador.cargar_datos()
    if not exito:
        raise Exception(mensaje)

    # Cargar tabla_de_relacion
    if not os.path.exists(tabla_de_relacion_path):
        raise Exception(f"No se encontró tabla_de_relacion: {tabla_de_relacion_path}")

    try:
        with open(tabla_de_relacion_path, 'r', encoding='utf-8') as f:
            tabla_obj = json.load(f)
    except Exception as e:
        raise Exception(f"Error leyendo {tabla_de_relacion_path}: {e}")

    # Normalizar a lista de registros
    if isinstance(tabla_obj, dict):
        # Si es dict y contiene una lista en alguna clave esperada
        if 'tabla' in tabla_obj and isinstance(tabla_obj['tabla'], list):
            registros = tabla_obj['tabla']
        elif 'registros' in tabla_obj and isinstance(tabla_obj['registros'], list):
            registros = tabla_obj['registros']
        else:
            # si es un dict que ya representa una lista envuelta
            registros = []
            for v in tabla_obj.values():
                if isinstance(v, list):
                    registros = v
                    break
    elif isinstance(tabla_obj, list):
        registros = tabla_obj
    else:
        raise Exception("Formato de tabla_de_relacion no reconocido")

    # Agrupar por SOLICITUD + FOLIO
    grupos = {}
    for reg in registros:
        solicitud = reg.get('SOLICITUD', '')
        folio = reg.get('FOLIO', '')
        clave = f"{solicitud}_{folio}"
        grupos.setdefault(clave, []).append(reg)

    # Preparar workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "EMA"

    encabezados = [
        "Número de solicitud",
        "Fecha de inspección",
        "Número de dictamen",
        "ESTATUS",
        "Número de Contrato",
        "Tipo de Documento Oficial Emitido",
        "Fecha de Documento Emitido",
        "Producto verificado",
        "Fecha de Desaduanamiento",
        "Fecha de visita",
        "Observaciones",
        "Inspector(es)",
        "Persona(s) de apoyo",
        "NOM"
    ]

    # Escribir encabezados con estilo
    for col, h in enumerate(encabezados, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    fila = 2
    filas_procesadas = 0

    # Construir mapa de estatus por folio a partir del historial (pendientes/cancelados)
    folio_status_map = {}
    try:
        for visita in generador.historial_visitas:
            est = str(visita.get('estatus', '')).strip().lower()
            folios_str = str(visita.get('folios_utilizados', '')).strip()
            if not folios_str:
                continue
            # Rango
            if ' - ' in folios_str:
                parts = folios_str.split(' - ')
                if len(parts) == 2:
                    try:
                        inicio = int(parts[0].strip())
                        fin = int(parts[1].strip())
                        for f in range(inicio, fin + 1):
                            folio_status_map[f] = est
                    except Exception:
                        pass
            else:
                # Puede ser lista separada por comas o único folio
                for part in [p.strip() for p in folios_str.split(',') if p.strip()]:
                    try:
                        folio_status_map[int(part)] = est
                    except Exception:
                        # intentar extraer dígitos del token (ej. CP000001 -> 000001)
                        digits = ''.join(ch for ch in part if ch.isdigit())
                        if digits:
                            try:
                                folio_status_map[int(digits)] = est
                            except Exception:
                                pass
                        pass
    except Exception:
        folio_status_map = {}

    # Iterar grupos ordenados por número de folio para detectar saltos
    sorted_items = []
    for clave, regs in grupos.items():
        # clave = f"{solicitud}_{folio}"
        try:
            fol = clave.split('_')[-1]
            # extraer dígitos del folio para obtener número entero
            fol_digits = ''.join(ch for ch in str(fol) if ch.isdigit())
            foln = int(fol_digits) if fol_digits else 0
        except Exception:
            try:
                foln = int(regs[0].get('FOLIO', 0))
            except Exception:
                foln = 0
        sorted_items.append((foln, clave, regs))
    sorted_items.sort(key=lambda x: x[0])

    prev_folio_num = None
    for foln, clave, regs in sorted_items:
        primer = regs[0]
        solicitud = primer.get('SOLICITUD', '')
        folio = primer.get('FOLIO', '')

        # Formatos y búsquedas
        try:
            # normalizar folio a entero extrayendo dígitos (soporta prefijos como CP000001)
            fol_digits = ''.join(ch for ch in str(folio) if ch.isdigit())
            folio_num = int(fol_digits) if fol_digits else 0
        except Exception:
            folio_num = 0

        cliente_info = generador.buscar_cliente_por_solicitud(solicitud, folio_num)

        # Construir campos
        numero_solicitud = generador.extraer_sol_ema(solicitud)
        fecha_inspeccion = generador._format_date(primer.get('FECHA DE VERIFICACION', 'N/A'))
        numero_dictamen = generador.formatear_folio_ema(folio)
        numero_contrato = cliente_info.get('NÚMERO_DE_CONTRATO', 'N/A') if cliente_info else 'N/A'
        tipo_raw = primer.get('TIPO DE DOCUMENTO', primer.get('TIPO DE DOCUMENTO OFICIAL EMITIDO', 'D'))
        _tt = str(tipo_raw).strip().upper() if tipo_raw is not None else ''
        if _tt == 'D':
            tipo_doc = 'Dictamen'
        elif _tt == 'C':
            tipo_doc = 'Constancia'
        else:
            tipo_doc = str(tipo_raw)
        fecha_doc_emitido = generador._format_date(primer.get('FECHA DE EMISION DE SOLICITUD', 'N/A'))

        # Productos, noms
        productos = set()
        noms = set()
        for r in regs:
            if r.get('DESCRIPCION'):
                productos.add(r.get('DESCRIPCION'))
            if r.get('CLASIF UVA'):
                noms.add(str(r.get('CLASIF UVA')))

        producto_verificado = ", ".join(productos) if productos else 'N/A'
        fecha_desaduanamiento = generador._format_date(primer.get('FECHA DE ENTRADA', 'N/A'))
        fecha_visita = generador._format_date(primer.get('FECHA DE VERIFICACION', 'N/A'))
        observaciones = 'N/A'

        # Inspector(es)
        firma = primer.get('FIRMA', '')
        inspector_nombre = generador.buscar_inspector_por_firma(firma)
        inspector_nombre = generador._normalize_name(inspector_nombre)

        personas_apoyo = 'N/A'
        nom_str = ", ".join(noms) if noms else 'N/A'

        # Determinar estatus desde el mapeo (si existe)
        est = folio_status_map.get(folio_num, '')
        est_display = est.upper() if est else 'COMPLETADO'

        fila_vals = [
            numero_solicitud,
            fecha_inspeccion,
            numero_dictamen,
            est_display,
            numero_contrato,
            tipo_doc,
            fecha_doc_emitido,
            producto_verificado,
            fecha_desaduanamiento,
            fecha_visita,
            observaciones,
            inspector_nombre,
            personas_apoyo,
            nom_str
        ]

        # Escribir fila
        for col, val in enumerate(fila_vals, 1):
            cel = ws.cell(row=fila, column=col, value=val)
            cel.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cel.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # Detectar salto (gap) con respecto al folio previo
        is_salto = False
        if prev_folio_num is not None and folio_num and (folio_num - prev_folio_num) > 1:
            is_salto = True

        # Detectar folio cancelado desde historial (solo 'cancelado'/'cancelada')
        est = folio_status_map.get(folio_num, '')
        is_cancelado = str(est).strip().lower() in ('cancelado', 'cancelada')

        # Aplicar color únicamente para folios cancelados (rojo). El resto sin color.
        if is_cancelado:
            fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            # Marcar negrita y aplicar relleno sólo en caso de cancelado
            for col in range(1, len(fila_vals) + 1):
                try:
                    cell = ws.cell(row=fila, column=col)
                    cell.fill = fill
                    cell.font = Font(bold=True)
                except Exception:
                    pass

        prev_folio_num = folio_num
        fila += 1
        filas_procesadas += 1

    # Ajustar anchos de columna
    for col in range(1, len(encabezados) + 1):
        col_letter = get_column_letter(col)
        if col in (1, 3, 4):
            ws.column_dimensions[col_letter].width = 18
        elif col in (2, 7):
            ws.column_dimensions[col_letter].width = 30
        else:
            ws.column_dimensions[col_letter].width = 20

    ws.freeze_panes = 'A2'

    try:
        wb.save(output_path)
    except Exception as e:
        raise Exception(f"Error guardando Excel EMA: {e}")

    return True

if __name__ == "__main__":
    exit(main())

