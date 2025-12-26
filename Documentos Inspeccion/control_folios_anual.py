import json
import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
        self._dictamen_cache = {}
        
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
                with open(clientes_path, 'r', encoding='utf-8') as f:
                    self.clientes = json.load(f)
                print(f"✅ Clientes cargados: {len(self.clientes)} registros")
            else:
                return False, f"No se encontró {clientes_path}"
            
            # Cargar Firmas.json
            firmas_path = os.path.join(self.data_dir, "Firmas.json")
            if os.path.exists(firmas_path):
                with open(firmas_path, 'r', encoding='utf-8') as f:
                    self.firmas = json.load(f)
                print(f"✅ Firmas cargadas: {len(self.firmas)} registros")
            else:
                return False, f"No se encontró {firmas_path}"
            
            # Cargar tabla_de_relacion.json
            tabla_path = os.path.join(self.data_dir, "tabla_de_relacion.json")
            if os.path.exists(tabla_path):
                with open(tabla_path, 'r', encoding='utf-8') as f:
                    self.tabla_relacion = json.load(f)
                print(f"✅ Tabla de relación cargada: {len(self.tabla_relacion)} registros")
            else:
                return False, f"No se encontró {tabla_path}"

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
                with open(historial_path, 'r', encoding='utf-8') as f:
                    hist_data = json.load(f)
                    if isinstance(hist_data, dict) and 'visitas' in hist_data:
                        self.historial_visitas = hist_data['visitas']
                        # Crear mapeo de folio a cliente
                        self._crear_mapeo_folio_cliente()
                        print(f"✅ Historial de visitas cargado: {len(self.historial_visitas)} registros")
            
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
                return inspector.get("NOMBRE DE INSPECTOR", "N/A")
        return "N/A"
    
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

    def _find_dictamen(self, solicitud: str, folio) -> Optional[Dict]:
        """Buscar dictamen JSON en data/Dictamenes que coincida con solicitud y folio."""
        try:
            # Normalizar folio y solicitud para mejorar coincidencias
            folio_s = str(folio)
            sol_search = str(solicitud) if solicitud is not None else ''
            # Si viene con formato 'XXXX/25', usar la parte antes de la barra
            if '/' in sol_search:
                sol_base = sol_search.split('/')[0]
            else:
                sol_base = sol_search
            dicts_dir = os.path.join(self.data_dir, 'Dictamenes')
            if dicts_dir in self._dictamen_cache:
                files = self._dictamen_cache[dicts_dir]
            else:
                if not os.path.exists(dicts_dir):
                    return None
                files = [os.path.join(dicts_dir, f) for f in os.listdir(dicts_dir) if f.lower().endswith('.json')]
                self._dictamen_cache[dicts_dir] = files

            for fp in files:
                try:
                    with open(fp, 'r', encoding='utf-8') as f:
                        d = json.load(f)
                    ident = d.get('identificacion', {})
                    sol = ident.get('solicitud') or ''
                    fol = str(ident.get('folio', ''))
                    cadena = ident.get('cadena_identificacion', '') or ''

                    # Comparar por folio cuando esté presente y coincida
                    if fol and folio and fol == folio_s:
                        return d

                    # Comparar por solicitud exacta o por base (sin /aa)
                    try:
                        if sol and sol_search and str(sol).endswith(sol_search):
                            return d
                    except Exception:
                        pass

                    try:
                        if sol and sol_base and str(sol).endswith(sol_base):
                            return d
                    except Exception:
                        pass

                    # Buscar en cadena_identificacion la base o la solicitud completa
                    if cadena and sol_search and sol_search in cadena:
                        return d
                    if cadena and sol_base and sol_base in cadena:
                        return d
                except Exception:
                    continue
        except Exception:
            return None
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

        # Intentar localizar dictamen JSON para extraer cadena_identificacion y norma completa
        dictamen_json = self._find_dictamen(solicitud, folio)
        cadena_ident = None
        norma_codigo = None
        if dictamen_json:
            ident = dictamen_json.get('identificacion', {})
            cadena_ident = ident.get('cadena_identificacion')
            norma = dictamen_json.get('norma', {})
            norma_codigo = norma.get('codigo') or norma.get('NOM')
        
        # Obtener información del inspector
        firma = primer_registro.get("FIRMA", "")
        nombre_inspector = self.buscar_inspector_por_firma(firma)
        
        # Extraer descripciones, marcas, NOMs y modelos de todos los registros
        descripciones = set()
        marcas = set()
        noms = set()
        modelos = []
        
        for reg in registros:
            if reg.get("DESCRIPCION"):
                descripciones.add(reg.get("DESCRIPCION"))
            if reg.get("MARCA"):
                marcas.add(reg.get("MARCA"))
            if reg.get("CLASIF UVA"):
                noms.add(str(reg.get("CLASIF UVA")))
            if reg.get("CODIGO"):
                modelos.append(str(reg.get("CODIGO")))
        
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
        tipo_display = "Dictamen" if str(tipo_raw).strip().upper() == 'D' else str(tipo_raw)

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

        documento_emitido = numero_solicitud_display

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
            "FECHA DE DOCUMENTO EMITIDO": primer_registro.get("FECHA DE EMISION DE SOLICITUD", "N/A"),
            "VERIFICADOR": nombre_inspector,
            "PEDIMENTO DE IMPORTACION": primer_registro.get("PEDIMENTO", "N/A"),
            "FECHA DE DESADUANAMIENTO (CUANDO APLIQUE)": primer_registro.get("FECHA DE ENTRADA", "N/A"),
            "FECHA DE VISITA (CUANDO APLIQUE)": primer_registro.get("FECHA DE VERIFICACION", "N/A"),
            "MODELOS": ", ".join(modelos) if modelos else "N/A",
            "SOL EMA": self.extraer_sol_ema(solicitud),
            "FOLIO EMA": self.formatear_folio_ema(folio),
            "INSP EMA": nombre_inspector
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
        
        # Usar la fecha de verificación para filtrar
        fecha_str = fila.get("FECHA DE VISITA (CUANDO APLIQUE)", "")
        
        if not fecha_str or fecha_str == "N/A":
            return False
        
        try:
            # Intentar parsear la fecha en diferentes formatos
            for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"]:
                try:
                    fecha = datetime.strptime(fecha_str, fmt)
                    break
                except ValueError:
                    continue
            else:
                # Si no se pudo parsear, incluir el registro
                return True
            
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
            # En caso de error, incluir el registro
            return True
    
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
            
            # Crear libro de trabajo
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Control de Folios"
            
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
            
            # Escribir encabezados
            for col, encabezado in enumerate(encabezados, 1):
                celda = ws.cell(row=1, column=col, value=encabezado)
                # Estilo para encabezados
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                celda.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            
            # Agrupar datos por dictamen
            dictamenes = self.agrupar_por_dictamen()
            print(f"📊 Dictámenes encontrados: {len(dictamenes)}")
            
            # Generar filas
            fila_actual = 2
            filas_procesadas = 0
            
            for dictamen in dictamenes:
                fila_datos = self.generar_fila_excel(dictamen)
                
                # Filtrar por fechas si se especificaron
                if not self.filtrar_por_fechas(fila_datos, fecha_inicio, fecha_fin):
                    continue
                
                # Escribir datos
                for col, encabezado in enumerate(encabezados, 1):
                    valor = fila_datos.get(encabezado, "N/A")
                    celda = ws.cell(row=fila_actual, column=col, value=valor)
                    celda.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    celda.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                
                fila_actual += 1
                filas_procesadas += 1
            
            # Ajustar ancho de columnas
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
            
            # Guardar archivo
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
    export_cache=None
):
    from datetime import datetime
    import os

    def normalizar(fecha):
        if not fecha:
            return None
        try:
            return datetime.strptime(fecha, "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            return None

    fecha_inicio = normalizar(start_date)
    fecha_fin = normalizar(end_date)

    # 👉 Corrección CLAVE
    base_dir = os.path.dirname(os.path.dirname(historial_path))
    data_dir = os.path.join(base_dir, "data")

    generador = ControlFoliosAnual(data_dir=data_dir)

    exito, mensaje = generador.cargar_datos()
    if not exito:
        raise Exception(mensaje)

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

    for clave, regs in grupos.items():
        primer = regs[0]
        solicitud = primer.get('SOLICITUD', '')
        folio = primer.get('FOLIO', '')

        # Formatos y búsquedas
        try:
            folio_num = int(folio) if folio not in (None, '') else 0
        except Exception:
            folio_num = 0

        cliente_info = generador.buscar_cliente_por_solicitud(solicitud, folio_num)

        # Construir campos
        numero_solicitud = generador.extraer_sol_ema(solicitud)
        fecha_inspeccion = primer.get('FECHA DE VERIFICACION', 'N/A')
        numero_dictamen = generador.formatear_folio_ema(folio)
        numero_contrato = cliente_info.get('NÚMERO_DE_CONTRATO', 'N/A') if cliente_info else 'N/A'
        tipo_doc = primer.get('TIPO DE DOCUMENTO', primer.get('TIPO DE DOCUMENTO OFICIAL EMITIDO', 'D'))
        fecha_doc_emitido = primer.get('FECHA DE EMISION DE SOLICITUD', 'N/A')

        # Productos, noms
        productos = set()
        noms = set()
        for r in regs:
            if r.get('DESCRIPCION'):
                productos.add(r.get('DESCRIPCION'))
            if r.get('CLASIF UVA'):
                noms.add(str(r.get('CLASIF UVA')))

        producto_verificado = ", ".join(productos) if productos else 'N/A'
        fecha_desaduanamiento = primer.get('FECHA DE ENTRADA', 'N/A')
        fecha_visita = primer.get('FECHA DE VERIFICACION', 'N/A')
        observaciones = 'N/A'

        # Inspector(es)
        firma = primer.get('FIRMA', '')
        inspector_nombre = generador.buscar_inspector_por_firma(firma)

        personas_apoyo = 'N/A'
        nom_str = ", ".join(noms) if noms else 'N/A'

        fila_vals = [
            numero_solicitud,
            fecha_inspeccion,
            numero_dictamen,
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

        for col, val in enumerate(fila_vals, 1):
            cel = ws.cell(row=fila, column=col, value=val)
            cel.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cel.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

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

