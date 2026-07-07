"""plantilla.py - Funciones de carga y preparación de datos"""
import pandas as pd
import json
from datetime import datetime
from collections import defaultdict
import os
import sys
import traceback
from etiqueta_dictamen import GeneradorEtiquetasDecathlon

# ---------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------
def obtener_ruta_recurso(ruta_relativa):
    """
    Obtiene la ruta absoluta del recurso, funciona tanto para .py como para .exe
    PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
    """
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        ruta_base = sys._MEIPASS
    except Exception:
        # Si no existe _MEIPASS, estamos corriendo como .py normal
        ruta_base = os.path.abspath(".")

    # Si el usuario o la app definieron un directorio de datos, comprobarlo primero
    try:
        env_data = os.environ.get('IMAGENESVC_DATA_DIR') or os.environ.get('FOLIO_DATA_DIR')
        if env_data:
            # Si la ruta relativa empieza por 'data/', ajustamos la cola
            tail = ruta_relativa
            if ruta_relativa.startswith('data' + os.path.sep) or ruta_relativa.startswith('data/'):
                tail = ruta_relativa.split(os.path.sep, 1)[-1] if os.path.sep in ruta_relativa else ruta_relativa.split('/', 1)[-1]
            candidate = os.path.join(env_data, tail)
            if os.path.exists(candidate):
                return candidate
    except Exception:
        pass

    # Preferir carpeta `data` ubicada junto al ejecutable (APP_DIR) si existe.
    try:
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.abspath(".")

        candidates = [
            os.path.join(app_dir, ruta_relativa),
            os.path.join(app_dir, '_internal', ruta_relativa),
            os.path.join(ruta_base, ruta_relativa),
            os.path.join(ruta_base, '_internal', ruta_relativa),
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    except Exception:
        pass

    return os.path.join(ruta_base, ruta_relativa)

# ---------------------------------------------------------
# FORMATEADORES DE FECHA
# ---------------------------------------------------------
def formatear_fecha_larga(fecha_str):
    if pd.isna(fecha_str) or fecha_str == "":
        return ""
    try:
        if isinstance(fecha_str, str) and "/" in fecha_str:
            fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
        else:
            fecha = datetime.strptime(str(fecha_str), "%Y-%m-%d")
        meses = {
            1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
            5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
            9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
        }
        return f"{fecha.day} de {meses[fecha.month]} de {fecha.year}"
    except:
        return str(fecha_str)

# ---------------------------------------------------------
# CARGA DE ARCHIVOS
# ---------------------------------------------------------
def cargar_tabla_relacion(ruta="data/tabla_de_relacion.json"):
    try:
        ruta_completa = obtener_ruta_recurso(ruta)
        with open(ruta_completa, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Detectar si el JSON cargado parece ser un "índice" (mapping code->destino)
        is_index_like = False
        try:
            if isinstance(data, dict):
                # verificar que todos los valores sean strings (destinos)
                if all(isinstance(v, str) or v is None for v in data.values()):
                    is_index_like = True
        except Exception:
            is_index_like = False

        if is_index_like:
            # Respaldar el archivo erróneo para análisis
            try:
                import shutil, time
                backup_path = ruta_completa + f".index_error_backup_{int(time.time())}.json"
                shutil.copy2(ruta_completa, backup_path)
                print(f"⚠️ Atención: {ruta} contiene un índice (code->destino). Se creó respaldo en: {backup_path}")
            except Exception:
                print(f"⚠️ Atención: {ruta} parece contener un índice (code->destino). No se pudo crear respaldo automático.")

            # Intentar restaurar desde el backup más reciente en data/tabla_relacion_backups/
            try:
                base_dir = os.path.dirname(ruta_completa)
                backups_dir = os.path.join(base_dir, 'tabla_relacion_backups')
                if os.path.isdir(backups_dir):
                    files = [os.path.join(backups_dir, f) for f in os.listdir(backups_dir) if f.lower().endswith('.json')]
                    if files:
                        files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                        latest = files[0]
                        shutil.copy2(latest, ruta_completa)
                        with open(ruta_completa, 'r', encoding='utf-8') as f2:
                            data = json.load(f2)
                        df = pd.DataFrame(data)
                        print(f"✅ Restaurado backup desde {latest}. Tabla de relación cargada: {len(df)} registros")
                        return df
            except Exception as e:
                print(f"⚠️ No se pudo restaurar backup automáticamente: {e}")

            # Si no hay backup o no se pudo restaurar, devolver DataFrame vacío y notificar
            print("❌ El archivo de tabla de relación parece incorrecto (índice). Por favor restaure el archivo correcto en 'data/tabla_de_relacion.json' o use la opción 'Cargar Tabla de Relación' con el archivo apropiado.")
            return pd.DataFrame()

        df = pd.DataFrame(data)
        print(f"✅ Tabla de relación cargada: {len(df)} registros")
        return df
    except Exception as e:
        print(f"❌ Error cargando tabla de relación: {e}")
        return pd.DataFrame()

def cargar_normas(ruta="data/Normas.json"):
    """Carga las NOMs y las indexa usando el número inicial (ej: 24, 50, 141)."""
    try:
        ruta_completa = obtener_ruta_recurso(ruta)
        with open(ruta_completa, "r", encoding="utf-8") as f:
            normas = json.load(f)

        normas_map = {}
        normas_info = {}

        for norma in normas:
            nom = norma.get("NOM", "").strip()
            nombre = norma.get("NOMBRE", "").strip()
            capitulo = norma.get("CAPITULO", "").strip()

            if not nom:
                continue

            try:
                numero_nom = nom.split("-")[1]
                numero_nom = str(int(numero_nom))
            except:
                continue

            normas_map[numero_nom] = nom
            normas_info[nom] = {
                "nombre": nombre,
                "capitulo": capitulo
            }

        print(f"✅ Normas cargadas correctamente: {len(normas_map)} mapeos")
        return normas_map, normas_info

    except Exception as e:
        print(f"❌ Error cargando NOMs: {e}")
        return {}, {}

def cargar_clientes(ruta="data/Clientes.json"):
    try:
        ruta_completa = obtener_ruta_recurso(ruta)
        with open(ruta_completa, "r", encoding="utf-8") as f:
            clientes = json.load(f)

        clientes_map = {}
        for cliente in clientes:
            marca = cliente.get("marca", "").strip().upper()
            if marca:
                clientes_map[marca] = {
                    "nombre": cliente.get("nombre", ""),
                    "rfc": cliente.get("rfc", "")
                }

        print(f"✅ Clientes cargados: {len(clientes_map)}")
        return clientes_map

    except Exception as e:
        print(f"⚠️ No se pudo cargar {ruta}: {e}")
        return {}

def cargar_firmas(ruta="data/Firmas.json"):
    """
    Carga el mapeo completo de firmas de inspectores desde Firmas.json.
    Incluye: nombre, imagen, normas acreditadas, puesto, etc.
    Indexado por código FIRMA para búsqueda rápida.
    """
    try:
        ruta_completa = obtener_ruta_recurso(ruta)
        with open(ruta_completa, "r", encoding="utf-8") as f:
            firmas = json.load(f)
        
        firmas_map = {}
        for firma in firmas:
            codigo = firma.get("FIRMA", "").strip()
            if codigo:
                firmas_map[codigo] = {
                    "nombre": firma.get("NOMBRE DE INSPECTOR", "").strip(),
                    "imagen": firma.get("IMAGEN", "").strip(),
                    "puesto": firma.get("Puesto", "").strip(),
                    "normas_acreditadas": firma.get("Normas acreditadas", []),
                    "vigencia": firma.get("VIGENCIA", ""),
                    "referencia": firma.get("Referencia", ""),
                    "fecha_acreditacion": firma.get("Fecha de acreditación", "")
                }
        
        print(f"✅ Firmas cargadas: {len(firmas_map)} inspectores")
        return firmas_map
    except Exception as e:
        print(f"❌ Error cargando firmas: {e}")
        return {}

# La información de normas acreditadas ahora está en Firmas.json

def validar_acreditacion_inspector(codigo_firma, norma_requerida, firmas_map):
    """
    Valida que el inspector esté acreditado para la NOM requerida.
    Retorna (nombre_inspector, ruta_imagen, acreditado) 
    - Si está acreditado: (nombre, imagen, True)
    - Si NO está acreditado: (nombre, imagen, False)
    - Si no existe: (None, None, False)
    """
    if codigo_firma not in firmas_map:
        print(f"   ⚠️ Código de firma '{codigo_firma}' no encontrado en Firmas.json")
        return None, None, False
    
    inspector = firmas_map[codigo_firma]
    nombre = inspector.get("nombre")
    imagen = inspector.get("imagen")
    normas_acreditadas = inspector.get("normas_acreditadas", [])

    # Normalizar valores para comparación
    try:
        req = (norma_requerida or "").strip().upper()
    except Exception:
        req = str(norma_requerida or "").strip().upper()

    normas_norm = [str(n).strip().upper() for n in normas_acreditadas if n]

    # Caso simple: coincidencia exacta
    if req and req in normas_norm:
        print(f"   ✅ Firma validada: {nombre} - {norma_requerida}")
        return nombre, imagen, True

    # Intentar coincidencia por partes (por número de norma o substring)
    try:
        import re
        nums = re.findall(r"\d+", req) if req else []
        for na in normas_norm:
            # coincidencia exacta
            if req and req == na:
                print(f"   ✅ Firma validada: {nombre} - {norma_requerida} == {na}")
                return nombre, imagen, True

            # coincidencia por substring (solo si la cadena de búsqueda es significativa)
            if req and len(req) > 3 and req in na:
                print(f"   ✅ Firma validada por substring: {nombre} - {norma_requerida} ~ {na}")
                return nombre, imagen, True

            # coincidencia por número: comparar tokens numéricos completos para evitar
            # que '4' coincida dentro de '141'. Normalizamos eliminando ceros a la izquierda.
            try:
                nums_na = re.findall(r"\d+", na) if na else []
                for n in nums:
                    if not n:
                        continue
                    n_norm = n.lstrip('0') or '0'
                    for m in nums_na:
                        m_norm = m.lstrip('0') or '0'
                        if n_norm == m_norm:
                            print(f"   ✅ Firma validada por número: {nombre} - {norma_requerida} ~ {na}")
                            return nombre, imagen, True
            except Exception:
                pass
    except Exception:
        pass

    # No acreditada
    print(f"   ⚠️ {nombre} NO está acreditado para {norma_requerida}")
    print(f"   📋 Normas acreditadas: {', '.join(normas_acreditadas)}")
    return nombre, imagen, False

# ---------------------------------------------------------
# PROCESAMIENTO DE FAMILIAS
# ---------------------------------------------------------
def procesar_familias(df):
    if df.empty:
        print("❌ DataFrame vacío")
        return {}

    familias = defaultdict(list)

    for _, row in df.iterrows():
        norma_uva = str(row.get("NORMA UVA", "")).strip()
        folio = str(row.get("FOLIO", "")).strip()
        solicitud = str(row.get("SOLICITUD", "")).strip()
        lista = str(row.get("LISTA", "")).strip()

        key = f"{norma_uva}_{folio}_{solicitud}_{lista}"
        familias[key].append(row.to_dict())

    print(f"✅ Familias procesadas: {len(familias)}")
    return dict(familias)

# ---------------------------------------------------------
# TABLA DE PRODUCTOS Y SUMA
# ---------------------------------------------------------
def preparar_datos_tabla(registros):
    filas_tabla = []
    total_cantidad = 0

    marca_global = ""
    for r in registros:
        if r.get("MARCA"):
            marca_global = str(r["MARCA"]).strip()
            break

    for r in registros:
        marca = str(r.get("MARCA", "") or marca_global).strip()

        codigos_raw = r.get("CODIGO", "")
        codigos = [c.strip() for c in str(codigos_raw).split(",")] if codigos_raw else [""]

        factura = str(r.get("FACTURA", "")).strip()

        try:
            cantidad = int(float(r.get("CANTIDAD", 0)))
        except:
            cantidad = 0

        for codigo in codigos:
            filas_tabla.append({
                "marca": marca,
                "codigo": codigo,
                "factura": factura,
                "cantidad": cantidad
            })

        total_cantidad += cantidad

    return filas_tabla, total_cantidad

# ---------------------------------------------------------
# PREPARAR DATOS POR FAMILIA
# ---------------------------------------------------------
def preparar_datos_familia(
    registros,
    normas_map,
    normas_info_completa,
    clientes_map,
    firmas_map,
    cliente_manual=None,
    rfc_manual=None
):
    """
    Prepara datos completos para el dictamen incluyendo validación de firmas.
    SIEMPRE genera el dictamen, con o sin firma válida.
    """

    r0 = registros[0]

    # FOLIO, SOLICITUD, LISTA
    folio = str(r0.get("FOLIO", "")).strip()
    solicitud_raw = str(r0.get("SOLICITUD", "")).strip()
    # Extraer año desde la solicitud si viene en formato 'NNNNNN/YY'
    # y obtener los dos dígitos del sufijo. Pero para la cadena final
    # usaremos el año actual para la parte UDC y el sufijo extraído
    # para el prefijo de "Solicitud de Servicio".
    solicitud_year_two = ''
    try:
        if '/' in solicitud_raw:
            parts = solicitud_raw.split('/')
            suf = parts[-1].strip()
            if suf.isdigit():
                solicitud_year_two = suf[-2:]
    except Exception:
        solicitud_year_two = ''

    # Si no se obtuvo el sufijo desde la solicitud, usar el año actual
    if not solicitud_year_two:
        solicitud_year_two = datetime.now().strftime("%y")

    # Solicitud: tomar la parte antes de '/' y formatearla a 6 dígitos si es numérica
    solicitud = solicitud_raw.split('/')[0].strip()
    if solicitud.isdigit():
        solicitud = solicitud.zfill(6)
    lista = str(r0.get("LISTA", "")).strip()

    # NORMA
    clasif = str(r0.get("CLASIF UVA", "")).strip()
    try:
        norma_num = str(int(float(clasif)))
    except (ValueError, TypeError):
        norma_num = "".join([c for c in clasif if c.isdigit()])

    norma = ""
    normades = ""
    capitulo = ""

    if norma_num in normas_map:
        norma = normas_map[norma_num]
        normades = normas_info_completa.get(norma, {}).get("nombre", "")
        capitulo = normas_info_completa.get(norma, {}).get("capitulo", "")
    else:
        print(f"⚠️ No se encontró la NOM para CLASIF UVA = {clasif}")

    # Formatear folio a 6 dígitos para la cadena de identificación
    folio_disp = folio.zfill(6) if folio and str(folio).isdigit() else folio
    current_year_two = datetime.now().strftime("%y")
    # `year` se expone en el dict devuelto; definirlo como el año usado
    # en la parte UDC (año actual, dos dígitos) para mantener compatibilidad.
    year = current_year_two
    cadena_identificacion = f"{current_year_two}049UDC{norma}{folio_disp} Solicitud de Servicio: {solicitud_year_two}049USD{norma}{solicitud}-{lista}"

    def fecha_corta(f):
        try:
            return datetime.strptime(str(f), "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            return str(f or "")

    fverificacion = fecha_corta(r0.get("FECHA DE VERIFICACION"))
    femision = fecha_corta(r0.get("FECHA DE ENTRADA"))
    fverificacionlarga = formatear_fecha_larga(r0.get("FECHA DE VERIFICACION"))

    marca = next((str(r.get("MARCA", "")).strip() for r in registros if r.get("MARCA")), "")
    descripcion = next((str(r.get("DESCRIPCION", "")).strip() for r in registros if r.get("DESCRIPCION")), "")

    marca_key = marca.upper()

    # Cliente y RFC
    if cliente_manual:
        cliente = cliente_manual
    else:
        cliente = clientes_map.get(marca_key, {}).get("nombre", marca)

    if rfc_manual:
        rfc = rfc_manual
    else:
        rfc = clientes_map.get(marca_key, {}).get("rfc", "")

    # Tabla de productos
    filas_tabla, total_cantidad = preparar_datos_tabla(registros)

    # Observaciones
    obs_raw = next((str(r.get("OBSERVACIONES DICTAMEN", "")).strip()
                    for r in registros if r.get("OBSERVACIONES DICTAMEN")), "")
    obs = "" if obs_raw.upper() == "NINGUNA" else obs_raw

    print("   🔍 Iniciando generación de etiquetas...")
    generador_etiquetas = GeneradorEtiquetasDecathlon()

    # Generar etiquetas usando clave compuesta para evitar reutilizar una etiqueta
    # de otro registro que comparte el mismo código pero pertenece a distinta
    # solicitud/marca/pais. Construimos una lista de entradas (dict) con los
    # campos mínimos: codigo, solicitud, marca, pais_origen.
    codigos_compuestos = []
    for r in registros:
        codigo = r.get("CODIGO") or r.get('EAN') or r.get('Codigo')
        if not codigo or str(codigo).strip() in ("", "None", "nan"):
            continue
        solicitud = r.get('SOLICITUD') or r.get('Solicitud') or r.get('solicitud') or ''
        marca = r.get('MARCA') or r.get('Marca') or r.get('marca') or ''
        pais = (
            r.get('PAIS ORIGEN') or r.get('PAIS_DE_ORIGEN') or r.get('PAIS DE ORIGEN') or
            r.get('PAIS') or r.get('PAIS ORIG') or ''
        )
        codigos_compuestos.append({
            'codigo': str(codigo).strip(),
            'solicitud': str(solicitud).strip(),
            'marca': str(marca).strip(),
            'pais_origen': str(pais).strip()
        })

    print(f"   📋 Claves compuestas encontradas: {len(codigos_compuestos)} entradas")

    etiquetas_generadas = []
    if codigos_compuestos:
        try:
            print(f"   🏷️ Generando etiquetas para {len(codigos_compuestos)} entradas (clave compuesta)...")
            etiquetas_generadas = generador_etiquetas.generar_etiquetas_por_codigos(codigos_compuestos)
            print(f"   ✅ Etiquetas generadas: {len(etiquetas_generadas)}")
        except Exception as e:
            print(f"   ⚠️ Error generando etiquetas: {e}")
            traceback.print_exc()
            etiquetas_generadas = []
    else:
        print("   ⚠️ No se encontraron códigos válidos en los registros")

    # Buscar entre todos los registros de la familia un inspector (código FIRMA)
    # que esté acreditado para la norma actual. Esto evita usar siempre el
    # primer registro y que se requiera que todos los inspectores tengan las
    # mismas acreditaciones.
    codigo_firma1 = ""
    nombre_firma1 = ""
    imagen_firma1 = ""
    firma1_acreditada = False
    razon_sin_firma = ""

    # Priorizar códigos explícitos bajo claves comunes en cada registro
    posibles_claves = ('FIRMA', 'firma', 'INSPECTOR', 'Inspector', 'inspector')
    for rec in registros:
        for k in posibles_claves:
            try:
                val = rec.get(k)
            except Exception:
                val = None
            if not val:
                continue
            cand = str(val).strip()
            if not cand:
                continue
            # Intentar validar el candidato como código de firma
            nom_tmp, img_tmp, ac_tmp = validar_acreditacion_inspector(cand, norma, firmas_map)
            if nom_tmp and ac_tmp:
                codigo_firma1 = cand
                nombre_firma1 = nom_tmp
                imagen_firma1 = img_tmp
                firma1_acreditada = True
                print(f"   ✅ Firma asignada (desde registros): {nombre_firma1} [{codigo_firma1}]")
                break
        if firma1_acreditada:
            break

    # Si no encontramos ningún inspector acreditado entre los registros,
    # intentar validar el código del primer registro como fallback (comportamiento previo)
    if not firma1_acreditada:
        codigo_fallback = str(r0.get("FIRMA", "")).strip()
        if codigo_fallback:
            print(f"   🔍 No se encontró firma acreditada en la familia; validando fallback: {codigo_fallback} para norma {norma}")
            nombre_firma1, imagen_firma1, firma1_acreditada = validar_acreditacion_inspector(
                codigo_fallback,
                norma,
                firmas_map
            )
            codigo_firma1 = codigo_fallback

        if not nombre_firma1:
            razon_sin_firma = f"Código de firma '{codigo_firma1}' no encontrado en Firmas.json"
            print(f"   ⚠️ DICTAMEN SIN FIRMA: {razon_sin_firma}")
            nombre_firma1 = ""
            imagen_firma1 = ""
            firma1_acreditada = False
        elif not firma1_acreditada:
            razon_sin_firma = f"Inspector {nombre_firma1} no acreditado para {norma}"
            print(f"   ⚠️ DICTAMEN SIN FIRMA: {razon_sin_firma}")
            nombre_firma1 = ""
            imagen_firma1 = ""
        else:
            # Firma válida desde fallback
            print(f"   ✅ Firma asignada (fallback): {nombre_firma1} [{codigo_firma1}]")

    firma_valida = firma1_acreditada

    nombre_firma2, imagen_firma2, aflores_acreditado = validar_acreditacion_inspector(
        "AFLORES", 
        norma, 
        firmas_map
    )
    
    if not nombre_firma2:
        print("   ⚠️ AFLORES no encontrado en Firmas.json")
        nombre_firma2 = ""
        imagen_firma2 = ""
    else:
        print(f"   ✅ Supervisor asignado: {nombre_firma2}")

    return {
        "cadena_identificacion": cadena_identificacion,
        "norma": norma,
        "normades": normades,
        "capitulo": capitulo,

        "year": year,
        "folio": folio,
        "solicitud": solicitud,
        "lista": lista,

        "fverificacion": fverificacion,
        "fverificacionlarga": fverificacionlarga,
        "femision": femision,

        "cliente": cliente,
        "rfc": rfc,

        "producto": descripcion,
        "pedimento": str(r0.get("PEDIMENTO", "")).strip(),

        "tabla_productos": filas_tabla,
        "total_cantidad": total_cantidad,
        "TCantidad": f"{total_cantidad} unidades",

        "obs": obs,

        "etiquetas_lista": etiquetas_generadas,

        "imagen_firma1": imagen_firma1,
        "imagen_firma2": imagen_firma2,
        "nfirma1": nombre_firma1,
        "nfirma2": nombre_firma2,
        
        "firma_valida": firma_valida,
        "razon_sin_firma": razon_sin_firma,
        "codigo_firma_solicitado": codigo_firma1
    }
    # Incluir datos originales de solicitud para uso por otros módulos
    # `solicitud_raw` conserva el sufijo de año (ej. '002334/25') y
    # `solicitud_year_two` contiene los dos dígitos del sufijo cuando existe.
    result = {
        **{
            "cadena_identificacion": cadena_identificacion,
            "norma": norma,
            "normades": normades,
            "capitulo": capitulo,

            "year": year,
            "folio": folio,
            "solicitud": solicitud,
            "lista": lista,

            "fverificacion": fverificacion,
            "fverificacionlarga": fverificacionlarga,
            "femision": femision,

            "cliente": cliente,
            "rfc": rfc,

            "producto": descripcion,
            "pedimento": str(r0.get("PEDIMENTO", "")).strip(),

            "tabla_productos": filas_tabla,
            "total_cantidad": total_cantidad,
            "TCantidad": f"{total_cantidad} unidades",

            "obs": obs,

            "etiquetas_lista": etiquetas_generadas,

            "imagen_firma1": imagen_firma1,
            "imagen_firma2": imagen_firma2,
            "nfirma1": nombre_firma1,
            "nfirma2": nombre_firma2,
            
            "firma_valida": firma_valida,
            "razon_sin_firma": razon_sin_firma,
            "codigo_firma_solicitado": codigo_firma1
        }
    }
    # Añadir metadatos de solicitud sin perder compatibilidad con consumidores existentes
    result["solicitud_raw"] = solicitud_raw
    result["solicitud_year_two"] = solicitud_year_two

    return result


