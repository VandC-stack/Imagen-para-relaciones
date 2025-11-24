"""plantilla.py - Funciones de carga y preparación de datos"""
import pandas as pd
import json
from datetime import datetime
from collections import defaultdict
import os
import traceback
from etiqueta_dictamen import GeneradorEtiquetasDecathlon

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
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        df = pd.DataFrame(data)
        print(f"✅ Tabla de relación cargada: {len(df)} registros")
        return df
    except Exception as e:
        print(f"❌ Error cargando tabla de relación: {e}")
        return pd.DataFrame()

def cargar_normas(ruta="data/Normas.json"):
    """Carga las NOMs y las indexa usando el número inicial (ej: 24, 50, 141)."""
    try:
        with open(ruta, "r", encoding="utf-8") as f:
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
        with open(ruta, "r", encoding="utf-8") as f:
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
        with open(ruta, "r", encoding="utf-8") as f:
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
    
    # Validar acreditación
    if norma_requerida in normas_acreditadas:
        print(f"   ✅ Firma validada: {nombre} - {norma_requerida}")
        return nombre, imagen, True
    else:
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

    # YEAR
    year = datetime.now().strftime("%y")

    # FOLIO, SOLICITUD, LISTA
    folio = str(r0.get("FOLIO", "")).strip()
    solicitud_raw = str(r0.get("SOLICITUD", "")).strip()
    solicitud = solicitud_raw.split("/")[0]
    lista = str(r0.get("LISTA", "")).strip()

    # NORMA
    clasif = str(r0.get("CLASIF UVA", "")).strip()
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

    cadena_identificacion = f"{year}049UDC{norma}{folio} Solicitud de Servicio: {year}049USD{norma}{solicitud}-{lista}"

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

    codigos = []
    for r in registros:
        codigo = r.get("CODIGO")
        if codigo and str(codigo).strip() not in ("", "None", "nan"):
            codigos.append(str(codigo).strip())
    
    print(f"   📋 Códigos encontrados: {codigos}")

    etiquetas_generadas = []
    if codigos:
        try:
            print(f"   🏷️ Generando etiquetas para {len(codigos)} códigos...")
            etiquetas_generadas = generador_etiquetas.generar_etiquetas_por_codigos(codigos)
            print(f"   ✅ Etiquetas generadas: {len(etiquetas_generadas)}")
        except Exception as e:
            print(f"   ⚠️ Error generando etiquetas: {e}")
            traceback.print_exc()
            etiquetas_generadas = []
    else:
        print("   ⚠️ No se encontraron códigos válidos en los registros")

    codigo_firma1 = str(r0.get("FIRMA", "")).strip()
    
    print(f"   🔍 Validando firma: {codigo_firma1} para norma {norma}")
    
    nombre_firma1, imagen_firma1, firma1_acreditada = validar_acreditacion_inspector(
        codigo_firma1, 
        norma, 
        firmas_map
    )
    
    firma_valida = False
    razon_sin_firma = ""
    
    if not nombre_firma1:
        # Código no encontrado
        razon_sin_firma = f"Código de firma '{codigo_firma1}' no encontrado en Firmas.json"
        print(f"   ⚠️ DICTAMEN SIN FIRMA: {razon_sin_firma}")
        nombre_firma1 = ""
        imagen_firma1 = ""
    elif not firma1_acreditada:
        # Inspector no acreditado para esta norma
        razon_sin_firma = f"Inspector {nombre_firma1} no acreditado para {norma}"
        print(f"   ⚠️ DICTAMEN SIN FIRMA: {razon_sin_firma}")
        nombre_firma1 = ""
        imagen_firma1 = ""
    else:
        # Firma válida
        firma_valida = True
        print(f"   ✅ Firma asignada: {nombre_firma1}")
    
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
