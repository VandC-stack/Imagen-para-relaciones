"""plantilla.py - Funciones de carga y preparación de datos"""
import pandas as pd
import json
from datetime import datetime
from collections import defaultdict
import os

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
    cliente_manual=None,
    rfc_manual=None
):

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
        normades = normas_info_completa[norma]["nombre"]
        capitulo = normas_info_completa[norma]["capitulo"]
    else:
        print(f"⚠️ No se encontró la NOM para CLASIF UVA = {clasif}")

    # CADENA IDENTIFICACIÓN
    cadena_identificacion = (
        f"{year}049UDC{norma}{folio} "
        f"Solicitud de Servicio: {year}049USD{norma}{solicitud}-{lista}"
    )

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

    # Cliente
    if cliente_manual:
        cliente = cliente_manual
    else:
        cliente = marca
        if clientes_map and marca_key in clientes_map:
            cliente = clientes_map[marca_key]["nombre"]

    # RFC
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

    # Obtener códigos EAN de los registros
    codigos = []
    for r in registros:
        codigo = r.get("CODIGO")
        if codigo and str(codigo).strip() not in ("", "None", "nan"):
            codigos.append(str(codigo).strip())
    
    print(f"   📋 Códigos encontrados: {codigos}")

    # Generar etiquetas
    etiquetas_generadas = []
    if codigos:
        print(f"   🏷️ Generando etiquetas para {len(codigos)} códigos...")
        etiquetas_generadas = generador_etiquetas.generar_etiquetas_por_codigos(codigos)
        print(f"   ✅ Etiquetas generadas: {len(etiquetas_generadas)}")
    else:
        print("   ⚠️ No se encontraron códigos válidos en los registros")

    nfirma1 = str(r0.get("FIRMA", "")).strip()

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

        "firma1": "________________________",
        "firma2": "________________________",
        "nfirma1": nfirma1,
        "nfirma2": "Responsable de Supervisión UI",
    }
