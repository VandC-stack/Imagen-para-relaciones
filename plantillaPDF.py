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

def cargar_firmas(ruta="data/Firmas.json"):
    """Carga el mapeo de firmas de inspectores indexado por código FIRMA."""
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            firmas = json.load(f)
        
        firmas_map = {}
        for firma in firmas:
            codigo = firma.get("FIRMA", "").strip()
            if codigo:
                firmas_map[codigo] = {
                    "nombre": firma.get("NOMBRE DE INSPECTOR", "").strip(),
                    "imagen": firma.get("IMAGEN", "").strip()
                }
        
        print(f"✅ Firmas cargadas: {len(firmas_map)} inspectores")
        return firmas_map
    except Exception as e:
        print(f"❌ Error cargando firmas: {e}")
        return {}

def cargar_inspectores_acreditados(ruta="data/Inspectores.json"):
    """Carga inspectores con sus normas acreditadas."""
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            inspectores = json.load(f)
        
        inspectores_normas = {}
        for inspector in inspectores:
            colaborador = inspector.get("Colaborador", "").strip()
            normas_str = inspector.get("Normas acreditadas", "").strip()
            
            if colaborador and normas_str:
                normas_list = [n.strip() for n in normas_str.split(",")]
                inspectores_normas[colaborador] = normas_list
        
        print(f"✅ Inspectores acreditados cargados: {len(inspectores_normas)}")
        return inspectores_normas
    except Exception as e:
        print(f"⚠️ No se pudo cargar Inspectores.json: {e}")
        return {}

def obtener_firma_validada(codigo_firma, norma_requerida, firmas_map, inspectores_normas):
    """
    Valida que el inspector esté acreditado para la NOM requerida.
    Retorna (nombre_inspector, ruta_imagen) o (None, None) si no está acreditado.
    """
    if codigo_firma not in firmas_map:
        print(f"   ⚠️ Código de firma no encontrado: {codigo_firma}")
        return None, None
    
    firma_data = firmas_map[codigo_firma]
    nombre_inspector = firma_data.get("nombre")
    ruta_imagen = firma_data.get("imagen")
    
    # Validar acreditación
    if nombre_inspector in inspectores_normas:
        normas_acreditadas = inspectores_normas[nombre_inspector]
        if norma_requerida in normas_acreditadas:
            print(f"   ✅ Firma validada: {nombre_inspector} - {norma_requerida}")
            return nombre_inspector, ruta_imagen
        else:
            print(f"   ⚠️ {nombre_inspector} NO está acreditado para {norma_requerida}")
            return None, None
    else:
        print(f"   ⚠️ Inspector {nombre_inspector} no encontrado en acreditaciones")
        return None, None

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
    inspectores_normas,
    cliente_manual,
    rfc_manual
):
    """
    PREPARA TODOS LOS DATOS REALES PARA GENERAR EL DICTAMEN
    Esta versión es compatible con generador_dictamen.py (8 parámetros).

    - Busca inspector en Firmas.json
    - Genera datos del dictamen
    - Prepara etiquetas
    - Prepara firmas
    - Construye y devuelve el diccionario final “datos”
    """

    try:
        # ---------------------------
        # 1) REGISTRO PRINCIPAL
        # ---------------------------
        reg = registros[0]  # primer registro de la lista/familia

        cliente = cliente_manual if cliente_manual else reg.get("CLIENTE", "")
        rfc = rfc_manual if rfc_manual else reg.get("RFC", "")
        inspector_nombre = reg.get("INSPECTOR", "").strip()

        # ---------------------------------
        # 2) BUSCAR INSPECTOR EN FIRMAS.JSON
        # ---------------------------------
        inspector = next(
            (i for i in firmas_map if i["NOMBRE DE INSPECTOR"].strip().lower() == inspector_nombre.lower()),
            None
        )

        if inspector is None:
            print(f"⚠️ Inspector {inspector_nombre} no encontrado en Firmas.json")
            return None

        # ------------------------
        # 3) DATOS DEL INSPECTOR
        # ------------------------
        puesto = inspector.get("Puesto", "")
        vigencia = inspector.get("VIGENCIA", "")
        normas = inspector.get("Normas acreditadas", [])
        referencia = inspector.get("Referencia", "")
        fecha_acredit = inspector.get("Fecha de acreditación", "")
        normas_str = ", ".join(normas) if isinstance(normas, list) else str(normas)

        # ------------------------
        # 4) DATOS DE ETIQUETAS
        # ------------------------
        # El generador espera lista de imágenes: etiquetas_lista
        etiquetas = reg.get("ETIQUETAS_GENERADAS", [])
        etiquetas_lista = []

        for e in etiquetas:
            etiquetas_lista.append({
                "imagen_bytes": e.get("imagen_bytes"),
                "tamaño_cm": e.get("tamaño_cm", (5, 5))
            })

        # ------------------------
        # 5) FIRMAS
        # ------------------------
        # Firma del inspector (firma1)
        ruta_firma1 = inspector.get("IMAGEN", "")
        if ruta_firma1.startswith("Firmas/"):
            ruta_firma1 = os.path.join("Firmas", ruta_firma1.replace("Firmas/", ""))

        # Supervisor fijo (el que tenga FIRMA = "AFLORES")
        supervisor = next((i for i in firmas_map if i.get("FIRMA") == "AFLORES"), None)

        ruta_firma2 = ""
        nombre_firma2 = ""
        if supervisor:
            ruta_firma2 = supervisor.get("IMAGEN", "")
            if ruta_firma2.startswith("Firmas/"):
                ruta_firma2 = os.path.join("Firmas", ruta_firma2.replace("Firmas/", ""))
            nombre_firma2 = supervisor["NOMBRE DE INSPECTOR"]

        # ------------------------
        # 6) TABLA DE PRODUCTOS
        # ------------------------
        tabla_productos = []
        total_cantidad = 0

        for r in registros:
            tabla_productos.append({
                "marca": r.get("MARCA", ""),
                "codigo": r.get("CODIGO", ""),
                "factura": r.get("FACTURA", ""),
                "cantidad": r.get("CANTIDAD", 0)
            })
            total_cantidad += r.get("CANTIDAD", 0)

        # ------------------------
        # 7) CONSTRUIR EL OBJETO FINAL
        # ------------------------
        datos = {
            # datos generales
            "cliente": cliente,
            "rfc": rfc,
            "producto": reg.get("PRODUCTO", ""),
            "norma": reg.get("NORMA", ""),
            "normades": normas_info_completa.get(reg.get("NORMA", ""), ""),
            "cadena_identificacion": reg.get("LISTA", ""),

            # fechas
            "fverificacion": reg.get("F_VERIFICACION", ""),
            "fverificacionlarga": reg.get("F_VERIFICACION_LARGA", ""),
            "femision": reg.get("F_EMISION", ""),

            # capítulo
            "capitulo": reg.get("CAPITULO", ""),

            # observaciones
            "obs": reg.get("OBSERVACIONES", ""),

            # tabla productos y lote
            "tabla_productos": tabla_productos,
            "TCantidad": f"{total_cantidad} unidades",

            # etiquetas
            "etiquetas_lista": etiquetas_lista,

            # firmas
            "imagen_firma1": ruta_firma1,
            "imagen_firma2": ruta_firma2,
            "nfirma1": inspector.get("NOMBRE DE INSPECTOR", ""),
            "nfirma2": nombre_firma2,

            # info inspector extra
            "puesto": puesto,
            "vigencia": vigencia,
            "normas_inspector": normas_str,
            "referencia_inspector": referencia,
            "acreditacion_inspector": fecha_acredit
        }

        return datos

    except Exception as e:
        print("❌ Error en preparar_datos_familia:", e)
        traceback.print_exc()
        return None

