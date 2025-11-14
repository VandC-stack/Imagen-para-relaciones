"""
ArmadoDictamen.py - Funciones de carga y preparación de datos para dictámenes
"""

import os
import json
import pandas as pd
from datetime import datetime
from collections import defaultdict
import re

def formatear_fecha_larga(fecha_str):
    """Convierte fecha de formato 04/09/2025 a '4 de septiembre de 2025'"""
    if pd.isna(fecha_str) or fecha_str == '':
        return ""
    
    try:
        # Si ya está en formato día/mes/año
        if isinstance(fecha_str, str) and '/' in fecha_str:
            fecha = datetime.strptime(fecha_str, '%d/%m/%Y')
        else:
            # Si viene en formato YYYY-MM-DD
            fecha = datetime.strptime(str(fecha_str), '%Y-%m-%d')
        
        # Diccionario de meses en español
        meses = {
            1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 
            5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
            9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
        }
        
        dia = fecha.day
        mes = meses[fecha.month]
        anio = fecha.year
        
        return f"{dia} de {mes} de {anio}"
        
    except Exception as e:
        print(f"⚠️  Error formateando fecha {fecha_str}: {e}")
        return str(fecha_str)

# FUNCIONES DE CARGA DE DATOS
def cargar_tabla_relacion():
    """Carga el archivo tabla_de_relacion.json con mejor manejo de errores"""
    print("🔍 Buscando tabla_de_relacion.json...")
    
    # Lista completa de posibles ubicaciones
    posibles_rutas = [
        'data/tabla_de_relacion.json',     
    ]
    
    # Verificar si la carpeta data existe
    if not os.path.exists('data'):
        print("❌ La carpeta 'data' no existe en el directorio actual")
        return None
    
    # Buscar el archivo
    archivo_encontrado = None
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            archivo_encontrado = ruta
            print(f"✅ Archivo encontrado: {ruta}")
            break
    
    if not archivo_encontrado:
        print("❌ No se encontró tabla_de_relacion.json en ninguna ubicación")
        return None
    
    # Intentar cargar el archivo
    try:
        with open(archivo_encontrado, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        
        print(f"✅ Tabla cargada exitosamente: {len(datos)} registros")
        return datos
        
    except Exception as e:
        print(f"❌ Error cargando archivo: {e}")
        return None

def cargar_normas():
    """Carga el archivo Normas.json y crea un mapeo de números a códigos de norma"""
    print("🔍 Buscando Normas.json...")
    
    posibles_rutas = [
        'data/Normas.json',
        'Normas.json',
        '../data/Normas.json'
    ]
    
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            try:
                with open(ruta, 'r', encoding='utf-8') as f:
                    normas_data = json.load(f)
                
                print(f"✅ Archivo Normas.json encontrado: {ruta}")
                
                # Crear mapeo de números a códigos de norma y almacenar información completa
                normas_map = {}
                normas_info_completa = {}
                
                if isinstance(normas_data, list):
                    print("📝 Procesando lista de normas...")
                    
                    for norma_item in normas_data:
                        if isinstance(norma_item, dict):
                            codigo_norma = norma_item.get('NOM', '')
                            nombre_norma = norma_item.get('NOMBRE', '')
                            capitulo_norma = norma_item.get('CAPITULO', '')
                            
                            # Guardar información completa de la norma
                            normas_info_completa[codigo_norma] = {
                                'nombre': nombre_norma,
                                'capitulo': capitulo_norma
                            }
                            
                            # Extraer números del código de norma para crear mapeos
                            numeros = re.findall(r'\d+', codigo_norma)
                            for num in numeros:
                                # Agregar el número tal cual (ej: "004")
                                normas_map[num] = codigo_norma
                                # También agregar sin ceros a la izquierda (ej: "4")
                                if num.startswith('0'):
                                    normas_map[num.lstrip('0')] = codigo_norma
                            
                            # También mapear el código completo a sí mismo
                            normas_map[codigo_norma] = codigo_norma
                    
                    print(f"✅ Mapeo de normas creado: {len(normas_map)} entradas")
                    print(f"✅ Información completa cargada para {len(normas_info_completa)} normas")
                    
                    return normas_map, normas_info_completa
                    
                else:
                    print("⚠️  Formato de normas no reconocido, usando valores por defecto")
                    normas_map = {
                        "24": "NOM-004-SE-2021",
                        "4": "NOM-004-SE-2021", 
                        "1": "NOM-001-SE-2021"
                    }
                    normas_info_completa = {
                        "NOM-004-SE-2021": {
                            "nombre": "Información Comercial- etiquetado de productos textiles, prendas de vestir, sus accesorios y ropa de casa",
                            "capitulo": "4 (Especificaciones de información comercial) y 5 (Instrumentación de la información comercial)"
                        }
                    }
                    return normas_map, normas_info_completa
                
            except Exception as e:
                print(f"❌ Error cargando {ruta}: {e}")
    
    print("⚠️  No se encontró Normas.json, usando valores por defecto")
    normas_map = {
        "24": "NOM-004-SE-2021",
        "4": "NOM-004-SE-2021", 
        "1": "NOM-001-SE-2021",
        "15": "NOM-015-SCFI-2007",
        "20": "NOM-020-SCFI-1997",
        "24": "NOM-024-SCFI-2013",
        "50": "NOM-050-SCFI-2004",
        "51": "NOM-051-SCFI/SSA1-2010",
        "141": "NOM-141-SSA1/SCFI-2012",
        "142": "NOM-142-SSA1/SCFI-2014",
        "189": "NOM-189-SSA1/SCFI-2018",
        "235": "NOM-235-SE-2020"
    }
    normas_info_completa = {
        "NOM-004-SE-2021": {
            "nombre": "Información Comercial- etiquetado de productos textiles, prendas de vestir, sus accesorios y ropa de casa",
            "capitulo": "4 (Especificaciones de información comercial) y 5 (Instrumentación de la información comercial)"
        },
        "NOM-024-SCFI-2013": {
            "nombre": "Etiquetado de productos y prendas de vestir, calzado y otros",
            "capitulo": "5 (Información Comercial)"
        }
    }
    return normas_map, normas_info_completa

def procesar_familias(tabla_datos):
    """Agrupa registros por LISTA"""
    if not tabla_datos:
        print("❌ No hay datos para procesar")
        return {}
    
    familias = defaultdict(list)
    for registro in tabla_datos:
        lista = registro.get('LISTA', '')
        familias[lista].append(registro)
    
    print(f"✅ {len(familias)} familias encontradas")
    return familias

def preparar_datos_familia(registros, normas_map, normas_info_completa, cliente_manual=None, rfc_manual=None):
    """Prepara datos para una familia específica - VERSIÓN CORREGIDA"""
    if not registros:
        return None
        
    primer_registro = registros[0]
    
    # Información básica
    year = datetime.now().strftime("%y")
    norma_uva = primer_registro.get('NORMA UVA', '')
    folio = str(primer_registro.get('FOLIO', ''))
    solicitud = str(primer_registro.get('SOLICITUD', ''))
    lista = str(primer_registro.get('LISTA', ''))
    
    # Mapear norma
    norma = "NOM-001"  # Valor por defecto
    capitulo = "4"  # Valor por defecto
    normades = "ESPECIFICACIONES DE SEGURIDAD"  # Valor por defecto
    
    if not pd.isna(norma_uva) and norma_uva != '':
        norma_str = str(int(norma_uva)) if isinstance(norma_uva, (int, float)) else str(norma_uva)
        
        # Buscar en el mapa de normas
        if norma_str in normas_map:
            norma_codigo = normas_map[norma_str]
            norma = norma_codigo
            
            # Buscar información completa de la norma
            if norma_codigo in normas_info_completa:
                normades = normas_info_completa[norma_codigo].get('nombre', 'ESPECIFICACIONES DE SEGURIDAD')
                capitulo = normas_info_completa[norma_codigo].get('capitulo', '4')
                print(f"   📋 Norma UVA {norma_str} → {norma}")
                print(f"   📖 Capítulo: {capitulo}")
                print(f"   📄 Descripción: {normades}")
            else:
                print(f"   ⚠️  No se encontró información completa para la norma {norma_codigo}")
        else:
            # Si no se encuentra, buscar coincidencias parciales
            norma_encontrada = None
            for norma_key, norma_value in normas_map.items():
                # Buscar por coincidencia exacta en claves numéricas
                if norma_key.isdigit() and norma_str == norma_key:
                    norma_encontrada = norma_value
                    break
            
            if norma_encontrada:
                norma = norma_encontrada
                # Buscar información completa de la norma encontrada
                if norma in normas_info_completa:
                    normades = normas_info_completa[norma].get('nombre', 'ESPECIFICACIONES DE SEGURIDAD')
                    capitulo = normas_info_completa[norma].get('capitulo', '4')
                print(f"   📋 Norma UVA {norma_str} → {norma} (por coincidencia exacta)")
            else:
                # Si aún no se encuentra, usar el formato NOM-XXX
                norma = f"NOM-{norma_str:03d}"
                print(f"   ⚠️  Norma UVA {norma_str} no encontrada en el mapeo, usando {norma}")
    
    # Fechas
    def formatear_fecha(fecha_str):
        if pd.isna(fecha_str) or fecha_str == '':
            return ""
        try:
            fecha = datetime.strptime(str(fecha_str), '%Y-%m-%d')
            return fecha.strftime('%d/%m/%Y')
        except:
            return str(fecha_str)
    
    fverificacion = formatear_fecha(primer_registro.get('FECHA DE VERIFICACION', ''))
    femision = formatear_fecha(primer_registro.get('FECHA DE ENTRADA', ''))
    
    # Fecha larga para el texto del dictamen
    fverificacionlarga = formatear_fecha_larga(primer_registro.get('FECHA DE VERIFICACION', ''))
    
    # INICIALIZAR VARIABLES CRÍTICAS PRIMERO
    marca = ""
    modelo = ""
    descripcion = ""
    
    # Buscar la primera marca, modelo y descripción no vacíos en todos los registros
    for registro in registros:
        if not pd.isna(registro.get('MARCA', '')) and registro.get('MARCA', '') != '':
            marca = registro.get('MARCA', '')
            break
    
    for registro in registros:
        if not pd.isna(registro.get('MODELO', '')) and registro.get('MODELO', '') != '':
            modelo = registro.get('MODELO', '')
            break
    
    for registro in registros:
        if not pd.isna(registro.get('DESCRIPCION', '')) and registro.get('DESCRIPCION', '') != '':
            descripcion = registro.get('DESCRIPCION', '')
            break
    
    # Cliente y RFC - USAR LOS VALORES MANUALES SI SE PROVEEN
    if cliente_manual and rfc_manual:
        cliente = cliente_manual
        rfc = rfc_manual
        print(f"   👤 Cliente manual: {cliente}")
    else:
        # Búsqueda automática (comportamiento original)
        cliente, rfc = marca, ""
        if not pd.isna(marca) and marca != '':
            marca_upper = marca.upper()
            # Cargar clientes para búsqueda automática
            clientes_list = cargar_clientes()
            for cliente_info in clientes_list:
                cliente_marca = cliente_info.get('MARCA', '').upper()
                if marca_upper == cliente_marca or marca_upper in cliente_marca:
                    cliente = cliente_info.get('CLIENTE', marca)
                    rfc = cliente_info.get('RFC', '')
                    break
        print(f"   👤 Cliente automático: {marca} → {cliente}")
    
    # Producto - usar la descripción encontrada
    producto = descripcion if descripcion else "Producto no especificado"
    if pd.isna(producto) or producto == '':
        producto = "Producto no especificado"
    
    # Códigos y facturas
    codigos = []
    facturas = []
    for registro in registros:
        codigo = registro.get('CODIGO', '')
        factura = registro.get('FACTURA', '')
        
        if not pd.isna(codigo) and codigo != '':
            if ',' in str(codigo):
                codigos.extend([c.strip() for c in str(codigo).split(',')])
            else:
                codigos.append(str(codigo))
        
        if not pd.isna(factura) and factura != '':
            if ',' in str(factura):
                facturas.extend([f.strip() for f in str(factura).split(',')])
            else:
                facturas.append(str(factura))
    
    rowCodigo = ', '.join(list(dict.fromkeys(codigos))) if codigos else ""
    rowFactura = ', '.join(list(dict.fromkeys(facturas))) if facturas else ""
    
    # Cantidades
    total_cantidad = 0
    for registro in registros:
        cantidad = registro.get('CANTIDAD', 0)
        if not pd.isna(cantidad) and isinstance(cantidad, (int, float)):
            total_cantidad += cantidad
    
    # Observaciones
    obs = ""
    for registro in registros:
        observaciones = registro.get('OBSERVACIONES DICTAMEN', '')
        if not pd.isna(observaciones) and observaciones and observaciones != '':
            obs = str(observaciones)
            break
    
    # Firmas
    firma = primer_registro.get('FIRMA', '')
    nfirma1 = firma if not pd.isna(firma) and firma != '' else "Inspector no asignado"
    
    # DATOS FINALES - AHORA CON VARIABLES SIEMPRE DEFINIDAS
    datos_finales = {
        'year': year,
        'norma': norma,
        'folio': folio,
        'solicitud': solicitud,
        'lista': lista,
        'fverificacion': fverificacion,
        'femision': femision,
        'fverificacionlarga': fverificacionlarga,  # Ahora con formato largo
        'cliente': cliente,
        'rfc': rfc,
        'producto': producto,
        'pedimento': str(primer_registro.get('PEDIMENTO', '')),
        'capitulo': capitulo,
        'normades': normades,
        'rowMarca': marca if not pd.isna(marca) and marca != '' else "",
        'rowModelo': modelo if not pd.isna(modelo) and modelo != '' else "",
        'rowCodigo': rowCodigo,
        'rowFactura': rowFactura,
        'rowCantidad': str(total_cantidad),
        'TCantidad': f"{total_cantidad} unidades",
        'obs': obs,
        'etiqueta1': '', 'etiqueta2': '', 'etiqueta3': '', 'etiqueta4': '', 'etiqueta5': '',
        'etiqueta6': '', 'etiqueta7': '', 'etiqueta8': '', 'etiqueta9': '', 'etiqueta10': '',
        'img1': '', 'img2': '', 'img3': '', 'img4': '', 'img5': '',
        'img6': '', 'img7': '', 'img8': '', 'img9': '', 'img10': '',
        'firma1': '________________________',
        'firma2': '________________________',
        'nfirma1': nfirma1,
        'nfirma2': 'Responsable de Supervisión UI'
    }
    
    print(f"   ✅ Datos preparados: Cliente={cliente}, Marca={marca}, Producto={producto[:50]}...")
    print(f"   📅 Fecha larga: {fverificacionlarga}")
    print(f"   📖 Capítulo final: {capitulo}")
    print(f"   📄 Descripción final: {normades}")
    return datos_finales

def cargar_clientes():
    """Carga el archivo Clientes.json (para búsqueda automática)"""
    try:
        with open('data/Clientes.json', 'r', encoding='utf-8') as f:
            clientes = json.load(f)
        return clientes
    except:
        return []