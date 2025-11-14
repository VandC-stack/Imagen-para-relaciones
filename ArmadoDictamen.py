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
    
    posibles_rutas = [
        'data/tabla_de_relacion.json',     
    ]
    
    if not os.path.exists('data'):
        print("❌ La carpeta 'data' no existe en el directorio actual")
        return None
    
    archivo_encontrado = None
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            archivo_encontrado = ruta
            print(f"✅ Archivo encontrado: {ruta}")
            break
    
    if not archivo_encontrado:
        print("❌ No se encontró tabla_de_relacion.json en ninguna ubicación")
        return None
    
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
                
                normas_map = {}
                normas_info_completa = {}
                
                if isinstance(normas_data, list):
                    print("📝 Procesando lista de normas...")
                    
                    for norma_item in normas_data:
                        if isinstance(norma_item, dict):
                            codigo_norma = norma_item.get('NOM', '')
                            nombre_norma = norma_item.get('NOMBRE', '')
                            capitulo_norma = norma_item.get('CAPITULO', '')
                            
                            normas_info_completa[codigo_norma] = {
                                'nombre': nombre_norma,
                                'capitulo': capitulo_norma
                            }
                            
                            numeros = re.findall(r'\d+', codigo_norma)
                            for num in numeros:
                                normas_map[num] = codigo_norma
                                if num.startswith('0'):
                                    normas_map[num.lstrip('0')] = codigo_norma
                            
                            normas_map[codigo_norma] = codigo_norma
                    
                    print(f"✅ Mapeo de normas creado: {len(normas_map)} entradas")
                    return normas_map, normas_info_completa
                    
                else:
                    print("⚠️  Formato de normas no reconocido")
                    return {}, {}
                
            except Exception as e:
                print(f"❌ Error cargando {ruta}: {e}")
    
    print("⚠️  No se encontró Normas.json")
    return {}, {}

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

def extraer_valor_seguro(registro, campo):
    """Extrae valores de manera segura SIN valores por defecto"""
    valor = registro.get(campo, '')
    if pd.isna(valor) or valor == '':
        return ''
    return str(valor).strip()

def preparar_datos_familia(registros, normas_map, normas_info_completa, cliente_manual=None, rfc_manual=None):
    """Prepara datos para una familia específica - VERSIÓN MEJORADA CON DEBUG"""
    if not registros:
        print("❌ No hay registros para procesar")
        return None
        
    primer_registro = registros[0]
    
    # Información básica con valores por defecto seguros
    year = datetime.now().strftime("%y")  # Formato de 2 dígitos
    
    # Extraer valores con manejo seguro de NaN/vacíos
    norma_uva = extraer_valor_seguro(primer_registro, 'NORMA UVA')
    folio = extraer_valor_seguro(primer_registro, 'FOLIO')
    
    # SOLICITUD: Extraer solo la parte antes de la barra "/"
    solicitud_raw = extraer_valor_seguro(primer_registro, 'SOLICITUD')
    if '/' in solicitud_raw:
        solicitud = solicitud_raw.split('/')[0].strip()
        print(f"   🔄 Solicitud modificada: '{solicitud_raw}' -> '{solicitud}'")
    else:
        solicitud = solicitud_raw
    
    lista = extraer_valor_seguro(primer_registro, 'LISTA')
    
    print(f"   📊 Valores extraídos - Norma UVA: '{norma_uva}', Folio: '{folio}', Solicitud: '{solicitud}', Lista: '{lista}'")
    
    # Mapear norma SIN VALORES POR DEFECTO
    norma = ""
    capitulo = ""
    normades = ""
    
    if norma_uva:
        norma_str = str(int(float(norma_uva))) if norma_uva.replace('.', '').isdigit() else str(norma_uva)
        norma_str = norma_str.strip()
        
        print(f"   🔍 Buscando norma: '{norma_str}' en mapeo de normas")
        
        if norma_str in normas_map:
            norma_codigo = normas_map[norma_str]
            norma = norma_codigo
            
            if norma_codigo in normas_info_completa:
                normades = normas_info_completa[norma_codigo].get('nombre', '')
                capitulo = normas_info_completa[norma_codigo].get('capitulo', '')
                print(f"   ✅ Norma UVA '{norma_str}' → '{norma}'")
            else:
                print(f"   ⚠️  No se encontró información completa para la norma '{norma_codigo}'")
        else:
            print(f"   ❌ Norma UVA '{norma_str}' no encontrada en mapeo")
    
    # VERIFICACIÓN CRÍTICA DE VARIABLES
    print(f"   🔍 VARIABLES PARA CADENA - year: '{year}', norma: '{norma}', folio: '{folio}', solicitud: '{solicitud}', lista: '{lista}'")
    
    # GENERAR PLANTILLA CORRECTA - FORMATO EXACTO
    try:
        cadena_identificacion = f"{year}049UDC{norma}{folio} Solicitud de Servicio: {year}049USD{norma}{solicitud}-{lista}"
        print(f"   ✅ Cadena de identificación GENERADA: {cadena_identificacion}")
    except Exception as e:
        print(f"   ❌ Error generando cadena: {e}")
        cadena_identificacion = ""

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
    fverificacionlarga = formatear_fecha_larga(primer_registro.get('FECHA DE VERIFICACION', ''))
    
    # Buscar datos en registros SIN VALORES POR DEFECTO
    marca = ""
    modelo = ""
    descripcion = ""
    
    for registro in registros:
        marca_temp = registro.get('MARCA', '')
        if not pd.isna(marca_temp) and marca_temp != '':
            marca = marca_temp
            break
    
    for registro in registros:
        modelo_temp = registro.get('MODELO', '')
        if not pd.isna(modelo_temp) and modelo_temp != '':
            modelo = modelo_temp
            break
    
    for registro in registros:
        descripcion_temp = registro.get('DESCRIPCION', '')
        if not pd.isna(descripcion_temp) and descripcion_temp != '':
            descripcion = descripcion_temp
            break
    
    # Cliente y RFC
    cliente = ""
    rfc = ""
    
    if cliente_manual and rfc_manual:
        cliente = cliente_manual
        rfc = rfc_manual
        print(f"   👤 Cliente manual: {cliente}")
    else:
        cliente = marca
        if marca:
            marca_upper = marca.upper()
            clientes_list = cargar_clientes()
            for cliente_info in clientes_list:
                cliente_marca = cliente_info.get('MARCA', '').upper()
                if marca_upper == cliente_marca or marca_upper in cliente_marca:
                    cliente = cliente_info.get('CLIENTE', marca)
                    rfc = cliente_info.get('RFC', '')
                    break
        print(f"   👤 Cliente: {cliente}")
    
    # Producto
    producto = descripcion if descripcion else ""
    
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
    nfirma1 = firma if not pd.isna(firma) and firma != '' else ""
    
    # DATOS FINALES - SIN VALORES POR DEFECTO
    datos_finales = {
        'year': year,
        'norma': norma,
        'folio': folio,
        'solicitud': solicitud,
        'lista': lista,
        'fverificacion': fverificacion,
        'femision': femision,
        'fverificacionlarga': fverificacionlarga,
        'cliente': cliente,
        'rfc': rfc,
        'producto': producto,
        'pedimento': extraer_valor_seguro(primer_registro, 'PEDIMENTO'),
        'capitulo': capitulo,
        'normades': normades,
        'cadena_identificacion': cadena_identificacion,
        'rowMarca': marca,
        'rowModelo': modelo,
        'rowCodigo': rowCodigo,
        'rowFactura': rowFactura,
        'rowCantidad': str(total_cantidad) if total_cantidad > 0 else "",
        'TCantidad': f"{total_cantidad} unidades" if total_cantidad > 0 else "",
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
    
    # VERIFICACIÓN FINAL
    print(f"   🔍 VERIFICACIÓN FINAL:")
    print(f"      year: '{datos_finales['year']}'")
    print(f"      norma: '{datos_finales['norma']}'")
    print(f"      folio: '{datos_finales['folio']}'")
    print(f"      solicitud: '{datos_finales['solicitud']}'")
    print(f"      lista: '{datos_finales['lista']}'")
    print(f"      CADENA COMPLETA: '{datos_finales['cadena_identificacion']}'")
    
    return datos_finales








def cargar_clientes():
    """Carga el archivo Clientes.json"""
    try:
        with open('data/Clientes.json', 'r', encoding='utf-8') as f:
            clientes = json.load(f)
        return clientes
    except Exception as e:
        print(f"⚠️  Error cargando clientes: {e}")
        return []