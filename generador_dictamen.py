"""
Generador de Dictámenes PDF - Versión Mejorada
- Mejor manejo de archivos y rutas
- Más información de depuración
"""

import os
import sys
import json
import pandas as pd
from datetime import datetime
from collections import defaultdict
import tempfile
import shutil
import traceback

# Importar tu plantilla
try:
    from DictamenPDF import PDFGenerator
    print("✅ DictamenPDF.py cargado correctamente")
except ImportError as e:
    print(f"❌ Error importando DictamenPDF: {e}")
    sys.exit(1)

# Importaciones de ReportLab
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

class PDFGeneratorConDatos(PDFGenerator):
    """Subclase que reemplaza placeholders con datos reales"""
    
    def __init__(self, datos):
        super().__init__()
        self.datos = datos
    
    def generar_pdf_con_datos(self, output_path):
        """Genera el PDF con datos reales"""
        print(f"🎯 Generando: {os.path.basename(output_path)}")
        
        try:
            # Configurar documento
            self.doc = SimpleDocTemplate(
                output_path,
                pagesize=letter,
                topMargin=1.5*inch,
                bottomMargin=1.5*inch,
                leftMargin=0.75*inch,
                rightMargin=0.75*inch
            )
            
            # Crear estilos y contenido
            self.crear_estilos()
            self.agregar_primera_pagina()
            self.agregar_segunda_pagina()
            
            # Reemplazar placeholders ANTES de construir
            self._reemplazar_todos_los_placeholders()
            
            # Construir PDF
            self.doc.build(
                self.elements,
                onFirstPage=self.agregar_encabezado_pie_pagina,
                onLaterPages=self.agregar_encabezado_pie_pagina
            )
            
            # Verificar que el archivo se creó
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"✅ PDF creado exitosamente: {os.path.basename(output_path)}")
                return True
            else:
                print(f"❌ El archivo no se creó o está vacío: {output_path}")
                return False
                
        except Exception as e:
            print(f"❌ Error generando PDF: {e}")
            traceback.print_exc()
            return False

    def _reemplazar_todos_los_placeholders(self):
        """Reemplaza todos los placeholders en los elementos"""
        nuevos_elementos = []
        
        for elemento in self.elements:
            if hasattr(elemento, 'text'):
                # Es un Paragraph - reemplazar texto
                texto_original = elemento.text
                texto_reemplazado = self._reemplazar_texto(texto_original)
                
                # Crear nuevo Paragraph con estilo preservado
                nuevo_para = Paragraph(texto_reemplazado, elemento.style)
                nuevos_elementos.append(nuevo_para)
            else:
                # Mantener otros elementos (Spacer, Table, etc.)
                nuevos_elementos.append(elemento)
        
        self.elements = nuevos_elementos

    def _reemplazar_texto(self, texto):
        """Reemplaza placeholders en texto"""
        if not texto:
            return texto
            
        for key, value in self.datos.items():
            placeholder = f"${{{key}}}"
            if placeholder in texto:
                texto = texto.replace(placeholder, str(value))
        return texto

# FUNCIONES MEJORADAS DE CARGA DE DATOS
def cargar_tabla_relacion():
    """Carga el archivo tabla_de_relacion_json con mejor manejo de errores"""
    print("🔍 Buscando tabla_de_relacion_json...")
    
    # Lista completa de posibles ubicaciones
    posibles_rutas = [
        'data/tabla_de_relacion.json'     # En directorio actual con extensión
    ]
    
    # Verificar si la carpeta data existe
    if not os.path.exists('data'):
        print("❌ La carpeta 'data' no existe en el directorio actual")
        print("📁 Creando carpeta 'data'...")
        os.makedirs('data', exist_ok=True)
    
    # Buscar el archivo
    archivo_encontrado = None
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            archivo_encontrado = ruta
            print(f"✅ Archivo encontrado: {ruta}")
            break
    
    if not archivo_encontrado:
        print("❌ No se encontró tabla_de_relacion_json en ninguna ubicación:")
        for ruta in posibles_rutas:
            existe = "✅" if os.path.exists(ruta) else "❌"
            print(f"   {existe} {ruta}")
        return None
    
    # Intentar cargar el archivo
    try:
        with open(archivo_encontrado, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        
        print(f"✅ Tabla cargada exitosamente: {len(datos)} registros")
        return datos
        
    except json.JSONDecodeError as e:
        print(f"❌ Error decodificando JSON: {e}")
        # Intentar cargar como texto plano para diagnóstico
        try:
            with open(archivo_encontrado, 'r', encoding='utf-8') as f:
                contenido = f.read()
                print(f"📄 Contenido del archivo (primeros 500 caracteres):")
                print(contenido[:500])
        except:
            pass
        return None
        
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
                
                # Crear mapeo de números a códigos de norma
                normas_map = {}
                
                if isinstance(normas_data, list):
                    print("📝 Procesando lista de normas...")
                    
                    for norma_item in normas_data:
                        if isinstance(norma_item, dict):
                            codigo_norma = norma_item.get('NOM', '')
                            
                            # Extraer números del código de norma para crear mapeos
                            # Ejemplo: "NOM-004-SE-2021" → extraer "004" y "4"
                            numeros_en_norma = []
                            
                            # Buscar números en el código de norma
                            import re
                            numeros = re.findall(r'\d+', codigo_norma)
                            for num in numeros:
                                # Agregar el número tal cual (ej: "004")
                                numeros_en_norma.append(num)
                                # También agregar sin ceros a la izquierda (ej: "4")
                                if num.startswith('0'):
                                    numeros_en_norma.append(num.lstrip('0'))
                            
                            # Crear entradas en el mapeo para cada número encontrado
                            for num in numeros_en_norma:
                                if num:  # Asegurarse de que no esté vacío
                                    normas_map[num] = codigo_norma
                            
                            # También mapear el código completo a sí mismo por si acaso
                            normas_map[codigo_norma] = codigo_norma
                    
                    print(f"✅ Mapeo de normas creado: {len(normas_map)} entradas")
                    
                
                # Mostrar algunas normas para verificación
                print("📋 Ejemplo de mapeo de normas:")
                normas_mostradas = 0
                for key, value in normas_map.items():
                    if len(key) <= 3:  # Mostrar solo mapeos con claves cortas (números)
                        print(f"   - {key} → {value}")
                        normas_mostradas += 1
                        if normas_mostradas >= 10:
                            break
                
                return normas_map
                
            except Exception as e:
                print(f"❌ Error cargando {ruta}: {e}")
                import traceback
                traceback.print_exc()
    
    print("⚠️  No se encontró Normas.json, usando valores por defecto")
    return {
        "4": "NOM-004-SE-2021", 
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





def cargar_clientes():
    """Carga el archivo Clientes.json"""
    print("🔍 Buscando Clientes.json...")
    
    posibles_rutas = [
        'data/Clientes.json',
        'Clientes.json',
        '../data/Clientes.json'
    ]
    
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            try:
                with open(ruta, 'r', encoding='utf-8') as f:
                    clientes = json.load(f)
                print(f"✅ Clientes cargados: {len(clientes)} clientes")
                return clientes
            except Exception as e:
                print(f"❌ Error cargando {ruta}: {e}")
    
    print("⚠️  No se encontró Clientes.json, usando valores por defecto")
    return ""

def verificar_estructura_datos():
    """Verifica que todos los archivos necesarios existan"""
    print("\n" + "="*50)
    print("VERIFICACIÓN DE ESTRUCTURA DE DATOS")
    print("="*50)
    
    # Verificar tabla_de_relacion_json
    tabla = cargar_tabla_relacion()
    if tabla is None:
        print("❌ CRÍTICO: No se pudo cargar tabla_de_relacion_json")
        return False
    
    # Verificar Normas.json
    normas = cargar_normas()
    
    # Verificar Clientes.json  
    clientes = cargar_clientes()
    
    print("✅ Estructura de datos verificada")
    return True

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

def preparar_datos_familia(registros, normas_map, clientes_list):
    """Prepara datos para una familia específica"""
    if not registros:
        return None
        
    primer_registro = registros[0]
    
    # Información básica
    year = datetime.now().strftime("%y")
    norma_uva = primer_registro.get('NORMA UVA', '')
    folio = str(primer_registro.get('FOLIO', ''))
    solicitud = str(primer_registro.get('SOLICITUD', ''))
    lista = str(primer_registro.get('LISTA', ''))
    
    # Mapear norma - CORREGIDO para usar el mapeo de normas
    norma = "NOM-001"  # Valor por defecto
    if not pd.isna(norma_uva) and norma_uva != '':
        norma_str = str(int(norma_uva)) if isinstance(norma_uva, (int, float)) else str(norma_uva)
        
        # Buscar en el mapa de normas
        if norma_str in normas_map:
            norma = normas_map[norma_str]
            print(f"   📋 Norma UVA {norma_str} → {norma}")
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
    
    # Cliente y RFC
    marca = primer_registro.get('MARCA', '')
    cliente, rfc = marca, ""
    if not pd.isna(marca) and marca != '':
        marca_upper = marca.upper()
        for cliente_info in clientes_list:
            cliente_marca = cliente_info.get('MARCA', '').upper()
            if marca_upper == cliente_marca or marca_upper in cliente_marca:
                cliente = cliente_info.get('CLIENTE', marca)
                rfc = cliente_info.get('RFC', '')
                break
    
    print(f"   👤 Cliente: {marca} → {cliente}")
    
    # Producto
    producto = primer_registro.get('DESCRIPCION', 'Producto no especificado')
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
    
    return {
        'year': year,
        'norma': norma,
        'folio': folio,
        'solicitud': solicitud,
        'lista': lista,
        'fverificacion': fverificacion,
        'femision': femision,
        'fverificacionlarga': fverificacion,  # Simplificado
        'cliente': cliente,
        'rfc': rfc,
        'producto': producto,
        'pedimento': str(primer_registro.get('PEDIMENTO', '')),
        'capitulo': '4',
        'normades': 'ESPECIFICACIONES DE SEGURIDAD',
        'rowMarca': marca if not pd.isna(marca) and marca != '' else "",
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


def generar_dictamenes_completos(directorio_destino):
    """Función principal que genera todos los dictámenes"""
    
    print("🚀 INICIANDO GENERACIÓN DE DICTÁMENES")
    print("="*60)
    
    # Primero verificar la estructura de datos
    if not verificar_estructura_datos():
        return False, "Error en la estructura de datos. Verifique los archivos requeridos.", None
    
    # Cargar datos nuevamente para el procesamiento
    tabla_datos = cargar_tabla_relacion()
    normas_map = cargar_normas()
    clientes_list = cargar_clientes()
    
    if not tabla_datos:
        return False, "No se pudieron cargar los datos de la tabla de relación", None
    
    # Procesar familias
    familias = procesar_familias(tabla_datos)
    
    if not familias:
        return False, "No se encontraron familias para procesar", None
    
    # Crear directorio de destino
    os.makedirs(directorio_destino, exist_ok=True)
    print(f"📁 Directorio de destino: {directorio_destino}")
    
    # Generar dictámenes
    dictamenes_generados = 0
    archivos_creados = []
    
    print(f"\n🛠️  Generando {len(familias)} dictámenes...")
    
    for lista, registros in familias.items():
        print(f"\n📄 Procesando familia LISTA {lista} ({len(registros)} registros)...")
        
        try:
            # Preparar datos
            datos = preparar_datos_familia(registros, normas_map, clientes_list)
            
            if not datos:
                print(f"   ⚠️  No se pudieron preparar datos para lista {lista}")
                continue
            
            # Generar PDF
            generador = PDFGeneratorConDatos(datos)
            nombre_archivo = f"Dictamen_Lista_{lista}.pdf"
            ruta_completa = os.path.join(directorio_destino, nombre_archivo)
            
            if generador.generar_pdf_con_datos(ruta_completa):
                dictamenes_generados += 1
                archivos_creados.append(ruta_completa)
                print(f"   ✅ Creado: {nombre_archivo}")
            else:
                print(f"   ❌ Error creando dictamen para lista {lista}")
                
        except Exception as e:
            print(f"   ❌ Error en familia {lista}: {e}")
            traceback.print_exc()
            continue
    
    # Resultado final
    resultado = {
        'directorio': directorio_destino,
        'total_generados': dictamenes_generados,
        'total_familias': len(familias),
        'archivos': archivos_creados
    }
    
    mensaje = f"Se generaron {dictamenes_generados} de {len(familias)} dictámenes"
    
    if dictamenes_generados == 0:
        return False, "No se pudo generar ningún dictamen", resultado
    else:
        return True, mensaje, resultado

# Función para la GUI
def generar_dictamenes_gui(callback_progreso=None, callback_finalizado=None):
    """Versión para interfaz gráfica"""
    try:
        if callback_progreso:
            callback_progreso(10, "Solicitando ubicación...")
        
        # Importar aquí para evitar problemas de dependencia
        import tkinter as tk
        from tkinter import filedialog
        
        # Crear ventana temporal para el diálogo
        root = tk.Tk()
        root.withdraw()  # Ocultar ventana principal
        
        directorio_destino = filedialog.askdirectory(
            title="Seleccione dónde guardar los dictámenes"
        )
        
        root.destroy()
        
        if not directorio_destino:
            if callback_finalizado:
                callback_finalizado(False, "Operación cancelada por el usuario", None)
            return False, "Operación cancelada", None
        
        # Crear subcarpeta con fecha
        from datetime import datetime
        carpeta_final = os.path.join(directorio_destino, f"Dictamenes_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        
        if callback_progreso:
            callback_progreso(30, "Verificando estructura de datos...")
        
        # Generar dictámenes
        exito, mensaje, resultado = generar_dictamenes_completos(carpeta_final)
        
        if callback_progreso:
            callback_progreso(100, mensaje)
        
        if callback_finalizado:
            callback_finalizado(exito, mensaje, resultado)
        
        return exito, mensaje, resultado
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print(f"❌ Error en generador GUI: {error_msg}")
        traceback.print_exc()
        
        if callback_finalizado:
            callback_finalizado(False, error_msg, None)
        return False, error_msg, None

if __name__ == "__main__":
    print("=" * 60)
    print("   GENERADOR DE DICTAMENES - PRUEBA DIRECTA")
    print("=" * 60)
    
    # Prueba directa
    carpeta_prueba = "dictamenes_prueba"
    exito, mensaje, resultado = generar_dictamenes_completos(carpeta_prueba)
    
    if exito:
        print(f"\n🎉 {mensaje}")
        print(f"📁 Ubicación: {resultado['directorio']}")
        
        # Listar archivos creados
        print("\n📄 Archivos creados:")
        for archivo in resultado['archivos']:
            print(f"   • {os.path.basename(archivo)}")
            
        # Verificar que los archivos existen
        print("\n🔍 Verificando archivos...")
        for archivo in resultado['archivos']:
            existe = "✅" if os.path.exists(archivo) else "❌"
            print(f"   {existe} {os.path.basename(archivo)}")
    else:
        print(f"\n❌ {mensaje}")