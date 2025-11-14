"""Generador de Dictámenes PDF - Con soporte para cliente manual
Este archivo genera Dictámenes de Cumplimiento en formato PDF de prueba si se ejecuta directamente."""

import os
import sys
import json
import pandas as pd
from datetime import datetime
import tempfile
import shutil
import traceback

# Importar funciones de carga de datos
try:
    from ArmadoDictamen import (
        cargar_tabla_relacion, 
        cargar_normas, 
        procesar_familias, 
        preparar_datos_familia, 
        cargar_clientes
    )
    print("✅ ArmadoDictamen.py cargado correctamente")
except ImportError as e:
    print(f"❌ Error importando ArmadoDictamen: {e}")
    sys.exit(1)

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

# Constantes para el tamaño de página
from reportlab.lib.pagesizes import letter
LETTER_WIDTH, LETTER_HEIGHT = letter

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
            
        # Reemplazo especial para el texto del dictamen
        if "De conformidad en lo dispuesto" in texto:
            texto = self._reemplazar_texto_dictamen(texto)
        else:
            # Reemplazo normal para otros placeholders
            for key, value in self.datos.items():
                placeholder = f"${{{key}}}"
                if placeholder in texto:
                    texto = texto.replace(placeholder, str(value))
        return texto

    def _reemplazar_texto_dictamen(self, texto):
        """Reemplazo especial para el texto principal del dictamen con campos en negritas"""
        # Obtener los datos necesarios
        cliente = self.datos.get('cliente', '')
        producto = self.datos.get('producto', '')
        pedimento = self.datos.get('pedimento', '')
        fverificacionlarga = self.datos.get('fverificacionlarga', '')
        capitulo = self.datos.get('capitulo', '')
        norma = self.datos.get('norma', '')
        normades = self.datos.get('normades', '')
        
        # Construir el texto del dictamen con campos específicos en negritas
        texto_dictamen = (
            f"De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, 68 y 140 de la Ley de Infraestructura "
            f"de la Calidad; 50 del Reglamento de la Ley Federal de Metrología y Normalización; Punto 2.4.8 Fracción III ACUERDO por "
            f"el que la Secretaría de Economía emite Reglas y criterios de carácter general en materia de comercio exterior; publicado "
            f"en el Diario Oficial de la Federación el 09 de mayo de 2022 y posteriores modificaciones; esta Unidad de Inspección a "
            f"solicitud de la persona moral denominada <b>{cliente}</b> dictamina el Producto: <b>{producto}</b>; que la mercancía importada bajo el "
            f"pedimento aduanal No. <b>{pedimento}</b> de fecha {fverificacionlarga}, fue etiquetada conforme a los requisitos de Información "
            f"Comercial en el capítulo <b>{capitulo}</b> de la Norma Oficial Mexicana <b>{norma}</b> <b>{normades}</b> Cualquier otro requisito "
            f"establecido en la norma referida, es responsabilidad del titular de este Dictamen."
        )
        
        return texto_dictamen

    def agregar_encabezado_pie_pagina(self, canvas, doc):
        """Agrega encabezado, pie de página y numeración a todas las páginas - SOBREESCRITO para usar self.datos"""
        
        canvas.saveState()
        
        width, height = doc.pagesize
        
        # Fondo
        image_path = "img/Fondo.jpeg"
        if os.path.exists(image_path):
            try:
                canvas.drawImage(image_path, 0, 0, width=width, height=height)
            except:
                pass
        
        # Encabezado
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(width/2, height-60, "DICTAMEN DE CUMPLIMIENTO")
        
        canvas.setFont("Helvetica", 10)
        codigo_text = self.datos.get('cadena_identificacion', '')
        if not codigo_text:
            year = self.datos.get('year', '')
            norma = self.datos.get('norma', '')
            folio = self.datos.get('folio', '')
            solicitud = self.datos.get('solicitud', '')
            lista = self.datos.get('lista', '')
            codigo_text = f"{year}049UDC{norma}{folio} Solicitud de Servicio: {year}049USD{norma}{solicitud}-{lista}"
        if len(codigo_text) > 100:
            codigo_text = codigo_text[:100] + "..."
        canvas.drawCentredString(width/2, height-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(width-72, height-50, numeracion)
        
        # Pie de página
        footer_text = "Este Dictamen de Cumplimiento se emitió por medios electrónicos, conforme al oficio de autorización DGN.312.05.2012.106 de fecha 10 de enero de 2012 expedido por la DGN a esta Unidad de Inspección."
        formato_text = "Formato: PT-F-208B-00-3"

        canvas.setFont("Helvetica", 7)

        # Dividir texto en líneas
        lines = []
        words = footer_text.split()
        current_line = ""
        for word in words:
            test_line = current_line + " " + word if current_line else word
            if len(test_line) <= 150:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)

        line_height = 8
        start_y = 60

        for i, line in enumerate(lines):
            text_width = canvas.stringWidth(line, "Helvetica", 7)
            available_width = width - 144
            if text_width < available_width * 0.8:
                x_position = (width - text_width) / 2
            else:
                x_position = 72
            canvas.drawString(x_position, start_y - (i * line_height), line)

        canvas.drawRightString(width - 72, start_y - (len(lines) * line_height) - 4, formato_text)

        canvas.restoreState()

    def crear_estilos(self):
        """Crear estilos que permitan HTML/negritas - VERSIÓN CORREGIDA"""
        # Llamar al método de la clase base primero
        super().crear_estilos()
        
        # Configurar los estilos para permitir HTML de manera segura
        try:
            # Obtener los nombres de los estilos de manera segura
            if hasattr(self.styles, '_styles'):
                # Para ReportLab, los estilos se almacenan en _styles
                for style_name, style_obj in self.styles._styles.items():
                    if hasattr(style_obj, 'allowHtml'):
                        style_obj.allowHtml = True
            elif hasattr(self.styles, 'byName'):
                # Alternativa: usar byName si está disponible
                for style_name, style_obj in self.styles.byName.items():
                    if hasattr(style_obj, 'allowHtml'):
                        style_obj.allowHtml = True
            else:
                # Si no podemos acceder a los estilos, crear uno específico para el texto del dictamen
                print("⚠️  No se pudieron configurar todos los estilos para HTML, continuando...")
                
        except Exception as e:
            print(f"⚠️  Error configurando estilos HTML: {e}")

def generar_dictamenes_completos(directorio_destino, cliente_manual=None, rfc_manual=None):
    """Función principal que genera todos los dictámenes"""
    
    print("🚀 INICIANDO GENERACIÓN DE DICTÁMENES")
    print("="*60)
    
    # Cargar datos
    print("📂 Cargando datos...")
    tabla_datos = cargar_tabla_relacion()
    normas_map, normas_info_completa = cargar_normas()
    
    if not tabla_datos:
        return False, "No se pudieron cargar los datos de la tabla de relación", None
    
    # Procesar familias
    familias = procesar_familias(tabla_datos)
    
    if not familias:
        return False, "No se encontraron familias para procesar", None
    
    # Crear directorio de destino
    os.makedirs(directorio_destino, exist_ok=True)
    print(f"📁 Directorio de destino: {directorio_destino}")
    
    # Informar sobre el cliente manual si se usa
    if cliente_manual:
        print(f"👤 Usando cliente manual: {cliente_manual}")
    
    # Generar dictámenes
    dictamenes_generados = 0
    archivos_creados = []
    
    print(f"\n🛠️  Generando {len(familias)} dictámenes...")
    
    for lista, registros in familias.items():
        print(f"\n📄 Procesando familia LISTA {lista} ({len(registros)} registros)...")
        
        try:
            # Preparar datos para esta familia
            datos = preparar_datos_familia(registros, normas_map, normas_info_completa, cliente_manual, rfc_manual)
            
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
def generar_dictamenes_gui(callback_progreso=None, callback_finalizado=None, cliente_manual=None, rfc_manual=None):
    """Versión para interfaz gráfica con soporte para cliente manual"""
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
        
        # Generar dictámenes con cliente manual si se proporciona
        exito, mensaje, resultado = generar_dictamenes_completos(carpeta_final, cliente_manual, rfc_manual)
        
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
    else:
        print(f"\n❌ {mensaje}")

