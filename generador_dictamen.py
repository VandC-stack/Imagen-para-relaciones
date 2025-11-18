"""Generador de Dictámenes PDF - VERSIÓN CORREGIDA CON TABLA DINÁMICA"""
import os
import sys
import json
import pandas as pd
from datetime import datetime
import traceback

# Importar funciones de carga de datos
try:
    from ArmadoDictamen import (
        cargar_tabla_relacion, 
        cargar_normas,
        cargar_clientes,  # Agregado import de cargar_clientes
        procesar_familias, 
        preparar_datos_familia
    )
    print("✅ ArmadoDictamen.py cargado correctamente")
except ImportError as e:
    print(f"❌ Error importando ArmadoDictamen: {e}")
    sys.exit(1)

# Importar tu plantilla base
try:
    from DictamenPDF import PDFGenerator
    print("✅ DictamenPDF.py cargado correctamente")
except ImportError as e:
    print(f"❌ Error importando DictamenPDF: {e}")
    sys.exit(1)

# Importaciones de ReportLab
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors

class PDFGeneratorConDatos(PDFGenerator):
    """Subclase que genera PDFs con datos reales y tablas dinámicas"""
    
    def __init__(self, datos):
        super().__init__()
        self.datos = datos
    
    def construir_tabla_productos(self):
        """Construye la tabla REAL usando tabla_productos de datos"""
        print("   📋 Construyendo tabla de productos...")

        # Encabezados
        tabla_data = [['MARCA', 'CÓDIGO', 'FACTURA', 'CANTIDAD']]

        filas = self.datos.get('tabla_productos', [])
        if not filas:
            print("   ⚠ No hay filas de productos")
            tabla_data.append(["", "", "", ""])
        else:
            for fila in filas:
                marca = fila.get('marca', '')
                codigo = fila.get('codigo', '')
                factura = fila.get('factura', '')
                cantidad = fila.get('cantidad', '')

                tabla_data.append([marca, codigo, factura, str(cantidad)])

        print(f"   ✅ {len(tabla_data)-1} filas agregadas a la tabla")

        tabla = Table(tabla_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.0*inch])

        tabla.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
        ]))

        return tabla

    def construir_tabla_lote(self):
        """Construye la tabla de tamaño del lote"""
        total_cantidad = self.datos.get('TCantidad', '0 unidades')
        
        tabla_data = [['TAMAÑO DEL LOTE', total_cantidad]]
        
        tabla = Table(tabla_data, colWidths=[4.5*inch, 1.5*inch])
        tabla.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (0,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
        ]))
        
        return tabla
    
    def generar_pdf_con_datos(self, output_path):
        """Genera el PDF con datos reales"""
        print(f"   🎯 Generando: {os.path.basename(output_path)}")
        
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
            self.agregar_primera_pagina_con_datos()
            self.agregar_segunda_pagina()
            
            # Construir PDF
            self.doc.build(
                self.elements,
                onFirstPage=self.agregar_encabezado_pie_pagina,
                onLaterPages=self.agregar_encabezado_pie_pagina
            )
            
            # Verificar creación
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"   ✅ PDF creado exitosamente")
                return True
            else:
                print(f"   ❌ El archivo no se creó correctamente")
                return False
                
        except Exception as e:
            print(f"   ❌ Error generando PDF: {e}")
            traceback.print_exc()
            return False
    
    def agregar_primera_pagina_con_datos(self):
        """Construye la primera página con datos reales y el texto original completo"""
        print("   📄 Construyendo primera página...")
        
        print(f"   🔍 DEBUG - Norma: {self.datos.get('norma', 'NO ENCONTRADA')}")
        print(f"   🔍 DEBUG - Normades: {self.datos.get('normades', 'NO ENCONTRADA')}")
        print(f"   🔍 DEBUG - Cliente: {self.datos.get('cliente', 'NO ENCONTRADO')}")
        print(f"   🔍 DEBUG - RFC: {self.datos.get('rfc', 'NO ENCONTRADO')}")

        # Fechas
        texto_fecha_inspeccion = f"<b>Fecha de Inspección:</b> {self.datos.get('fverificacion', '')}"
        self.elements.append(Paragraph(texto_fecha_inspeccion, self.normal_style))

        texto_fecha_emision = f"<b>Fecha de Emisión:</b> {self.datos.get('femision', '')}"
        self.elements.append(Paragraph(texto_fecha_emision, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        # Cliente y RFC
        texto_cliente = f"<b>Cliente:</b> {self.datos.get('cliente', '')}"
        self.elements.append(Paragraph(texto_cliente, self.normal_style))

        texto_rfc = f"<b>RFC:</b> {self.datos.get('rfc', '')}"
        self.elements.append(Paragraph(texto_rfc, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        # Texto del dictamen con NORMA y NORMADES
        texto_dictamen = (
            "De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, "
            "68 y 140 de la Ley de Infraestructura de la Calidad; 50 del Reglamento de la Ley Federal "
            "de Metrología y Normalización; Punto 2.4.8 Fracción III ACUERDO por el que la Secretaría "
            "de Economía emite Reglas y criterios de carácter general en materia de comercio exterior; "
            "publicado en el Diario Oficial de la Federación el 09 de mayo de 2022 y posteriores "
            "modificaciones; esta Unidad de Inspección a solicitud de la persona moral denominada "
            f"<b>{self.datos.get('cliente', '')}</b> dictamina el Producto: <b>{self.datos.get('producto', '')}</b>; "
            f"que la mercancía importada bajo el pedimento aduanal No. <b>{self.datos.get('pedimento', '')}</b> "
            f"de fecha <b>{self.datos.get('fverificacionlarga', '')}</b>, fue etiquetada conforme a los requisitos "
            f"de Información Comercial en el capítulo <b>{self.datos.get('capitulo', '')}</b> "
            f"de la Norma Oficial Mexicana <b>{self.datos.get('norma', '')}</b> <b>{self.datos.get('normades', '')}</b>. "
            "Cualquier otro requisito establecido en la norma referida es responsabilidad del titular de este Dictamen."
        )

        self.elements.append(Paragraph(texto_dictamen, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        # TABLA DE PRODUCTOS
        tabla_productos = self.construir_tabla_productos()
        self.elements.append(tabla_productos)
        self.elements.append(Spacer(1, 0.2 * inch))

        # TAMAÑO DEL LOTE
        tabla_lote = self.construir_tabla_lote()
        self.elements.append(tabla_lote)
        self.elements.append(Spacer(1, 0.2 * inch))

        # OBSERVACIONES
        obs1 = (
            "<b>OBSERVACIONES:</b> La imagen amparada en el dictamen es una muestra de etiqueta "
            "que aplica para todos los modelos declarados en el presente dictamen; lo anterior fue "
            "constatado durante la inspección."
        )
        self.elements.append(Paragraph(obs1, self.normal_style))

        obs2 = f"<b>OBSERVACIONES:</b> {self.datos.get('obs', '')}"
        self.elements.append(Paragraph(obs2, self.normal_style))
        self.elements.append(Spacer(1, 0.3 * inch))


    def agregar_encabezado_pie_pagina(self, canvas, doc):
        """Sobrescribe el método para agregar encabezado y pie con datos reales"""
        canvas.saveState()
        
        # Fondo
        image_path = "img/Fondo.jpeg"
        if os.path.exists(image_path):
            try:
                canvas.drawImage(image_path, 0, 0, width=8.5*inch, height=11*inch)
            except:
                pass
        
        # Encabezado
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(8.5*inch/2, 11*inch-60, "DICTAMEN DE CUMPLIMIENTO")
        
        canvas.setFont("Helvetica", 10)
        codigo_text = self.datos.get('cadena_identificacion', '')
        canvas.drawCentredString(8.5*inch/2, 11*inch-80, codigo_text)
        
        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(8.5*inch-72, 11*inch-50, numeracion)
        
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
            canvas.drawCentredString(8.5*inch/2, start_y - (i * line_height), line)
        
        canvas.drawRightString(8.5*inch - 72, start_y - (len(lines) * line_height) - 4, formato_text)
        
        canvas.restoreState()

def limpiar_nombre_archivo(nombre):
    """Reemplaza caracteres inválidos en nombres de archivos."""
    prohibidos = '\\/:*?"<>|'
    for p in prohibidos:
        nombre = nombre.replace(p, "_")
    return nombre

def generar_dictamenes_completos(directorio_destino, cliente_manual=None, rfc_manual=None):
    """Función principal que genera todos los dictámenes"""
    
    print("🚀 INICIANDO GENERACIÓN DE DICTÁMENES")
    print("="*60)
    
    # Cargar datos
    print("📂 Cargando datos...")
    tabla_datos = cargar_tabla_relacion()
    normas_map, normas_info_completa = cargar_normas()
    clientes_map = cargar_clientes()  # Agregado carga de clientes
    
    if tabla_datos is None or tabla_datos.empty:
        return False, "No se pudieron cargar los datos de la tabla de relación", None
    
    # Procesar familias
    familias = procesar_familias(tabla_datos)
    
    if familias is None or len(familias) == 0:
        return False, "No se encontraron familias para procesar", None
    
    # Crear directorio de destino
    os.makedirs(directorio_destino, exist_ok=True)
    print(f"📁 Directorio de destino: {directorio_destino}")
    
    # Informar sobre el cliente manual si se usa
    if cliente_manual:
        print(f"👤 Usando cliente manual: {cliente_manual}")
    if rfc_manual:
        print(f"🆔 Usando RFC manual: {rfc_manual}")
    
    # Generar dictámenes
    dictamenes_generados = 0
    archivos_creados = []
    
    print(f"\n🛠️  Generando {len(familias)} dictámenes...")
    
    for lista, registros in familias.items():
        print(f"\n📄 Procesando familia LISTA {lista} ({len(registros)} registros)...")
        
        try:
            datos = preparar_datos_familia(
                registros, 
                normas_map, 
                normas_info_completa, 
                clientes_map,
                cliente_manual, 
                rfc_manual
            )
            
            if datos is None:
                print(f"   ⚠️  No se pudieron preparar datos para lista {lista}")
                continue
            
            # Generar PDF
            generador = PDFGeneratorConDatos(datos)
            nombre_archivo = limpiar_nombre_archivo(f"Dictamen_Lista_{lista}.pdf")

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
        root.withdraw()
        
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
    print("="*60)
    print("   GENERADOR DE DICTAMENES - PRUEBA DIRECTA")
    print("="*60)
    
    # Prueba directa
    carpeta_prueba = "dictamenes_prueba"
    exito, mensaje, resultado = generar_dictamenes_completos(carpeta_prueba)
    
    if exito:
        print(f"\n🎉 {mensaje}")
        print(f"📁 Ubicación: {resultado['directorio']}")
    else:
        print(f"\n❌ {mensaje}")
