"""Generador de Dictámenes PDF con Datos Reales e Imágenes de Etiquetas"""
import os
import sys
import json
import pandas as pd
from datetime import datetime
import traceback

from plantillaPDF import (
    cargar_tabla_relacion,
    cargar_normas,
    cargar_clientes,
    cargar_firmas,
    procesar_familias,
    preparar_datos_familia
)

from DictamenPDF import PDFGenerator

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as RLImage, PageBreak, KeepTogether
)
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle

def obtener_ruta_recurso(ruta_relativa):
    """
    Obtiene la ruta absoluta del recurso, funciona tanto para .py como para .exe.
    PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS.
    """
    try:
        base_path = sys._MEIPASS   # ruta temporal del .exe
    except Exception:
        base_path = os.path.abspath(".")  # ruta local en modo .py

    return os.path.join(base_path, ruta_relativa)

class PDFGeneratorConDatos(PDFGenerator):
    """Subclase que genera PDFs con datos reales y tablas dinámicas
       Evita saltos de página vacíos y calcula correctamente total_pages.
    """

    def __init__(self, datos):
        super().__init__()
        self.datos = datos or {}
        # Calcular total_pages basándose en etiquetas (no añadimos página extra para firmas)
        self.calcular_total_paginas()

    def calcular_total_paginas(self):
        """Calcula el número total de páginas correctamente."""
        etiquetas = self.datos.get('etiquetas_lista', []) or []
        num_etiquetas = len(etiquetas)
        
        paginas_datos = 1
        max_por_pagina = 6
        paginas_etiquetas = (num_etiquetas + max_por_pagina - 1) // max_por_pagina if num_etiquetas > 0 else 0
        
        self.total_pages = paginas_datos + max(1, paginas_etiquetas)
        print(f"   📊 Total de páginas calculado: {self.total_pages}")

    # ---------------- tablas auxiliares ----------------
    def construir_tabla_productos(self):
        print("   📋 Construyendo tabla de productos...")
        tabla_data = [['MARCA', 'CÓDIGO', 'FACTURA', 'CANTIDAD']]
        filas = self.datos.get('tabla_productos', []) or []
        if not filas:
            tabla_data.append(["", "", "", ""])
        else:
            for fila in filas:
                tabla_data.append([
                    str(fila.get('marca', '')),
                    str(fila.get('codigo', '')),
                    str(fila.get('factura', '')),
                    str(fila.get('cantidad', ''))
                ])
        tabla = Table(tabla_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.0*inch])
        tabla.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
        ]))
        return tabla

    def construir_tabla_lote(self):
        total_cantidad = str(self.datos.get('TCantidad', '0 unidades'))
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

    # ---------------- generación ----------------
    def generar_pdf_con_datos(self, output_path):
        """Genera el PDF con datos reales."""
        print(f"   🎯 Generando: {os.path.basename(output_path)}")
        try:
            self.doc = SimpleDocTemplate(
                output_path,
                pagesize=letter,
                topMargin=1.5*inch,
                bottomMargin=1.5*inch,
                leftMargin=0.75*inch,
                rightMargin=0.75*inch
            )

            self.crear_estilos()
            if not hasattr(self, 'elements') or self.elements is None:
                self.elements = []

            self.agregar_primera_pagina_con_datos()
            self.agregar_segunda_pagina_con_etiquetas()

            self.doc.build(self.elements,
                           onFirstPage=self.agregar_encabezado_pie_pagina,
                           onLaterPages=self.agregar_encabezado_pie_pagina)

            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print("   ✅ PDF creado exitosamente")
                return True
            else:
                print("   ❌ El archivo no se creó correctamente")
                return False

        except Exception as e:
            print(f"   ❌ Error generando PDF: {e}")
            traceback.print_exc()
            return False

    # ---------------- páginas ----------------
    def agregar_primera_pagina_con_datos(self):
        print("   📄 Construyendo primera página...")
        texto_fecha_inspeccion = f"<b>Fecha de Inspección:</b> {str(self.datos.get('fverificacion',''))}"
        texto_fecha_emision = f"<b>Fecha de Emisión:</b> {str(self.datos.get('femision',''))}"
        self.elements.append(Paragraph(texto_fecha_inspeccion, self.normal_style))
        self.elements.append(Paragraph(texto_fecha_emision, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        texto_cliente = f"<b>Cliente:</b> {str(self.datos.get('cliente',''))}"
        texto_rfc = f"<b>RFC:</b> {str(self.datos.get('rfc',''))}"
        self.elements.append(Paragraph(texto_cliente, self.normal_style))
        self.elements.append(Paragraph(texto_rfc, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        texto_dictamen = (
            "De conformidad en lo dispuesto en los artículos 53, 56 fracción I, 60 fracción I, 62, 64, "
            "68 y 140 de la Ley de Infraestructura de la Calidad; 50 del Reglamento de la Ley Federal "
            "de Metrología y Normalización; Punto 2.4.8 Fracción III ACUERDO por el que la Secretaría "
            "de Economía emite Reglas y criterios de carácter general en materia de comercio exterior; "
            "publicado en el Diario Oficial de la Federación el 09 de mayo de 2022 y posteriores "
            "modificaciones; esta Unidad de Inspección a solicitud de la persona moral denominada "
            f"<b>{str(self.datos.get('cliente',''))}</b> dictamina el Producto: <b>{str(self.datos.get('producto',''))}</b>; "
            f"que la mercancía importada bajo el pedimento aduanal No. <b>{str(self.datos.get('pedimento',''))}</b> "
            f"de fecha <b>{str(self.datos.get('fverificacionlarga',''))}</b>, fue etiquetada conforme a los requisitos "
            f"de Información Comercial en el capítulo <b>{str(self.datos.get('capitulo',''))}</b> "
            f"de la Norma Oficial Mexicana <b>{str(self.datos.get('norma',''))}</b> <b>{str(self.datos.get('normades',''))}</b>. "
            "Cualquier otro requisito establecido en la norma referida es responsabilidad del titular de este Dictamen."
        )
        self.elements.append(Paragraph(texto_dictamen, self.normal_style))
        self.elements.append(Spacer(1, 0.2 * inch))

        tabla_productos = self.construir_tabla_productos()
        self.elements.append(tabla_productos)
        self.elements.append(Spacer(1, 0.2 * inch))

        tabla_lote = self.construir_tabla_lote()
        self.elements.append(tabla_lote)
        self.elements.append(Spacer(1, 0.2 * inch))

        obs1 = ("<b>OBSERVACIONES:</b> La imagen amparada en el dictamen es una muestra de etiqueta "
                "que aplica para todos los modelos declarados en el presente dictamen; lo anterior fue "
                "constatado durante la inspección.")
        self.elements.append(Paragraph(obs1, self.normal_style))

        obs2 = f"<b>OBSERVACIONES:</b> {str(self.datos.get('obs',''))}"
        self.elements.append(Paragraph(obs2, self.normal_style))
        self.elements.append(Spacer(1, 0.3 * inch))

    def agregar_segunda_pagina_con_etiquetas(self):
        """Genera las páginas de etiquetas con firmas al final."""
        print("   📄 Construyendo página(s) de etiquetas...")

        etiquetas = self.datos.get('etiquetas_lista', []) or []
        
        if not etiquetas:
            print("   ⚠️ No hay etiquetas para mostrar")

        etiquetas_por_fila = 2
        max_por_pagina = 6

        paginas_contenido = []
        total = len(etiquetas)
        total_paginas_etq = (total + max_por_pagina - 1) // max_por_pagina if total else 1

        for pagina_idx in range(total_paginas_etq):
            pagina = []

            inicio = pagina_idx * max_por_pagina
            fin = inicio + max_por_pagina
            etiquetas_pagina = etiquetas[inicio:fin]

            for i in range(0, len(etiquetas_pagina), etiquetas_por_fila):
                fila = etiquetas_pagina[i:i + etiquetas_por_fila]
                imgs = []
                colwidths = []
                for etq in fila:
                    img_bytes = etq.get('imagen_bytes')
                    size_cm = etq.get('tamaño_cm', (5,5))
                    if img_bytes:
                        img_bytes.seek(0)
                        w_cm, h_cm = size_cm
                        img = RLImage(img_bytes,
                                    width=w_cm*0.393701*inch,
                                    height=h_cm*0.393701*inch)
                        imgs.append(img)
                        colwidths.append((w_cm*0.393701 + 0.2)*inch)

                if imgs:
                    tabla = Table([imgs], colWidths=colwidths)
                    tabla.setStyle(TableStyle([
                        ("ALIGN", (0,0), (-1,-1), "CENTER"),
                        ("VALIGN", (0,0), (-1,-1), "MIDDLE")
                    ]))
                    pagina.append(tabla)
                    pagina.append(Spacer(1, 0.15 * inch))

            if pagina_idx == total_paginas_etq - 1:
                pagina.append(Spacer(1, 0.25 * inch))

                imagen_firma1 = obtener_ruta_recurso(self.datos.get('imagen_firma1',''))
                imagen_firma2 = obtener_ruta_recurso(self.datos.get('imagen_firma2',''))

                
                # Columna 1: Primera firma (Inspector)
                col1_elementos = []
                if imagen_firma1 and os.path.exists(imagen_firma1):
                    try:
                        img1 = RLImage(imagen_firma1, width=2.0*inch, height=0.6*inch)
                        col1_elementos.append(img1)
                    except Exception as e:
                        print(f"   ⚠️ Error cargando firma1: {e}")
                
                # Línea divisoria
                col1_elementos.append(Paragraph("_______________________________", self.normal_style))
                
                # Nombre en negritas y centrado
                nombre_insp = str(self.datos.get('nfirma1',''))
                bold_style = ParagraphStyle('BoldCenter', parent=self.normal_style, fontName='Helvetica-Bold', alignment=1)
                col1_elementos.append(Paragraph(nombre_insp, bold_style))
                col1_elementos.append(Paragraph("Inspector", bold_style))
                
                # Columna 3: Segunda firma (Responsable de Supervisión UI)
                col3_elementos = []
                if imagen_firma2 and os.path.exists(imagen_firma2):
                    try:
                        img2 = RLImage(imagen_firma2, width=2.0*inch, height=0.6*inch)
                        col3_elementos.append(img2)
                    except Exception as e:
                        print(f"   ⚠️ Error cargando firma2: {e}")
                
                # Línea divisoria
                col3_elementos.append(Paragraph("_______________________________", self.normal_style))
                
                # Nombre en negritas y centrado
                nombre_sup = str(self.datos.get('nfirma2',''))
                col3_elementos.append(Paragraph(nombre_sup, bold_style))
                col3_elementos.append(Paragraph("Responsable de Supervisión UI", bold_style))
                
                firmas_data = [[col1_elementos, '', col3_elementos]]
                
                firmas_table = Table(firmas_data, colWidths=[2.5*inch, 0.5*inch, 2.5*inch])
                firmas_table.setStyle(TableStyle([
                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                    ('VALIGN',(0,0),(-1,-1),'TOP'),
                ]))
                pagina.append(firmas_table)

            paginas_contenido.append(pagina)

        for idx, pagina in enumerate(paginas_contenido):
            if idx > 0:
                self.elements.append(PageBreak())
            self.elements.extend(pagina)

    def agregar_encabezado_pie_pagina(self, canvas, doc):
        canvas.saveState()
        
        image_path = obtener_ruta_recurso("img/Fondo.jpeg")
        if os.path.exists(image_path):
            try:
                canvas.drawImage(image_path, 0, 0, width=8.5*inch, height=11*inch)
            except:
                pass

        # Encabezado
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(8.5*inch/2, 11*inch-60, "DICTAMEN DE CUMPLIMIENTO")
        
        year = datetime.now().strftime("%y")
        norma = str(self.datos.get('norma', '')).strip()
        folio = str(self.datos.get('folio', '')).strip()
        solicitud = str(self.datos.get('solicitud', '')).strip()
        lista = str(self.datos.get('lista', '')).strip()
        
        # Formato: ${year}049UDC${norma}${folio} Solicitud de Servicio: ${year}049USD${norma}${solicitud}-${lista}
        linea_completa = f"{year}049UDC{norma}{folio}   Solicitud de Servicio: {year}049USD{norma}{solicitud}-{lista}"
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(8.5*inch/2, 11*inch-80, linea_completa)

        # Numeración
        pagina_actual = canvas.getPageNumber()
        numeracion = f"Página {pagina_actual} de {self.total_pages}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(8.5*inch-72, 11*inch-50, numeracion)

        # Pie
        footer_text = ("Este Dictamen de Cumplimiento se emitió por medios electrónicos, conforme al oficio "
                       "de autorización DGN.312.05.2012.106 de fecha 10 de enero de 2012 expedido por la DGN a esta Unidad de Inspección.")
        formato_text = "Formato: PT-F-208B-00-3"
        canvas.setFont("Helvetica", 7)

        words = footer_text.split()
        lines = []
        current_line = ""
        for w in words:
            test = f"{current_line} {w}".strip()
            if len(test) <= 150:
                current_line = test
            else:
                lines.append(current_line)
                current_line = w
        if current_line:
            lines.append(current_line)

        line_height = 8
        start_y = 60
        for i, line in enumerate(lines):
            canvas.drawCentredString(8.5*inch/2, start_y - (i * line_height), line)
        canvas.drawRightString(8.5*inch - 72, start_y - (len(lines) * line_height) - 4, formato_text)

        canvas.restoreState()

# ---------------- resto (funciones auxiliares y flujo) ----------------
def limpiar_nombre_archivo(nombre):
    prohibidos = '\\/:*?"<>|'
    for p in prohibidos:
        nombre = nombre.replace(p, "_")
    return nombre

def generar_dictamenes_completos(directorio_destino, cliente_manual=None, rfc_manual=None):
    print("🚀 INICIANDO GENERACIÓN DE DICTÁMENES")
    print("="*60)

    # Cargar datos
    tabla_datos = cargar_tabla_relacion()
    normas_map, normas_info_completa = cargar_normas()
    clientes_map = cargar_clientes()
    firmas_map = cargar_firmas()

    if tabla_datos is None or tabla_datos.empty:
        return False, "No se pudieron cargar los datos de la tabla de relación", None

    familias = procesar_familias(tabla_datos)
    if not familias:
        return False, "No se encontraron familias para procesar", None

    os.makedirs(directorio_destino, exist_ok=True)
    
    dictamenes_generados = 0
    dictamenes_con_firma = 0
    dictamenes_sin_firma = 0
    dictamenes_error = 0
    
    archivos_creados = []
    sin_firma_detalle = []

    for lista, registros in familias.items():
        print(f"\n📄 Procesando familia LISTA {lista} ({len(registros)} registros)...")
        try:
            datos = preparar_datos_familia(
                registros,
                normas_map,
                normas_info_completa,
                clientes_map,
                firmas_map,
                cliente_manual,
                rfc_manual
            )
            
            if datos is None:
                dictamenes_error += 1
                print(f"   ❌ ERROR: No se pudieron preparar datos para lista {lista}")
                continue

            tiene_firma = datos.get("firma_valida", False)
            
            generador = PDFGeneratorConDatos(datos)
            nombre_archivo = limpiar_nombre_archivo(f"Dictamen_Lista_{lista}.pdf")
            ruta_completa = os.path.join(directorio_destino, nombre_archivo)

            if generador.generar_pdf_con_datos(ruta_completa):
                dictamenes_generados += 1
                archivos_creados.append(ruta_completa)
                
                if tiene_firma:
                    dictamenes_con_firma += 1
                    print(f"   ✅ Creado CON FIRMA: {nombre_archivo}")
                else:
                    dictamenes_sin_firma += 1
                    print(f"   ⚠️ Creado SIN FIRMA: {nombre_archivo}")
                    
                    sin_firma_detalle.append({
                        "lista": lista,
                        "norma": datos.get("norma", ""),
                        "firma_solicitada": datos.get("codigo_firma_solicitado", ""),
                        "razon": datos.get("razon_sin_firma", "Desconocida")
                    })
            else:
                dictamenes_error += 1
                print(f"   ❌ Error creando dictamen para lista {lista}")

        except Exception as e:
            dictamenes_error += 1
            print(f"   ❌ Error en familia {lista}: {e}")
            traceback.print_exc()
            continue

    print("\n" + "="*60)
    print("📊 RESUMEN DE GENERACIÓN")
    print("="*60)
    print(f"✅ Total generados: {dictamenes_generados}/{len(familias)}")
    print(f"✅ Con firma válida: {dictamenes_con_firma}")
    print(f"⚠️  Sin firma: {dictamenes_sin_firma}")
    
    if dictamenes_error > 0:
        print(f"❌ Con errores: {dictamenes_error}")
    
    if sin_firma_detalle:
        print("\n" + "="*60)
        print("⚠️  DICTÁMENES SIN FIRMA - DETALLE")
        print("="*60)
        for item in sin_firma_detalle:
            print(f"\n📄 Lista: {item['lista']}")
            print(f"   Norma: {item['norma']}")
            print(f"   Firma solicitada: {item['firma_solicitada']}")
            print(f"   Razón: {item['razon']}")
    
    print("\n" + "="*60)

    resultado = {
        'directorio': directorio_destino,
        'total_generados': dictamenes_generados,
        'con_firma': dictamenes_con_firma,
        'sin_firma': dictamenes_sin_firma,
        'con_error': dictamenes_error,
        'total_familias': len(familias),
        'archivos': archivos_creados,
        'sin_firma_detalle': sin_firma_detalle
    }
    
    mensaje = f"Se generaron {dictamenes_generados} dictámenes ({dictamenes_con_firma} con firma, {dictamenes_sin_firma} sin firma)"
    success = dictamenes_generados > 0
    return success, mensaje if success else "No se pudo generar ningún dictamen", resultado

def generar_dictamenes_gui(callback_progreso=None, callback_finalizado=None, cliente_manual=None, rfc_manual=None):
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        directorio_destino = filedialog.askdirectory(title="Seleccione dónde guardar los dictámenes")
        root.destroy()
        if not directorio_destino:
            if callback_finalizado:
                callback_finalizado(False, "Operación cancelada por el usuario", None)
            return False, "Operación cancelada", None

        carpeta_final = os.path.join(directorio_destino, f"Dictamenes_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        if callback_progreso:
            callback_progreso(10, "Iniciando...")
        exito, mensaje, resultado = generar_dictamenes_completos(carpeta_final, cliente_manual, rfc_manual)
        if callback_progreso:
            callback_progreso(100, mensaje)
        if callback_finalizado:
            callback_finalizado(exito, mensaje, resultado)
        return exito, mensaje, resultado

    except Exception as e:
        traceback.print_exc()
        if callback_finalizado:
            callback_finalizado(False, str(e), None)
        return False, str(e), None

if __name__ == "__main__":
    carpeta_prueba = "dictamenes_prueba"
    exito, mensaje, resultado = generar_dictamenes_completos(carpeta_prueba)
    if exito:
        print(f"\n🎉 {mensaje}")
        print(f"📁 Ubicación: {resultado['directorio']}")
    else:
        print(f"\n❌ {mensaje}")
