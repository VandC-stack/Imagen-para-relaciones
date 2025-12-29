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
        """Calcula correctamente las páginas según modo y estructura final."""

        modo = self.datos.get("modo_insertado", "etiqueta")

        # --- SIEMPRE existe HOJA 1 = DATOS ---
        paginas = 1

        # -------------------------------------------------------------------
        # MODO: PEGADO DE EVIDENCIA
        # -------------------------------------------------------------------
        if modo == "evidencia":
            print("📌 MODO EVIDENCIA → Datos + Evidencia + Firmas")
            paginas += 1            # Hoja de evidencia
            paginas += 1            # Hoja de firmas
            self.total_pages = paginas
            return

        # -------------------------------------------------------------------
        # MODO: MIXTO (ULTA con NOM-024)
        # -------------------------------------------------------------------
        if modo == "mixto":
            print("📌 MODO MIXTO → Datos + Mixta + Firmas")
            paginas += 1            # Hoja mixta
            paginas += 1            # Firmas
            self.total_pages = paginas
            return

        # -------------------------------------------------------------------
        # MODO ETIQUETADO NORMAL / BASE ETIQUETAS
        # -------------------------------------------------------------------
        if modo in ("etiqueta", "base_etiquetado"):
            etiquetas = self.datos.get("etiquetas_lista", []) or []
            
            # Si no hay etiquetas, solo hay DATOS + FIRMAS = 2 páginas
            if not etiquetas:
                print(f"📌 MODO ETIQUETA → SIN ETIQUETAS (solo Datos + Firmas)")
                paginas += 1           # Hoja de firmas
                self.total_pages = paginas
                return
            
            # Si hay etiquetas, calcular páginas de etiquetas
            max_por_pagina = 6
            paginas_etq = (len(etiquetas) + max_por_pagina - 1) // max_por_pagina

            print(f"📌 MODO ETIQUETA → {paginas_etq} páginas de etiquetas ({len(etiquetas)} etiquetas)")

            paginas += paginas_etq     # Agregar páginas de etiquetas
            paginas += 1               # Hoja de firmas

            self.total_pages = paginas
            return

        # -------------------------------------------------------------------
        # FALLBACK (por si llega un modo desconocido)
        # -------------------------------------------------------------------
        print(f"⚠️ MODO DESCONOCIDO: {modo}, asignando modo etiqueta")
        self.total_pages = 2  # Datos + Firmas mínimo

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

            modo = self.datos.get("modo_insertado", "etiqueta")

            # 🚀 RUTEAMOS SEGÚN RAZÓN SOCIAL
            if modo == "evidencia":
                print("   📌 MODO: SOLO EVIDENCIA")
                self.agregar_hoja_evidencia()

            elif modo == "mixto":
                print("   📌 MODO: MIXTO (EVIDENCIA + ETIQUETAS EN UNA HOJA)")
                self.agregar_hoja_mixta()

            elif modo == "etiqueta":
                # agregar_segunda_pagina_con_etiquetas devolverá True si ya colocó las firmas
                firmas_colocadas = self.agregar_segunda_pagina_con_etiquetas()
                if not firmas_colocadas:
                    # Agregar firmas en página separada
                    self.agregar_hoja_firmas()

            elif modo == "base_etiquetado":
                print("   📌 MODO: BASE DE ETIQUETADO (Decathlon)")
                firmas_colocadas = self.agregar_segunda_pagina_con_etiquetas()
                if not firmas_colocadas:
                    self.agregar_hoja_firmas()

            else:
                print(f"   ⚠️ Modo desconocido: {modo}, se usa modo etiqueta.")
                self.agregar_segunda_pagina_con_etiquetas()


            # Use NumberedCanvas to ensure accurate "Página X de Y" numeration
            from DictamenPDF import NumberedCanvas
            self.doc.build(self.elements,
                           onFirstPage=self.agregar_encabezado_pie_pagina,
                           onLaterPages=self.agregar_encabezado_pie_pagina,
                           canvasmaker=NumberedCanvas)

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

            paginas_contenido.append(pagina)

        for idx, pagina in enumerate(paginas_contenido):
            # If this is not the first etiqueta page, add a page break
            if idx > 0:
                self.elements.append(PageBreak())

            # If we're on the last etiqueta page AND there are less than 5 etiquetas,
            # append the firmas flowables to this page so they remain on the same page
            if idx == (len(paginas_contenido) - 1) and len(etiquetas) < 5:
                print("   📌 Firmas mostradas ABAJO de etiquetas (menos de 5 etiquetas)")
                # add a small spacer then the firmas
                pagina.append(Spacer(1, 0.15 * inch))
                for e in self._get_firmas_elements():
                    pagina.append(e)
                # extend the page content (no extra PageBreak)
                self.elements.extend(pagina)
                return True

            # Otherwise just extend the page content normally
            self.elements.extend(pagina)

        # If we reach here, firmas were not placed; caller should add a separate firmas page
        print("   📌 Firmas mostradas en PÁGINA SEPARADA (5+ etiquetas)")
        return False

    # Agregar hoja para pegado de evidencias fotograficas #
    def agregar_hoja_evidencia(self):
        """Hoja en blanco para evidencia + hoja de firmas."""
        print("   📄 Generando hoja en blanco para evidencia...")

        # HOJA 2 – Evidencia
        self.elements.append(PageBreak())
        self.elements.append(Spacer(1, 0.25 * inch))

        evidencias = self.datos.get('evidencias_lista', []) or []

        if not evidencias:
            # Si no hay evidencias asignadas, mantener el placeholder informativo
            self.elements.append(Paragraph(
                "<b>${IMAGEN}</b>",
                ParagraphStyle('Center', parent=self.normal_style, alignment=1, fontSize=12)
            ))
        else:
            # Insertar cada evidencia. `evidencias` puede ser lista de rutas (str) o de BytesIO/dicts
            from io import BytesIO
            for ev in evidencias:
                try:
                    if isinstance(ev, str):
                        if os.path.exists(ev):
                            img = RLImage(ev, width=4.5*inch, height=4.5*inch)
                            self.elements.append(img)
                            self.elements.append(Spacer(1, 0.25 * inch))
                        else:
                            # ruta no existe -> intentar si es JSON-like dict con imagen_bytes
                            continue
                    elif isinstance(ev, dict):
                        img_bytes = ev.get('imagen_bytes') or ev.get('imagen_path_bytes')
                        if img_bytes:
                            if hasattr(img_bytes, 'seek'):
                                img_bytes.seek(0)
                                img = RLImage(img_bytes, width=4.5*inch, height=4.5*inch)
                                self.elements.append(img)
                                self.elements.append(Spacer(1, 0.25 * inch))
                        else:
                            # maybe it's a path stored under 'imagen_path'
                            p = ev.get('imagen_path')
                            if p and os.path.exists(p):
                                img = RLImage(p, width=4.5*inch, height=4.5*inch)
                                self.elements.append(img)
                                self.elements.append(Spacer(1, 0.25 * inch))
                    else:
                        # assume BytesIO-like
                        try:
                            ev.seek(0)
                            img = RLImage(ev, width=4.5*inch, height=4.5*inch)
                            self.elements.append(img)
                            self.elements.append(Spacer(1, 0.25 * inch))
                        except Exception:
                            continue
                except Exception:
                    continue

        # HOJA 3 – Firmas
        self.agregar_hoja_firmas()

    # Funcion para el caso de ULTA BEAUTY ya que para la norma 024 es pegado de evidencia y pegado de etiquetas para las demas normas #
    def agregar_hoja_mixta(self):
        """Mezcla en una sola hoja evidencia y etiquetas."""
        evidencias = self.datos.get('evidencias_lista', []) or []
        etiquetas = self.datos.get('etiquetas_lista', []) or []

        self.elements.append(PageBreak())
        self.elements.append(Paragraph("<b>EVIDENCIA Y ETIQUETAS</b>", self.normal_style))
        self.elements.append(Spacer(1, 0.25 * inch))

        # --- Mostrar evidencia ---
        if evidencias:
            for ev in evidencias:
                img_bytes = ev.get('imagen_bytes')
                if img_bytes:
                    img_bytes.seek(0)
                    img = RLImage(img_bytes, width=4.5*inch, height=4.5*inch)
                    self.elements.append(img)
                    self.elements.append(Spacer(1, 0.25 * inch))

        # --- Mostrar etiquetas a un tamaño menor ---
        if etiquetas:
            for etq in etiquetas:
                img_bytes = etq.get('imagen_bytes')
                w_cm, h_cm = etq.get("tamaño_cm", (5,5))
                if img_bytes:
                    img_bytes.seek(0)
                    img = RLImage(img_bytes, width=w_cm*0.393701*inch/1.4,
                                            height=h_cm*0.393701*inch/1.4)
                    self.elements.append(img)
                    self.elements.append(Spacer(1, 0.15 * inch))

    def agregar_hoja_firmas(self):
        """Agrega una hoja con las firmas al final (PÁGINA SEPARADA)."""
        print("   🖊 Agregando hoja de firmas (PÁGINA SEPARADA)")
        self.elements.append(PageBreak())
        for e in self._get_firmas_elements():
            self.elements.append(e)

    def _get_firmas_elements(self):
        """Devuelve la lista de flowables que representan las firmas (sin PageBreak)."""
        elems = []
        bold_style = ParagraphStyle('BoldCenter', parent=self.normal_style, fontName='Helvetica-Bold', alignment=1)

        ruta_firma1 = self.datos.get('imagen_firma1', '')
        ruta_firma2 = self.datos.get('imagen_firma2', '')
        imagen_firma1 = obtener_ruta_recurso(ruta_firma1) if ruta_firma1 else None
        imagen_firma2 = obtener_ruta_recurso(ruta_firma2) if ruta_firma2 else None

        col1 = []
        if imagen_firma1 and os.path.exists(imagen_firma1):
            img1 = RLImage(imagen_firma1, width=2.2*inch, height=0.9*inch)
            col1.append(img1)
        col1.append(Paragraph("_______________________________", self.normal_style))
        col1.append(Paragraph(self.datos.get("nfirma1",""), bold_style))
        col1.append(Paragraph("Inspector", bold_style))

        col3 = []
        if imagen_firma2 and os.path.exists(imagen_firma2):
            img2 = RLImage(imagen_firma2, width=2.2*inch, height=0.9*inch)
            col3.append(img2)
        col3.append(Paragraph("_______________________________", self.normal_style))
        col3.append(Paragraph(self.datos.get("nfirma2",""), bold_style))
        col3.append(Paragraph("Responsable de Supervisión UI", bold_style))

        firmas_table = Table([[col1, "", col3]], colWidths=[2.5*inch, 0.5*inch, 2.5*inch])
        firmas_table.setStyle(TableStyle([
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('VALIGN',(0,0),(-1,-1),'TOP'),
        ]))

        elems.append(Spacer(1, 1 * inch))
        elems.append(firmas_table)
        return elems

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
        
        # Formato folio a 6 dígitos: ${year}049UDC${norma}${folio} Solicitud de Servicio: ${year}049USD${norma}${solicitud}-${lista}
        folio_formateado = f"{int(folio) if folio.isdigit() else 0:06d}"
        solicitud_formateado = f"{int(solicitud) if solicitud.isdigit() else 0:06d}"
        linea_completa = f"{year}049UDC{norma}{folio_formateado}   Solicitud de Servicio: {year}049USD{norma}{solicitud_formateado}-{lista}"
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

def convertir_dictamen_a_json(datos):
    """
    Convierte los datos del dictamen a formato JSON serializable.
    Extrae solo los datos relevantes, excluyendo objetos binarios como imágenes.
    """
    # Construir cadena_identificacion asegurando folio y solicitud a 6 dígitos
    year = str(datos.get("year", "")).strip()
    norma = str(datos.get("norma", "")).strip()
    folio_raw = str(datos.get("folio", "")).strip()
    solicitud_raw = str(datos.get("solicitud", "")).strip()
    lista = str(datos.get("lista", "")).strip()

    if folio_raw.isdigit():
        folio_formateado = f"{int(folio_raw):06d}"
    else:
        folio_formateado = folio_raw

    if solicitud_raw.isdigit():
        solicitud_formateado = f"{int(solicitud_raw):06d}"
    else:
        solicitud_formateado = solicitud_raw

    cadena_identificacion = (
        f"{year}049UDC{norma}{folio_formateado} Solicitud de Servicio: {year}049USD{norma}{solicitud_formateado}-{lista}"
    )

    json_data = {
        "identificacion": {
            "cadena_identificacion": cadena_identificacion,
            "year": datos.get("year", ""),
            "folio": datos.get("folio", ""),
            "solicitud": datos.get("solicitud", ""),
            "lista": datos.get("lista", "")
        },
        "norma": {
            "codigo": datos.get("norma", ""),
            "descripcion": datos.get("normades", ""),
            "capitulo": datos.get("capitulo", "")
        },
        "fechas": {
            "verificacion": datos.get("fverificacion", ""),
            "verificacion_larga": datos.get("fverificacionlarga", ""),
            "emision": datos.get("femision", "")
        },
        "cliente": {
            "nombre": datos.get("cliente", ""),
            "rfc": datos.get("rfc", "")
        },
        "producto": {
            "descripcion": datos.get("producto", ""),
            "pedimento": datos.get("pedimento", "")
        },
        "tabla_productos": datos.get("tabla_productos", []),
        "cantidad_total": {
            "valor": datos.get("total_cantidad", 0),
            "texto": datos.get("TCantidad", "")
        },
        "observaciones": datos.get("obs", ""),
        "firmas": {
            "firma1": {
                "nombre": datos.get("nfirma1", ""),
                "valida": datos.get("firma_valida", False),
                "codigo_solicitado": datos.get("codigo_firma_solicitado", ""),
                "razon_sin_firma": datos.get("razon_sin_firma", "")
            },
            "firma2": {
                "nombre": datos.get("nfirma2", "")
            }
        },
        "modo_insertado": datos.get("modo_insertado", "etiqueta"),
        "etiquetas": {
            "cantidad": len(datos.get("etiquetas_lista", []))
        }
    }
    return json_data

def guardar_dictamen_json(datos, lista, directorio_json):
    """
    Guarda los datos del dictamen en formato JSON.
    Retorna (exito, mensaje_error)
    """
    try:
        # Crear directorio si no existe
        os.makedirs(directorio_json, exist_ok=True)
        
        # Convertir datos a JSON
        json_data = convertir_dictamen_a_json(datos)
        
        # Nombre base del archivo JSON (limpiar caracteres no válidos)
        base_nombre = limpiar_nombre_archivo(f"Dictamen_Lista_{lista}.json")
        ruta_json = os.path.join(directorio_json, base_nombre)

        # Si ya existe, añadir timestamp para preservar archivos anteriores
        if os.path.exists(ruta_json):
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            nombre_archivo = limpiar_nombre_archivo(f"Dictamen_Lista_{lista}_{ts}.json")
            ruta_json = os.path.join(directorio_json, nombre_archivo)

        # Guardar archivo JSON
        with open(ruta_json, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        
        return True, None
    except Exception as e:
        return False, str(e)

def detectar_flujo_cliente(cliente_nombre, norma_nombre=""):
    """
    Detecta automáticamente qué flujo debe usar el cliente.
    Retorna: 'evidencia', 'etiqueta', 'mixto', o 'etiqueta' (default)
    """
    cliente_upper = str(cliente_nombre).upper().strip()
    norma_upper = str(norma_nombre).upper().strip()
    
    # ─────────────────────────────────────────────
    # CLIENTES QUE PEGAN EVIDENCIA
    # ─────────────────────────────────────────────
    CLIENTES_EVIDENCIA = {
        "BASECO SAPI DE CV",
        "BLUE STRIPES SA DE CV",
        "GRUPO GUESS S DE RL DE CV",
        "EAST COAST MODA SA DE CV",
        "I NOSTRI FRATELLI S DE RL DE CV",
        "LEDERY MEXICO SA DE CV",
        "MODA RAPSODIA SA DE CV",
        "MULTIBRAND OUTLET STORES SAPI DE CV",
        "RED STRIPES SA DE CV",
        "ROBERT BOSCH S DE RL DE CV",
        "UNILEVER MANUFACTURERA S DE RL DE CV",
        "UNILEVER DE MÉXICO S DE RL DE CV",
    }
    
    # ─────────────────────────────────────────────
    # CLIENTES QUE PEGAN ETIQUETAS
    # ─────────────────────────────────────────────
    CLIENTES_ETIQUETA = {
        "ARTICULOS DEPORTIVOS DECATHLON SA DE CV",
        "FERRAGAMO MEXICO S DE RL DE CV",
    }
    
    # ─────────────────────────────────────────────
    # ULTA BEAUTY: MIXTO PARA NOM-024, ETIQUETA PARA OTRAS
    # ─────────────────────────────────────────────
    if "ULTA BEAUTY" in cliente_upper:
        if "NOM-024" in norma_upper:
            return "mixto"
        else:
            return "etiqueta"
    
    if cliente_upper in CLIENTES_EVIDENCIA:
        return "evidencia"
    
    if cliente_upper in CLIENTES_ETIQUETA:
        return "etiqueta"
    
    # Default: etiqueta
    return "etiqueta"

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

    # Construir índice global de evidencias a partir de rutas guardadas por la UI
    evidencia_cfg = {}
    try:
        ruta_evidence_cfg = os.path.join(os.path.dirname(__file__), 'data', 'evidence_paths.json')
        if os.path.exists(ruta_evidence_cfg):
            with open(ruta_evidence_cfg, 'r', encoding='utf-8') as f:
                evidencia_cfg = json.load(f) or {}
    except Exception:
        evidencia_cfg = {}

    IMG_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}
    import re
    def _normalizar(s):
        return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

    def _construir_indice_de_carpetas(cfg):
        """Construye un índice dict: base_norm -> [paths]
        Evita duplicados y captura errores de acceso a carpetas.
        """
        index = {}
        total = 0
        for grp, lst in (cfg or {}).items():
            for carpeta in lst:
                try:
                    for root, _, files in os.walk(carpeta):
                        for nombre in files:
                            base, ext = os.path.splitext(nombre)
                            if ext.lower() not in IMG_EXTS:
                                continue
                            path = os.path.join(root, nombre)
                            key = _normalizar(base)
                            if not key:
                                continue
                            index.setdefault(key, []).append(path)
                            total += 1
                except Exception:
                    continue
        return index, total

    indice_evidencias_global, total_indexadas = _construir_indice_de_carpetas(evidencia_cfg)
    try:
        claves = list(indice_evidencias_global.keys())[:20]
    except Exception:
        claves = []
    print(f"🔎 Índice global de evidencias construido: {total_indexadas} imágenes indexadas, {len(indice_evidencias_global)} claves únicas. Muestras: {claves}")

    os.makedirs(directorio_destino, exist_ok=True)
    
    # Crear directorio para JSON dentro de 'data/Dictamenes' para centralizar los dictámenes
    directorio_json = os.path.join(os.path.dirname(__file__), 'data', 'Dictamenes')
    os.makedirs(directorio_json, exist_ok=True)
    
    dictamenes_generados = 0
    dictamenes_con_firma = 0
    dictamenes_sin_firma = 0
    dictamenes_error = 0
    
    json_generados = 0
    json_errores = 0
    json_errores_detalle = []
    
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

            # 🎯 DETECTAR Y ASIGNAR FLUJO AUTOMÁTICAMENTE
            # --- Intentar asignar evidencias a partir del índice global ---
            try:
                # Construir lista de códigos a buscar a partir de los registros (campo CODIGO)
                etiquetas = datos.get('etiquetas_lista', []) or []
                codigos_a_buscar = []
                try:
                    for r in registros:
                        c = r.get('CODIGO') or r.get('codigo') or r.get('EAN') or r.get('ean')
                        if c and str(c).strip() not in ("", "None", "nan"):
                            codigos_a_buscar.append(str(c).strip())
                except Exception:
                    codigos_a_buscar = []

                def _buscar_imagen(code):
                    key = _normalizar(code)
                    if not key:
                        return None
                    # 1) coincidencia exacta en claves del índice
                    if key in indice_evidencias_global:
                        lst = indice_evidencias_global.get(key) or []
                        if lst:
                            return lst[0]

                    # 2) coincidencias parciales en las claves (key dentro de clave o viceversa)
                    candidatos = [k for k in indice_evidencias_global.keys() if key in k or k in key]
                    if candidatos:
                        # ordenar por diferencia de longitud (más cercano) y lexicográficamente
                        candidatos.sort(key=lambda k: (abs(len(k) - len(key)), k))
                        primera = candidatos[0]
                        lst = indice_evidencias_global.get(primera) or []
                        if lst:
                            return lst[0]

                    # 3) fallback: buscar claves que contienen la mayor secuencia numérica del código
                    import re as _re
                    dig = ''.join(_re.findall(r"\d+", key))
                    if dig:
                        candidatos2 = [k for k in indice_evidencias_global.keys() if dig in k]
                        if candidatos2:
                            candidatos2.sort(key=lambda k: (abs(len(k) - len(dig)), k))
                            lst = indice_evidencias_global.get(candidatos2[0]) or []
                            if lst:
                                return lst[0]

                    # No encontrado
                    return None

                rutas_encontradas = []
                if codigos_a_buscar:
                    print(f"   🔎 Buscando evidencias para códigos: {codigos_a_buscar}")
                    for codigo in codigos_a_buscar:
                        p = _buscar_imagen(codigo)
                        print(f"      → {codigo} => {p}")
                        if p:
                            rutas_encontradas.append(p)
                            # Si etiquetas son dicts, anexar la ruta a la etiqueta correspondiente
                            if etiquetas and isinstance(etiquetas[0], dict):
                                for e in etiquetas:
                                    if str(e.get('codigo')) == str(codigo) or str(e.get('ean')) == str(codigo):
                                        e['imagen_path'] = p

                if rutas_encontradas:
                    datos['evidencias_lista'] = rutas_encontradas
                    print(f"   ✅ Evidencias asignadas: {rutas_encontradas}")
            except Exception:
                pass

            cliente = datos.get('cliente', 'DESCONOCIDO')
            cliente = datos.get('cliente', 'DESCONOCIDO')
            norma = datos.get('norma', '')
            flujo_detectado = detectar_flujo_cliente(cliente, norma)
            datos['modo_insertado'] = flujo_detectado
            print(f"   📌 Flujo detectado: {flujo_detectado.upper()} (Cliente: {cliente})")
            
            tiene_firma = datos.get("firma_valida", False)
            
            # 🎯 CREAR CARPETA POR SOLICITUD (SOL{solicitud})
            solicitud = str(datos.get('solicitud', '000000')).strip()
            solicitud_formateado = f"{int(solicitud) if solicitud.isdigit() else 0:06d}"
            carpeta_solicitud = os.path.join(directorio_destino, f"SOL {solicitud_formateado}")
            os.makedirs(carpeta_solicitud, exist_ok=True)
            
            generador = PDFGeneratorConDatos(datos)
            nombre_archivo = limpiar_nombre_archivo(f"Dictamen_Lista_{lista}.pdf")
            ruta_completa = os.path.join(carpeta_solicitud, nombre_archivo)

            if generador.generar_pdf_con_datos(ruta_completa):
                dictamenes_generados += 1
                archivos_creados.append(ruta_completa)
                
                # Guardar JSON del dictamen
                exito_json, error_json = guardar_dictamen_json(datos, lista, directorio_json)
                if exito_json:
                    json_generados += 1
                    print(f"   💾 JSON guardado: Dictamen_Lista_{lista}.json")
                else:
                    json_errores += 1
                    json_errores_detalle.append({
                        "lista": lista,
                        "error": error_json
                    })
                    print(f"   ⚠️ Error guardando JSON: {error_json}")
                
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
    
    print("\n" + "="*60)
    print("📄 RESUMEN DE ARCHIVOS JSON")
    print("="*60)
    print(f"✅ JSON generados: {json_generados}/{dictamenes_generados}")
    try:
        ruta_relativa = os.path.relpath(directorio_json)
    except:
        ruta_relativa = "data/dictamenes/"
    print(f"📂 Ubicación: {ruta_relativa} (carpeta interna del proyecto)")
    if json_errores > 0:
        print(f"❌ Errores JSON: {json_errores}")
    
    if json_errores_detalle:
        print("\n" + "="*60)
        print("⚠️  ERRORES AL GUARDAR JSON - DETALLE")
        print("="*60)
        for item in json_errores_detalle:
            print(f"\n📄 Lista: {item['lista']}")
            print(f"   Error: {item['error']}")
    
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
        'sin_firma_detalle': sin_firma_detalle,
        'json_errores': json_errores,
        'json_errores_detalle': json_errores_detalle
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
    else:
        print(f"\n❌ {mensaje}")
