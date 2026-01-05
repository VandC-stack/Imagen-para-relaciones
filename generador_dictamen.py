"""Generador de Dictámenes PDF con Datos Reales e Imágenes de Etiquetas"""
import os
import sys
import json
import pandas as pd
from datetime import datetime
import traceback
import time

# Evitar UnicodeEncodeError en consolas Windows (CP1252) al imprimir emojis u
# otros caracteres Unicode. Intentar reconfigurar stdout/stderr a UTF-8 cuando
# sea posible.
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from plantillaPDF import (
    cargar_tabla_relacion,
    cargar_normas,
    cargar_clientes,
    cargar_firmas,
    procesar_familias,
    preparar_datos_familia
)

from DictamenPDF import PDFGenerator
import folio_manager

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as RLImage, PageBreak, KeepTogether
)
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader

def obtener_ruta_recurso(ruta_relativa):
    """
    Obtiene la ruta absoluta del recurso, funciona tanto para .py como para .exe.
    PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS.
    """
    # Preferir carpeta junto al ejecutable (portable `data`), luego _MEIPASS, luego cwd
    try:
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            candidate = os.path.join(exe_dir, ruta_relativa)
            # si existe la ruta o su carpeta padre existe, usarla
            parent = os.path.dirname(candidate)
            try:
                if os.path.exists(candidate) or (parent and os.path.exists(parent)):
                    return candidate
            except Exception:
                pass
    except Exception:
        pass

    try:
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            return os.path.join(meipass, ruta_relativa)
    except Exception:
        pass

    return os.path.join(os.path.abspath('.'), ruta_relativa)


# ---------------- Folio counter (reserva atómica) ----------------
def _get_folio_paths():
    carpeta = obtener_ruta_recurso('data')
    try:
        os.makedirs(carpeta, exist_ok=True)
    except Exception:
        pass
    return os.path.join(carpeta, 'folio_counter.json'), os.path.join(carpeta, 'folio_counter.lock')

def _acquire_lock(lock_path, timeout=5.0):
    start = time.time()
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return True
        except FileExistsError:
            if (time.time() - start) >= timeout:
                return False
            time.sleep(0.1)

def _release_lock(lock_path):
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass

def reservar_siguiente_folio(timeout=5.0):
    """Reserva el siguiente folio delegando al nuevo módulo `folio_manager`."""
    try:
        return folio_manager.reserve_next(timeout=timeout)
    except Exception as e:
        raise RuntimeError(f"No se pudo reservar siguiente folio: {e}")
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
        """Hoja en blanco para evidencia + hoja de firmas.

        Distribuye hasta 5 evidencias por página en una cuadrícula (2 columnas).
        Si la última página de evidencias tiene menos de 5 imágenes, las firmas
        se colocan en esa misma página en lugar de crear una página separada.
        """
        print("   📄 Generando hoja para evidencias (máx. 5 por hoja)...")

        # HOJA – iniciar sección de evidencias
        self.elements.append(PageBreak())
        self.elements.append(Spacer(1, 0.25 * inch))

        evidencias = self.datos.get('evidencias_lista', []) or []

        if not evidencias:
            # placeholder cuando no hay evidencias
            self.elements.append(Paragraph(
                "<b>${IMAGEN}</b>",
                ParagraphStyle('Center', parent=self.normal_style, alignment=1, fontSize=12)
            ))
            # luego agregar la página de firmas normalmente
            self.agregar_hoja_firmas()
            return

        # Procesar evidencias y normalizarlas a flowables RLImage
        from io import BytesIO
        from PIL import Image as PILImage
        import traceback

        image_flowables = []
        for idx, ev in enumerate(evidencias, start=1):
            try:
                bio = None
                # cadena = ruta en disco
                if isinstance(ev, str):
                    ruta = os.path.normpath(ev)
                    if not os.path.exists(ruta):
                        # intentar variantes simples
                        ruta_alt = ruta.replace('\\', '/')
                        if os.path.exists(ruta_alt):
                            ruta = ruta_alt
                        else:
                            ruta_alt2 = ruta.replace('/', '\\')
                            if os.path.exists(ruta_alt2):
                                ruta = ruta_alt2
                            else:
                                print(f"         ⚠️ Ruta no encontrada: {ruta} (omitida)")
                                continue
                    with PILImage.open(ruta) as im:
                        im.verify()
                    with PILImage.open(ruta) as im2:
                        if im2.mode != 'RGB':
                            im2 = im2.convert('RGB')
                        bio = BytesIO()
                        im2.save(bio, format='JPEG', quality=90, optimize=True)
                        bio.seek(0)

                elif isinstance(ev, dict):
                    img_bytes = ev.get('imagen_bytes') or ev.get('imagen_path_bytes')
                    if img_bytes:
                        bio_in = img_bytes if hasattr(img_bytes, 'read') else BytesIO(img_bytes)
                        try:
                            with PILImage.open(bio_in) as im:
                                im.verify()
                            bio_in.seek(0)
                            with PILImage.open(bio_in) as im2:
                                if im2.mode != 'RGB':
                                    im2 = im2.convert('RGB')
                                bio = BytesIO()
                                im2.save(bio, format='JPEG', quality=90, optimize=True)
                                bio.seek(0)
                        except Exception:
                            traceback.print_exc()
                            continue
                    else:
                        p = ev.get('imagen_path')
                        if p and os.path.exists(p):
                            with PILImage.open(p) as im:
                                im.verify()
                            with PILImage.open(p) as im2:
                                if im2.mode != 'RGB':
                                    im2 = im2.convert('RGB')
                                bio = BytesIO()
                                im2.save(bio, format='JPEG', quality=90, optimize=True)
                                bio.seek(0)
                        else:
                            print(f"         ⚠️ imagen_path no existe o inválida: {p}")
                            continue

                else:
                    # file-like / BytesIO
                    if hasattr(ev, 'seek'):
                        try:
                            ev.seek(0)
                        except Exception:
                            pass
                    bio = ev

                if not bio:
                    continue

                # Crear RLImage con tamaño adecuado para 2 columnas x 2 filas (4 por página)
                try:
                    img = RLImage(bio, width=3.4*inch, height=3.0*inch)
                except Exception:
                    # reintentar creando desde BytesIO copy
                    try:
                        tmp = BytesIO(bio.read() if hasattr(bio, 'read') else bio)
                        tmp.seek(0)
                        img = RLImage(tmp, width=3.8*inch, height=3.0*inch)
                    except Exception:
                        traceback.print_exc()
                        continue

                image_flowables.append(img)
            except Exception:
                traceback.print_exc()
                continue

        # Dividir en páginas, 4 imágenes por página (2x2)
        pages = [image_flowables[i:i+4] for i in range(0, len(image_flowables), 4)]

        for p_idx, chunk in enumerate(pages):
            # Construir filas para una tabla de 2 columnas (2 filas para hasta 4 imágenes)
            rows = []
            imgs = chunk
            if len(imgs) == 4:
                rows.append([imgs[0], imgs[1]])
                rows.append([imgs[2], imgs[3]])
            elif len(imgs) == 3:
                rows.append([imgs[0], imgs[1]])
                rows.append([imgs[2], ''])
            elif len(imgs) == 2:
                rows.append([imgs[0], imgs[1]])
            elif len(imgs) == 1:
                rows.append([imgs[0], ''])

            # Crear tabla y añadir a elementos
            col_width = (8.5*inch - 2*0.75*inch) / 2  # aproximación con márgenes
            tbl = Table(rows, colWidths=[col_width, col_width])
            tbl.setStyle(TableStyle([
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('LEFTPADDING',(0,0),(-1,-1),10),
                ('RIGHTPADDING',(0,0),(-1,-1),10),
                ('TOPPADDING',(0,0),(-1,-1),12),
                ('BOTTOMPADDING',(0,0),(-1,-1),12),
            ]))

            self.elements.append(tbl)
            # aumentar espacio entre bloques de evidencias
            self.elements.append(Spacer(1, 0.35*inch))

            # Si es la última página y tiene menos de 5 imágenes, colocar firmas aquí
            is_last = (p_idx == len(pages) - 1)
            if is_last and len(chunk) < 5:
                print("   📌 Firmas mostradas ABAJO de evidencias (última página incompleta)")
                self.elements.append(Spacer(1, 0.15 * inch))
                for e in self._get_firmas_elements():
                    self.elements.append(e)
                return
            else:
                # Si no es la última página, o la última tiene 5 imágenes, añadir salto de página
                if not is_last:
                    self.elements.append(PageBreak())

        # Si llegamos aquí y la última página fue completa (5 imágenes), añadir página de firmas
        print("   📌 Firmas mostradas en PÁGINA SEPARADA (última página completa)")
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
                    try:
                        img = RLImage(img_bytes, width=4.5*inch, height=4.5*inch)
                    except Exception:
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
        
        # Preferir el `year` incluido en los datos del dictamen si existe (dos dígitos),
        # en caso contrario usar el año actual.
        year = str(self.datos.get('year', '')).strip()
        if not year:
            year = datetime.now().strftime("%y")
        # Si viene como 4 dígitos, usar los últimos dos
        if year and year.isdigit() and len(year) == 4:
            year = year[-2:]

        norma = str(self.datos.get('norma', '')).strip()
        folio = str(self.datos.get('folio', '')).strip()
        solicitud = str(self.datos.get('solicitud', '')).strip()
        lista = str(self.datos.get('lista', '')).strip()

        # Normalizar solicitud: si viene con '/', tomar parte antes del '/'
        if '/' in solicitud:
            solicitud = solicitud.split('/')[0].strip()

        # Formato folio y solicitud a 6 dígitos cuando son numéricos
        folio_formateado = folio.zfill(6) if folio.isdigit() else folio
        solicitud_formateado = solicitud.zfill(6) if solicitud.isdigit() else solicitud
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
    norma = str(datos.get("norma", "")).strip()
    folio_raw = str(datos.get("folio", "")).strip()
    solicitud_raw = str(datos.get("solicitud", "")).strip()
    lista = str(datos.get("lista", "")).strip()

    # Extraer año desde la solicitud si está presente (p. ej. "006669/25").
    # Preferimos el año indicado en la solicitud por encima del campo 'year'
    year_from_solicitud = ''
    try:
        if '/' in solicitud_raw:
            parts = solicitud_raw.split('/')
            suf = parts[-1].strip()
            if suf.isdigit():
                # Tomar los dos últimos dígitos (p. ej. 2025 -> 25)
                year_from_solicitud = suf[-2:]
    except Exception:
        year_from_solicitud = ''

    # Determinar year definitivo: preferir el extraído desde la solicitud
    if year_from_solicitud:
        year = year_from_solicitud
    else:
        year = str(datos.get("year", "")).strip()

    # Si viene como 4 dígitos (ej. 2025), usar los últimos dos
    if year and year.isdigit() and len(year) == 4:
        year = year[-2:]

    # Formatear folio a 6 dígitos si tiene dígitos
    folio_digits = ''.join([c for c in folio_raw if c.isdigit()])
    folio_formateado = folio_digits.zfill(6) if folio_digits else folio_raw

    # Formatear solicitud: tomar la parte antes de '/' o los dígitos
    solicitud_num = ''
    if solicitud_raw:
        if '/' in solicitud_raw:
            solicitud_num = solicitud_raw.split('/')[0].strip()
        else:
            solicitud_num = ''.join([c for c in solicitud_raw if c.isdigit()])
    solicitud_formateado = solicitud_num.zfill(6) if solicitud_num and solicitud_num.isdigit() else solicitud_num

    # Construir cadena_identificacion siempre (asegurar variable definida)
    cadena_identificacion = (
        f"{year}049UDC{norma}{folio_formateado} Solicitud de Servicio: {year}049USD{norma}{solicitud_formateado}-{lista}"
    )

    json_data = {
        "identificacion": {
            "cadena_identificacion": cadena_identificacion,
            "year": year,
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

def guardar_dictamen_json(datos, lista, directorio_json, metadata=None):
    """
    Guarda los datos del dictamen en formato JSON. Añade un campo opcional
    'metadata' dentro del JSON para almacenar estado del PDF u otros datos
    de diagnóstico. Retorna (exito, mensaje_error)
    """
    try:       # Crear directorio si no existe
        os.makedirs(directorio_json, exist_ok=True)

        # Convertir datos a JSON base
        json_data = convertir_dictamen_a_json(datos)

        # Añadir metadata si se proporcionó (no sobrescribe campos existentes)
        if metadata and isinstance(metadata, dict):
            try:
                json_data.setdefault('metadata', {})
                for k, v in metadata.items():
                    json_data['metadata'][k] = v
            except Exception:
                pass

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
    # CLIENTES QUE PEGAN ETIQUETAS (EXCEPCIONES)
    # ─────────────────────────────────────────────
    # Todos los clientes se tratan como flujo de EVIDENCIA por defecto,
    # salvo los listados aquí. Añadimos la regla especial de ULTA.
    CLIENTES_ETIQUETA = {
        "ARTICULOS DEPORTIVOS DECATHLON SA DE CV",
        "FERRAGAMO MEXICO S DE RL DE CV",
        "ULTA BEAUTY SAPI DE CV",
    }
    
    # ─────────────────────────────────────────────
    # ULTA BEAUTY: MIXTO PARA NOM-024, ETIQUETA PARA OTRAS
    # ─────────────────────────────────────────────
    if "ULTA BEAUTY" in cliente_upper:
        if "NOM-024" in norma_upper:
            return "mixto"
        else:
            return "etiqueta"

    # Si está en la lista de etiquetas -> modo etiqueta
    if cliente_upper in CLIENTES_ETIQUETA:
        return "etiqueta"

    # Por defecto todos los demás clientes usan modo evidencia
    return "evidencia"

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
        ruta_evidence_cfg = obtener_ruta_recurso('data/evidence_paths.json')
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
    directorio_json = obtener_ruta_recurso('data/Dictamenes')
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
    folios_usados_set = set()

    # Calcular bloque de folios a asignar para este proceso.
    total_needed = len(familias)
    last_known = None
    try:
        last_known = folio_manager.get_last()
    except Exception:
        last_known = None

    # Fallback adicional: leer directamente `data/folio_counter.json` usando
    # `obtener_ruta_recurso` (maneja rutas para .py y .exe). Esto ayuda cuando
    # folio_manager pueda estar apuntando a una ruta diferente inesperadamente
    # o cuando fue editado manualmente fuera del flujo del módulo.
    try:
        if not last_known or int(last_known) == 0:
            contador_path = os.path.join(obtener_ruta_recurso('data'), 'folio_counter.json')
            if os.path.exists(contador_path):
                try:
                    with open(contador_path, 'r', encoding='utf-8') as cf:
                        j = json.load(cf) or {}
                        v = int(j.get('last', 0))
                        if v and v > 0:
                            last_known = v
                            print(f"   ℹ️ folio_counter.json (fallback) leído: {last_known}")
                            # Intentar sincronizar con folio_manager para asegurar que
                            # las próximas reservas usen este valor como referencia.
                            try:
                                current_mgr = folio_manager.get_last()
                                if current_mgr is None or int(current_mgr) < int(v):
                                    try:
                                        folio_manager.set_last(int(v))
                                        print(f"   ℹ️ folio_manager sincronizado a {int(v)}")
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass

    # Detectar si la tabla ya trae folios asignados por familia. Si es así,
    # respetamos esos folios y evitamos reservar de nuevo (para no duplicar
    # el avance del contador). Si no hay folios preasignados, intentamos reservar
    # un bloque atómico aquí.
    preassigned_map = {}
    try:
        assigned_set = set()
        for lista, registros in familias.items():
            found = None
            for rec in registros:
                # considerar 'FOLIO' o 'folio'
                val = rec.get('FOLIO') if 'FOLIO' in rec else rec.get('folio')
                if val is not None and str(val).strip() != "":
                    try:
                        n = int(float(str(val)))
                        found = int(n)
                        break
                    except Exception:
                        continue
            if found is not None:
                preassigned_map[lista] = found
                assigned_set.add(found)
    except Exception:
        preassigned_map = {}

    use_preassigned = False
    try:
        # Only treat table folios as preassigned if every folio is a valid positive int
        # and is greater than the current persisted last known folio. This avoids
        # reusing placeholder or old folios present in the table (e.g., 000001..)
        current_last = int(last_known) if last_known is not None else 0
        if total_needed > 0 and len(preassigned_map) == total_needed and len(assigned_set) == total_needed:
            try:
                min_assigned = min(int(x) for x in assigned_set)
                if min_assigned > current_last:
                    use_preassigned = True
                else:
                    use_preassigned = False
            except Exception:
                use_preassigned = False
    except Exception:
        use_preassigned = False

    next_folio_to_assign = None
    reserved_here = False

    if use_preassigned:
        print(f"🔎 Se detectaron folios preasignados en la tabla; se usarán sin reservar aquí.")
    else:
        if total_needed > 0:
            try:
                start = folio_manager.reserve_block(total_needed)
                next_folio_to_assign = start
                reserved_here = True
            except Exception:
                # Si reserve_block falla, pero conocemos el último contador,
                # asignamos a partir de last_known+1 y tratamos de persistir
                # el nuevo último con set_last().
                if last_known is not None:
                    try:
                        start = int(last_known) + 1
                        next_folio_to_assign = start
                        try:
                            folio_manager.set_last(int(last_known) + int(total_needed))
                        except Exception:
                            pass
                    except Exception:
                        next_folio_to_assign = None
                else:
                    # Fallback robusto: inspeccionar archivos locales para calcular el siguiente
                    try:
                        carpeta = os.path.dirname(_get_folio_paths()[0])
                        dirp = os.path.join(carpeta, 'folios_visitas')
                        maxf = 0
                        if os.path.exists(dirp):
                            import re
                            for fn in os.listdir(dirp):
                                if fn.startswith('folios_') and fn.endswith('.json'):
                                    pathf = os.path.join(dirp, fn)
                                    try:
                                        with open(pathf, 'r', encoding='utf-8') as fh:
                                            arr = json.load(fh) or []
                                            for entry in arr:
                                                fol = entry.get('FOLIOS') or ''
                                                nums = re.findall(r"\d+", str(fol))
                                                for d in nums:
                                                    try:
                                                        n = int(d)
                                                        if n > maxf:
                                                            maxf = n
                                                    except Exception:
                                                        pass
                                    except Exception:
                                        continue
                        next_folio_to_assign = maxf + 1
                    except Exception:
                        next_folio_to_assign = None

        else:
            next_folio_to_assign = None

        # Asegurar coherencia: si conocemos el último folio persistido, el siguiente
        # folio a asignar no debe ser menor que last_known + 1.
        try:
            if last_known is not None and not use_preassigned:
                minimo = int(last_known) + 1
                if next_folio_to_assign is None:
                    print(f"   ℹ️ Usando folio_counter como inicio: {minimo}")
                    next_folio_to_assign = minimo
                    try:
                        folio_manager.set_last(int(last_known) + int(total_needed))
                    except Exception:
                        pass
                else:
                    if int(next_folio_to_assign) < minimo:
                        print(f"   ⚠️ Ajustando inicio de folios {next_folio_to_assign} -> {minimo} (según folio_counter)")
                        next_folio_to_assign = minimo
        except Exception:
            pass

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

            # ---------------- Asignar folio automático por familia (LISTA) ----------------
            try:
                # Si pre-calculamos un bloque, usarlo y avanzar la variable local;
                # si no, usar el mecanismo de reserva atómica por compatibilidad.
                if use_preassigned:
                    # usar folio preasignado por la tabla (por lista)
                    folio_num = preassigned_map.get(lista)
                    if folio_num is None:
                        # fallback: reservar uno-a-uno
                        folio_num = reservar_siguiente_folio()
                else:
                    if next_folio_to_assign is None:
                        folio_num = reservar_siguiente_folio()
                    else:
                        folio_num = next_folio_to_assign
                        next_folio_to_assign += 1

                datos['folio'] = str(folio_num)
                print(f"   🔢 Folio asignado automáticamente: {int(folio_num):06d}")
                # Propagar el folio asignado a cada registro de la familia (columna 'FOLIO')
                try:
                    for rec in registros:
                        try:
                            rec['FOLIO'] = int(folio_num)
                        except Exception:
                            rec['FOLIO'] = str(folio_num)
                except Exception:
                    pass
            except Exception as e:
                print(f"   ⚠️ No se pudo reservar folio automáticamente: {e}")
                traceback.print_exc()
                # Intentar reserva uno-a-uno como fallback antes de usar folios preexistentes
                try:
                    folio_num = reservar_siguiente_folio()
                    datos['folio'] = str(folio_num)
                    print(f"   🔁 Reserva fallback exitosa: {int(folio_num):06d}")
                    try:
                        for rec in registros:
                            try:
                                rec['FOLIO'] = int(folio_num)
                            except Exception:
                                rec['FOLIO'] = str(folio_num)
                    except Exception:
                        pass
                except Exception as e2:
                    print(f"   ⚠️ Fallback de reserva uno-a-uno falló: {e2}")
                    # Último recurso: mantener folio existente en registros si lo hubiera
                    try:
                        posible = registros[0].get('FOLIO') or registros[0].get('folio')
                        if posible:
                            datos['folio'] = str(posible)
                            print(f"   ℹ️ Usando folio preexistente: {datos['folio']}")
                    except Exception:
                        pass

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

            pdf_ok = False
            try:
                pdf_ok = generador.generar_pdf_con_datos(ruta_completa)
            except Exception as e:
                pdf_ok = False
                pdf_error_msg = str(e)
            else:
                pdf_error_msg = None

            if pdf_ok:
                dictamenes_generados += 1
                archivos_creados.append(ruta_completa)
                try:
                    used_folio = int(str(datos.get('folio') or folio_num))
                    folios_usados_set.add(used_folio)
                except Exception:
                    pass

                # Guardar JSON del dictamen con metadata indicando PDF creado
                meta = {'pdf_generado': True, 'pdf_path': os.path.abspath(ruta_completa)}
                exito_json, error_json = guardar_dictamen_json(datos, lista, directorio_json, metadata=meta)
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
                # Incluso si el PDF falló, intentar guardar JSON con metadata de error
                try:
                    meta = {'pdf_generado': False, 'pdf_error': str(pdf_error_msg or 'Error desconocido')}
                    exito_json, error_json = guardar_dictamen_json(datos, lista, directorio_json, metadata=meta)
                    if exito_json:
                        json_generados += 1
                        print(f"   💾 JSON guardado (error): Dictamen_Lista_{lista}.json")
                    else:
                        json_errores += 1
                        json_errores_detalle.append({
                            "lista": lista,
                            "error": error_json
                        })
                        print(f"   ⚠️ Error guardando JSON tras fallo de PDF: {error_json}")
                except Exception:
                    pass

        except Exception as e:
            dictamenes_error += 1
            print(f"   ❌ Error en familia {lista}: {e}")
            traceback.print_exc()
            continue

    # Actualizar folio_counter.json al último folio asignado para este lote
    try:
        # Si usamos folios preasignados por la tabla, asumimos que ya se hizo
        # la reserva (o que la tabla fue preparada por el usuario) y NO
        # actualizamos el contador aquí para evitar duplicados.
        if reserved_here:
            # Si reservamos un bloque, ajustar el contador al máximo folio efectivamente usado.
            try:
                if folios_usados_set:
                    max_used = int(max(folios_usados_set))
                    try:
                        folio_manager.set_last(max_used)
                        print(f"   🔢 Reserva ajustada: folio_counter fijado a {int(max_used):06d} (máx. folio usado)")
                    except Exception:
                        # Fallback atómico
                        try:
                            counter_path = os.path.join(os.path.dirname(__file__), 'data', 'folio_counter.json')
                            tmp = counter_path + '.tmp'
                            with open(tmp, 'w', encoding='utf-8') as tf:
                                json.dump({'last': int(max_used)}, tf)
                            try:
                                os.replace(tmp, counter_path)
                            except Exception:
                                if os.path.exists(counter_path):
                                    os.remove(counter_path)
                                os.replace(tmp, counter_path)
                            print(f"   🔢 Reserva ajustada (fallback): folio_counter fijado a {int(max_used):06d}")
                        except Exception as e:
                            print(f"   ⚠️ No se pudo ajustar reserva del contador: {e}")
                else:
                    # No se usó ningún folio: revertir al valor conocido
                    if last_known is not None:
                        try:
                            folio_manager.set_last(int(last_known))
                            print(f"   🔁 Reserva revertida: folio_counter restaurado a {int(last_known):06d}")
                        except Exception:
                            try:
                                counter_path = os.path.join(os.path.dirname(__file__), 'data', 'folio_counter.json')
                                tmp = counter_path + '.tmp'
                                with open(tmp, 'w', encoding='utf-8') as tf:
                                    json.dump({'last': int(last_known)}, tf)
                                try:
                                    os.replace(tmp, counter_path)
                                except Exception:
                                    if os.path.exists(counter_path):
                                        os.remove(counter_path)
                                    os.replace(tmp, counter_path)
                                print(f"   🔁 Reserva revertida (fallback): folio_counter restaurado a {int(last_known):06d}")
                            except Exception as e:
                                print(f"   ⚠️ No se pudo revertir reserva del contador: {e}")
            except Exception:
                pass
        else:
            if (not use_preassigned) and next_folio_to_assign is not None and not reserved_here:
                try:
                    last_to_write = int(next_folio_to_assign) - 1
                    folio_manager.set_last(last_to_write)
                except Exception:
                    pass
    except Exception:
        pass

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

    # Preparar información de folios utilizados para feedback en UI
    folios_info = None
    folios_list = None
    try:
        if folios_usados_set:
            used_sorted = sorted(int(x) for x in folios_usados_set)
            folios_list = [f"{x:06d}" for x in used_sorted]
            if len(used_sorted) == 1:
                folios_info = f"{used_sorted[0]:06d}"
            else:
                folios_info = f"{used_sorted[0]:06d} - {used_sorted[-1]:06d}"
    except Exception:
        folios_info = None

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
        'json_errores_detalle': json_errores_detalle,
        'folios_utilizados': folios_info,
        'folios_usados_list': folios_list
    }

    # Exportar una copia plana de la tabla de relación con los folios actualizados
    try:
        tabla_out = []
        # `familias` es un dict lista->registros donde cada registro ya contiene 'FOLIO'
        for lista_k, regs in familias.items():
            for r in regs:
                try:
                    # Asegurar que FOLIO esté en formato int o string
                    fol = r.get('FOLIO') if 'FOLIO' in r else r.get('folio')
                except Exception:
                    fol = r.get('folio', '')
                # Normalizar algunos campos que `guardar_folios_visita` espera
                tabla_out.append({
                    'FOLIO': fol,
                    'MARCA': r.get('MARCA') or r.get('marca') or '',
                    'SOLICITUD': r.get('SOLICITUD') or r.get('SOLICITUDES') or r.get('solicitud') or '',
                    'FECHA DE VERIFICACION': r.get('FECHA DE VERIFICACION') or r.get('fverificacion') or '',
                    'TIPO DE DOCUMENTO': r.get('TIPO DE DOCUMENTO') or r.get('tipo_documento') or 'D',
                    'INSPECTOR': r.get('INSPECTOR') or r.get('inspector') or r.get('Inspector') or '',
                    'LISTA': lista_k,
                    'CODIGO': r.get('CODIGO') or r.get('codigo') or r.get('Codigo') or '',
                    # Añadir campos necesarios para que la app pueda mapear supervisor y norma
                    'FIRMA': r.get('FIRMA') or r.get('firma') or r.get('Firma') or '',
                    'CLASIF UVA': r.get('CLASIF UVA') or r.get('CLASIF_UVA') or r.get('CLASIF_Uva') or r.get('clasif_uva') or None,
                    'NORMA UVA': r.get('NORMA UVA') or r.get('NORMA_UVA') or r.get('NORMA_Uva') or r.get('norma_uva') or None
                })
        if tabla_out:
            ts = datetime.now().strftime('%Y%m%d%H%M%S')
            out_path = os.path.join(directorio_json, f"tabla_relacion_generada_{ts}.json")
            try:
                with open(out_path, 'w', encoding='utf-8') as of:
                    json.dump(tabla_out, of, ensure_ascii=False, indent=2)
                resultado['tabla_relacion_actualizada'] = out_path
                print(f"   💾 Tabla de relación actualizada guardada: {out_path}")
            except Exception:
                pass
    except Exception:
        pass

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
