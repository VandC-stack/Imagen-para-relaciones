# IMPRIME LAS ETIQUETAS DENTRO DEL DICTAMEN #  

import json
import os
import sys
from PIL import Image, ImageDraw, ImageFont
import textwrap
import ast
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from io import BytesIO
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
from reportlab.lib.utils import ImageReader

class GeneradorEtiquetasDecathlon:
    def __init__(self):
        # Detectar ruta de `data` en tres lugares (preferir carpeta junto al exe):
        # 1) carpeta junto al ejecutable (APP_DIR / exe dir)
        # 2) PyInstaller _MEIPASS (bundle interno)
        # 3) directorio actual
        exe_dir = None
        try:
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
        except Exception:
            exe_dir = None

        candidates = []
        if exe_dir:
            candidates.append(os.path.join(exe_dir, 'data'))
        # _MEIPASS (if running from a bundle)
        try:
            meipass = getattr(sys, '_MEIPASS', None)
            if meipass:
                candidates.append(os.path.join(meipass, 'data'))
        except Exception:
            pass

        candidates.append(os.path.join(os.path.abspath('.'), 'data'))

        # Choose the first existing 'data' folder, otherwise fallback to first candidate
        data_dir = None
        for c in candidates:
            try:
                if os.path.isdir(c):
                    data_dir = c
                    break
            except Exception:
                continue
        if data_dir is None:
            data_dir = candidates[0]

        self.data_dir = data_dir

        # Rutas completas (resolviendo posibles variantes de nombre de archivo)
        base_etiquetado_path = None
        for name in ("BASE_ETIQUETADO.json", "base_etiquetado.json", "Base_Etiquetado.json"):
            p = os.path.join(self.data_dir, name)
            if os.path.exists(p):
                base_etiquetado_path = p
                break

        tabla_relacion_path = None
        for name in ("TABLA_DE_RELACION.json", "tabla_de_relacion.json", "tabla_relacion.json"):
            p = os.path.join(self.data_dir, name)
            if os.path.exists(p):
                tabla_relacion_path = p
                break

        config_etiquetas_path = os.path.join(self.data_dir, "config_etiquetas.json")

        # Pasar None si no existen; la función `cargar_datos` manejará la ausencia
        self.cargar_datos(base_etiquetado_path, tabla_relacion_path)
        self.configuraciones = self.cargar_configuraciones(config_etiquetas_path)
        self.mapeo_norma_uva = self.crear_mapeo_norma_uva()

    def cargar_datos(self, base_etiquetado_path, tabla_relacion_path):
        """Carga los datos de la base de etiquetado y tabla de relación"""
        try:
            if base_etiquetado_path and os.path.exists(base_etiquetado_path):
                with open(base_etiquetado_path, 'r', encoding='utf-8') as f:
                    self.base_etiquetado = json.load(f)
            else:
                print(f"⚠️ BASE_ETIQUETADO no encontrado en {self.data_dir}")
                self.base_etiquetado = []

            if tabla_relacion_path and os.path.exists(tabla_relacion_path):
                with open(tabla_relacion_path, 'r', encoding='utf-8') as f:
                    self.tabla_relacion = json.load(f)
            else:
                print(f"⚠️ TABLA_DE_RELACION no encontrada en {self.data_dir}")
                self.tabla_relacion = []

            print("✅ Archivos JSON procesados (si existían)")
        except Exception as e:
            print(f"❌ Error cargando archivos: {e}")
            self.base_etiquetado = []
            self.tabla_relacion = []
    
    def insertar_etiquetas_en_dictamen(self, dictamen_path, etiquetas, output_pdf="DICTAMEN_FINAL.pdf"):
        try:
            print("🔧 Insertando etiquetas dentro del dictamen...")

            # 1) Leer PDF original
            reader = PdfReader(dictamen_path)
            writer = PdfWriter()

            # 2) Copiar todas las hojas del dictamen al writer
            for page in reader.pages:
                writer.add_page(page)

            # 3) Crear una página nueva para etiquetas
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=letter)

            x = 40
            y = 720

            for etiqueta in etiquetas:
                ancho_cm, alto_cm = etiqueta["tamaño_cm"]
                ancho_pt = ancho_cm * 28.35
                alto_pt = alto_cm * 28.35

                img_bytes = etiqueta["imagen_bytes"]

                # Insertar etiqueta en PDF temporal (usar ImageReader para BytesIO/objetos)
                try:
                    c.drawImage(ImageReader(img_bytes), x, y - alto_pt, width=ancho_pt, height=alto_pt)
                except Exception:
                    # Fallback: intentar usar la ruta si está disponible como str
                    try:
                        c.drawImage(img_bytes, x, y - alto_pt, width=ancho_pt, height=alto_pt)
                    except Exception:
                        # si falla, saltar esa etiqueta
                        continue

                x += ancho_pt + 20
                if x > 500:
                    x = 40
                    y -= alto_pt + 40

            c.save()

            packet.seek(0)
            nueva_pagina = PdfReader(packet).pages[0]

            # 4) Añadir la nueva página al final del PDF
            writer.add_page(nueva_pagina)

            # 5) Guardar PDF final
            with open(output_pdf, "wb") as f:
                writer.write(f)

            print(f"✅ Dictamen final generado: {output_pdf}")
            return True

        except Exception as e:
            print(f"❌ Error insertando etiquetas en dictamen: {e}")
            return False

    def cargar_configuraciones(self, config_etiquetas_path):
        """Carga las configuraciones desde un archivo JSON"""
        try:
            with open(config_etiquetas_path, 'r', encoding='utf-8') as f:
                configs = json.load(f)
            
            # Procesar tamaños (convertir de string a tupla)
            for norma, config in configs.items():
                tamaño_str = config.get('tamaño_cm', '(0,0)')
                # Convertir string a tupla
                try:
                    config['tamaño'] = ast.literal_eval(tamaño_str)
                except:
                    config['tamaño'] = (5.0, 5.0)  # Tamaño por defecto
                    
            print("✅ Configuraciones de etiquetas cargadas")
            return configs
        except Exception as e:
            print(f"❌ Error cargando configuraciones: {e}")
            return {}
    
    def crear_mapeo_norma_uva(self):
        """Crea el mapeo completo entre NORMA UVA y las configuraciones de etiquetas"""
        return {
            4: {
                "con_insumos": "NOM-004-SE-2021",
                "sin_insumos": "NOM-004-TEX"
            },
            15: "NOM-015-SCFI-2007",
            20: "NOM-020-SCFI-1997", 
            24: "NOM-024-SCFI-2013",
            50: "NOM-050-SCFI-2004-1",
            141: "NOM-141-SSA1 SCFI-2012",
            189: "NOM-189-SSA1/SCFI-2018",
            142: "NOM-142-SSA1/SCFI-2014",
            51: "NOM-051-SCFI/SSA1-2010"
        }
    
    def buscar_en_tabla_relacion(self, codigo):
        """Busca un código en la tabla de relación (que es una lista)"""
        for item in self.tabla_relacion:
            # Intentar buscar por EAN primero
            if str(item.get('EAN', '')).strip() == str(codigo).strip():
                return item
            # Si no encuentra por EAN, intentar por CODIGO
            if str(item.get('CODIGO', '')).strip() == str(codigo).strip():
                return item
        return None

    def buscar_en_tabla_relacion_compuesta(self, codigo, solicitud='', marca='', pais_origen=''):
        """Busca un registro en la tabla de relación usando una clave compuesta.

        Intenta encontrar una entrada que coincida en `CODIGO`/`EAN` y además en
        `SOLICITUD`, `MARCA` y `PAIS` cuando esos valores estén presentes.
        Si no encuentra una coincidencia estricta, hace fallback a `buscar_en_tabla_relacion(codigo)`.
        """
        def _norm(v):
            try:
                return str(v or '').strip().upper()
            except Exception:
                return ''

        target_codigo = _norm(codigo)
        target_solicitud = _norm(solicitud)
        target_marca = _norm(marca)
        target_pais = _norm(pais_origen)

        # Preferir coincidencias que cumplan todos los campos proporcionados
        best = None
        import re as _re
        def _digits(s):
            try:
                d = _re.findall(r"\d+", str(s or ''))
                return d[0] if d else ''
            except Exception:
                return ''

        for item in self.tabla_relacion:
            ean = _norm(item.get('EAN'))
            cod = _norm(item.get('CODIGO'))
            if not (ean == target_codigo or cod == target_codigo):
                continue

            # comprobar campos opcionales
            sol_i_raw = item.get('SOLICITUD') or item.get('Solicitud') or item.get('solicitud') or ''
            sol_i = _norm(sol_i_raw)
            marca_i = _norm(item.get('MARCA') or item.get('Marca') or item.get('marca'))
            pais_i = _norm(item.get('PAIS_DE_ORIGEN') or item.get('PAIS') or item.get('PAIS DE ORIGEN'))

            # si se proporcionaron valores y coinciden, devolver inmediatamente
            # comparar solicitudes preferentemente por dígitos (ej. '000191/26' vs '191')
            target_sol_digits = _digits(target_solicitud)
            sol_i_digits = _digits(sol_i)
            if target_sol_digits and sol_i_digits:
                ok_sol = (not target_solicitud) or (sol_i_digits == target_sol_digits)
            else:
                ok_sol = (not target_solicitud) or (sol_i == target_solicitud)
            ok_marca = (not target_marca) or (marca_i == target_marca)
            ok_pais = (not target_pais) or (pais_i == target_pais)

            if ok_sol and ok_marca and ok_pais:
                return item

            # mantener el primer match por código como fallback
            if best is None:
                best = item

        # fallback: devolver primer match por código si no hubo coincidencia compuesta
        if best:
            return best
        return None
    
    def buscar_producto_por_ean(self, ean):
        """Busca un producto en la base por EAN"""
        for producto in self.base_etiquetado:
            if str(producto.get('EAN', '')).strip() == str(ean).strip():
                return producto
        return None
    
    def determinar_norma_por_uva(self, norma_uva, producto):
        """Determina la norma específica basada en NORMA UVA"""
        if norma_uva in self.mapeo_norma_uva:
            mapeo = self.mapeo_norma_uva[norma_uva]
            
            # Si es un diccionario (como norma 4), verificar insumos
            if isinstance(mapeo, dict):
                insumos = producto.get('INSUMOS', '')
                tiene_insumos = insumos and str(insumos).upper() not in ['N/A', '', 'NaN']
                
                if tiene_insumos:
                    return mapeo["con_insumos"]
                else:
                    return mapeo["sin_insumos"]
            else:
                # Si es un string directo, retornarlo
                return mapeo
        
        # Si no hay mapeo específico, usar norma por defecto
        return "NOM-050-SCFI-2004-1"
    
    def formatear_dato(self, campo, valor):
        """Formatea los datos según el campo"""
        if str(valor).upper() in ['N/A', 'NAN', ''] or not valor:
            return None
        # Normalizar nombre de campo y aceptar variantes como
        # 'PAIS DE ORIGEN', 'PAIS_ORIGEN' o 'PAIS ORIGEN'
        try:
            cn = str(campo or '').upper().replace('_', ' ').strip()
        except Exception:
            cn = str(campo or '').upper()

        if cn in ('PAIS DE ORIGEN','PAIS'):
            return f"HECHO EN {valor}"
        
        if campo == 'TALLA':
            return f"TALLA {valor}"
        
        return str(valor)
    
    def cm_a_pixeles(self, cm, dpi=300):
        """Convierte centímetros a píxeles"""
        return int(cm * dpi / 2.54)
    
    def organizar_campos_por_seccion(self, campos, producto):
        """Organiza los campos en tres secciones: encabezado, centro y pie
        Si no hay talla, el importador va al pie"""
        encabezado = ['EAN', 'MARCA','INSUMOS']
        pie_preferente = ['TALLA']
        
        campos_encabezado = []
        campos_centro = []
        campos_pie = []
        
        # Verificar si hay talla
        tiene_talla = False
        for campo in campos:
            if campo == 'TALLA':
                valor = producto.get(campo, '')
                if valor and str(valor).upper() not in ['N/A', 'NAN', '']:
                    tiene_talla = True
                    break
        
        for campo in campos:
            valor = producto.get(campo, '')
            
            if campo == 'EAN':
                campos_encabezado.append(campo)
            elif campo == 'MARCA' and valor and str(valor).upper() not in ['N/A', 'NAN', '']:
                campos_encabezado.append(campo)
            elif campo == 'IMPORTADOR' and not tiene_talla:
                campos_pie.append(campo)
            elif campo == 'TALLA' and tiene_talla:
                campos_pie.append(campo)
            # Todo lo demás va al centro
            else:
                if valor and str(valor).upper() not in ['N/A', 'NAN', '']:
                    campos_centro.append(campo)
        
        return campos_encabezado, campos_centro, campos_pie
    
    def crear_etiqueta(self, producto, config, output_path):
        """Crea una imagen de etiqueta con layout optimizado: encabezado (EAN siempre), centro centrado, pie"""
        ancho_cm, alto_cm = config['tamaño']
        ancho = self.cm_a_pixeles(ancho_cm)
        alto = self.cm_a_pixeles(alto_cm)
        
        img = Image.new('RGB', (ancho, alto), 'white')
        draw = ImageDraw.Draw(img)
        
        # Configurar fuentes
        try:
            font_paths = [
                "arialbd.ttf",  # Arial Bold
                "Arial Bold.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
                "C:/Windows/Fonts/arialbd.ttf"
            ]
            font = None
            for font_path in font_paths:
                try:
                    area = ancho_cm * alto_cm
                    if area < 25:
                        font_size = 22  # Aumentado de 20 a 22
                    elif area < 35:
                        font_size = 26  # Aumentado de 24 a 26
                    else:
                        font_size = 28  # Aumentado de 28 a 30
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except:
                    continue
            if font is None:
                font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        # Dibujar borde
        draw.rectangle([0, 0, ancho-1, alto-1], outline='black', width=2)
        
        margin_x = 50  # Aumentado de 40 a 50 para márgenes más amplios y uniformes
        margin_y = 40  # Aumentado de 35 a 40
        line_height = font.size + 12
        max_caracteres = 30 if ancho_cm < 5 else 40  # Reducido de 35/45 a 30/40
        
        campos_encabezado, campos_centro, campos_pie = self.organizar_campos_por_seccion(config['campos'], producto)
        
        # ============ SECCIÓN ENCABEZADO (EAN siempre + MARCA si existe) ============
        y_actual = margin_y
        for campo in campos_encabezado:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor) if campo != 'EAN' else str(valor)
            
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                for line in lines:
                    if hasattr(draw, 'textbbox'):
                        bbox = draw.textbbox((0, 0), line, font=font)
                        text_width = bbox[2] - bbox[0]
                    else:
                        text_width = draw.textsize(line, font=font)[0]
                    
                    x_centered = (ancho - text_width) / 2
                    # Asegurar que no se salga de los márgenes
                    if x_centered < margin_x:
                        x_centered = margin_x
                    if x_centered + text_width > ancho - margin_x:
                        x_centered = margin_x
                    
                    draw.text((x_centered, y_actual), line, font=font, fill='black')
                    y_actual += line_height
        
        y_actual += 30
        
        # ============ CALCULAR ESPACIO PARA PIE ============
        lineas_pie = []
        for campo in campos_pie:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor)
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                lineas_pie.extend(lines)
        
        altura_pie = len(lineas_pie) * line_height + margin_y if lineas_pie else margin_y
        
        # ============ SECCIÓN CENTRO (INSUMOS, PAIS ORIGEN centrado) ============
        y_centro_inicio = y_actual
        altura_disponible_centro = alto - y_actual - altura_pie
        
        lineas_centro_total = []
        for campo in campos_centro:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor)
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                lineas_centro_total.extend([(campo, line) for line in lines])
        
        altura_contenido_centro = len(lineas_centro_total) * (line_height + 5)
        
        # Si hay espacio extra, centrar verticalmente
        if altura_contenido_centro < altura_disponible_centro:
            y_actual += (altura_disponible_centro - altura_contenido_centro) / 2
        
        # Dibujar líneas del centro
        for campo, line in lineas_centro_total:
            if y_actual >= alto - altura_pie - margin_y:
                break
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            # Asegurar que no se salga de los márgenes
            if x_centered < margin_x:
                x_centered = margin_x
            if x_centered + text_width > ancho - margin_x:
                x_centered = margin_x
            
            draw.text((x_centered, y_actual), line, font=font, fill='black')
            y_actual += line_height + 5
        
        # ============ SECCIÓN PIE (TALLA o IMPORTADOR) ============
        y_pie = alto - margin_y
        for line in reversed(lineas_pie):
            y_pie -= line_height
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            if x_centered < margin_x:
                x_centered = margin_x
            if x_centered + text_width > ancho - margin_x:
                x_centered = margin_x
            
            draw.text((x_centered, y_pie), line, font=font, fill='black')
        
        # Guardar imagen
        img.save(output_path, 'PNG', dpi=(300, 300))
        return True
    
    def generar_etiquetas_por_codigos(self, codigos, output_dir="etiquetas_generadas", guardar_en_disco=False):
        """Genera etiquetas para una lista de códigos EAN y retorna objetos BytesIO para inserción directa"""
        etiquetas_generadas = []
        for entrada in codigos:
            # aceptar tanto strings como dicts compuestos
            if isinstance(entrada, dict):
                codigo = entrada.get('codigo')
                solicitud = entrada.get('solicitud', '')
                marca = entrada.get('marca', '')
                pais_origen = entrada.get('pais_origen', '')
            else:
                codigo = entrada
                solicitud = ''
                marca = ''
                pais_origen = ''

            print(f"   🔍 Procesando código EAN: {codigo} (solicitud={solicitud}, marca={marca}, pais={pais_origen})")
            
            producto_relacionado = self.buscar_en_tabla_relacion_compuesta(codigo, solicitud=solicitud, marca=marca, pais_origen=pais_origen)
            if not producto_relacionado:
                print(f"      ❌ EAN {codigo} no encontrado en tabla de relación (clave compuesta)")
                continue
            
            norma_uva = producto_relacionado.get('NORMA UVA')
            if norma_uva is None:
                print(f"      ❌ No se encontró NORMA UVA para EAN {codigo}")
                continue
            
            print(f"      📋 NORMA UVA encontrada: {norma_uva}")
            
            # Obtener producto base desde la base de etiquetado (puede ser None)
            producto_base = self.buscar_producto_por_ean(codigo)
            if not producto_base:
                # No existe en la base; generar con los datos mínimos de tabla_relacion
                producto = {}
            else:
                # Copiar para no mutar la base
                producto = dict(producto_base)

            # Campos desde la tabla de relación para respetar
            # marca/pais/descripcion/insumos específicos por solicitud.
            # Esto evita reutilizar la misma etiqueta base cuando el mismo EAN
            # corresponde a productos distintos según solicitud/marca/pais.
            try:
                # Normalizar y mapear variantes de campos desde tabla_relacion
                mapping = {
                    'MARCA': ['MARCA', 'Marca', 'marca'],
                    # Aceptar múltiples variantes de nombre para el país/origen
                    'PAIS ORIGEN': ['PAIS DE ORIGEN', 'PAIS_DE_ORIGEN'],
                    'DESCRIPCION': ['DESCRIPCION', 'DESCRIPCIÓN'],
                    'INSUMOS': ['INSUMOS', 'INSUMO'],
                    'TALLA': ['TALLA'],
                    'IMPORTADOR': ['IMPORTADOR'],
                    'EAN': ['EAN', 'CODIGO', 'Codigo', 'codigo']
                }

                for target_key, variants in mapping.items():
                    for var in variants:
                        if var in producto_relacionado and producto_relacionado.get(var) not in (None, ''):
                            producto[target_key] = producto_relacionado.get(var)
                            break

                # Asegurar EAN presente (fallback adicional)
                if 'EAN' not in producto or not producto.get('EAN'):
                    producto['EAN'] = producto_relacionado.get('EAN') or producto_relacionado.get('CODIGO') or codigo
            except Exception:
                pass
            
            norma = self.determinar_norma_por_uva(norma_uva, producto)
            if not norma:
                print(f"      ❌ No se pudo determinar la norma para NORMA UVA {norma_uva}")
                continue
            
            print(f"      🏷️ Norma determinada: {norma}")
            
            config = self.configuraciones.get(norma)
            if not config:
                print(f"      ❌ No hay configuración para la norma {norma}")
                continue

            # Usar copia local de la configuración para no mutar el objeto global
            config_local = dict(config)
            campos_config = list(config_local.get('campos', []))

            # Si el producto tiene información de país en alguna variante, pero
            # la configuración no incluye ningún campo de país, añadir 'PAIS DE ORIGEN'
            pais_variants = ['PAIS ORIGEN', 'PAIS DE ORIGEN', 'PAIS', 'PAIS_DE_ORIGEN', 'PAIS ORIGEN']
            has_pais = any(producto.get(k) for k in pais_variants)
            if has_pais and not any(v in campos_config for v in ('PAIS DE ORIGEN', 'PAIS_ORIGEN', 'PAIS ORIGEN', 'PAIS')):
                # Insertar antes de IMPORTADOR si existe, sino al final
                try:
                    idx = campos_config.index('IMPORTADOR')
                    campos_config.insert(idx, 'PAIS DE ORIGEN')
                except ValueError:
                    campos_config.append('PAIS DE ORIGEN')
                print(f"      ℹ️ Añadido 'PAIS DE ORIGEN' a campos de la norma {norma} para el código {codigo}")

            # Actualizar la copia local de la config con los campos ajustados
            config_local['campos'] = campos_config

            # Asegurar que los campos esperados por la configuración también
            # estén presentes en `producto` usando nuestras claves canónicas.
            # Ej: la config puede usar 'PAIS DE ORIGEN' pero nosotros normalizamos
            # a 'PAIS ORIGEN' desde la tabla_relacion; copiar ambos para compat.
            try:
                # Mapeo de canónicas a posibles variantes que aparecen en configs
                canonical_map = {
                    'PAIS ORIGEN': ['PAIS DE ORIGEN', 'PAIS_ORIGEN', 'PAIS ORIGEN', 'PAIS'],
                    'DESCRIPCION': ['DESCRIPCION', 'DESCRIPCIÓN'],
                    'INSUMOS': ['INSUMOS', 'INSUMOS O INGREDIENTES', 'INSUMOS O INGREDIENTES'],
                    'TALLA': ['TALLA']
                }

                for canon, variants in canonical_map.items():
                    if producto.get(canon):
                        for var in variants:
                            if var in campos_config and not producto.get(var):
                                producto[var] = producto.get(canon)
            except Exception:
                pass
            
            try:
                # Crear imagen en memoria
                ancho_cm, alto_cm = config_local['tamaño']
                ancho = self.cm_a_pixeles(ancho_cm)
                alto = self.cm_a_pixeles(alto_cm)
                
                img = Image.new('RGB', (ancho, alto), 'white')
                draw = ImageDraw.Draw(img)
                
                # Reutilizar la lógica de dibujo (usar config_local ajustada)
                self._dibujar_etiqueta_en_imagen(img, draw, producto, config_local)
                
                # Guardar en BytesIO en lugar de archivo
                img_bytes = BytesIO()
                img.save(img_bytes, format='PNG', dpi=(300, 300))
                img_bytes.seek(0)
                
                etiquetas_generadas.append({
                    'codigo': codigo,
                    'ean': producto.get('EAN'),
                    'norma': norma,
                    'imagen_bytes': img_bytes,
                    'tamaño_cm': config_local['tamaño']
                })
                print(f"      ✅ Etiqueta generada en memoria")
            except Exception as e:
                print(f"      ❌ Error generando etiqueta para {codigo}: {e}")
                import traceback
                traceback.print_exc()
        
        return etiquetas_generadas
    
    def _dibujar_etiqueta_en_imagen(self, img, draw, producto, config):
        """Dibuja el contenido de la etiqueta en la imagen proporcionada"""
        ancho, alto = img.size
        ancho_cm, alto_cm = config['tamaño']
        
        try:
            font_paths = [
                "arialbd.ttf",
                "Arial Bold.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
                "C:/Windows/Fonts/arialbd.ttf"
            ]
            font = None
            for font_path in font_paths:
                try:
                    area = ancho_cm * alto_cm
                    if area < 25:
                        font_size = 22
                    elif area < 35:
                        font_size = 26
                    else:
                        font_size = 30
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except:
                    continue
            if font is None:
                font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        # Dibujar borde
        draw.rectangle([0, 0, ancho-1, alto-1], outline='black', width=2)
        
        margin_x = 50  # Aumentado de 40 a 50 para márgenes más amplios y uniformes
        margin_y = 40  # Aumentado de 35 a 40
        line_height = font.size + 10
        
        ancho_disponible = ancho - (2 * margin_x)
        
        # Calcular max_caracteres basado en el ancho real de la fuente
        char_width_estimate = font.size * 0.6  # Estimación del ancho promedio de carácter
        max_caracteres = int(ancho_disponible / char_width_estimate)
        max_caracteres = max(20, min(max_caracteres, 45))  # Entre 20 y 45 caracteres
        
        campos_encabezado, campos_centro, campos_pie = self.organizar_campos_por_seccion(config['campos'], producto)
        
        # ENCABEZADO
        y_actual = margin_y
        for campo in campos_encabezado:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor) if campo != 'EAN' else str(valor)
            
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                for line in lines:
                    if hasattr(draw, 'textbbox'):
                        bbox = draw.textbbox((0, 0), line, font=font)
                        text_width = bbox[2] - bbox[0]
                    else:
                        text_width = draw.textsize(line, font=font)[0]
                    
                    x_centered = (ancho - text_width) / 2
                    # Asegurar que no se salga de los márgenes
                    if x_centered < margin_x:
                        x_centered = margin_x
                    if x_centered + text_width > ancho - margin_x:
                        x_centered = margin_x
                    
                    draw.text((x_centered, y_actual), line, font=font, fill='black')
                    y_actual += line_height
        
        y_actual += 10
        
        # CALCULAR PIE
        lineas_pie = []
        for campo in campos_pie:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor)
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                lineas_pie.extend(lines)
        
        altura_pie = len(lineas_pie) * line_height + margin_y if lineas_pie else margin_y
        
        # CENTRO
        altura_disponible_centro = alto - y_actual - altura_pie
        
        lineas_centro_total = []
        for campo in campos_centro:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor)
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                lineas_centro_total.extend([(campo, line) for line in lines])
        
        altura_contenido_centro = len(lineas_centro_total) * (line_height + 5)
        
        if altura_contenido_centro < altura_disponible_centro:
            y_actual += (altura_disponible_centro - altura_contenido_centro) / 2
        
        for campo, line in lineas_centro_total:
            if y_actual >= alto - altura_pie - margin_y:
                break
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            if x_centered < margin_x:
                x_centered = margin_x
            if x_centered + text_width > ancho - margin_x:
                x_centered = margin_x
            
            draw.text((x_centered, y_actual), line, font=font, fill='black')
            y_actual += line_height + 5
        
        # PIE
        y_pie = alto - margin_y
        for line in reversed(lineas_pie):
            y_pie -= line_height
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            if x_centered < margin_x:
                x_centered = margin_x
            if x_centered + text_width > ancho - margin_x:
                x_centered = margin_x
            
            draw.text((x_centered, y_pie), line, font=font, fill='black')
    
    def crear_pdf_etiquetas(self, etiquetas_generadas, output_pdf="etiquetas.pdf"):
        """Crea un PDF con todas las etiquetas generadas"""
        try:
            c = canvas.Canvas(output_pdf, pagesize=letter)
            ancho_pagina, alto_pagina = letter
            
            x = 1 * cm
            y = alto_pagina - 2 * cm
            max_y = 2 * cm
            
            for i, etiqueta in enumerate(etiquetas_generadas):
                if y < max_y:
                    c.showPage()
                    y = alto_pagina - 2 * cm
                    x = 1 * cm
                
                ancho_cm, alto_cm = etiqueta['tamaño_cm']
                ancho_pt = ancho_cm * 28.35
                alto_pt = alto_cm * 28.35
                
                try:
                    c.drawImage(ImageReader(etiqueta['imagen_bytes']), x, y - alto_pt, width=ancho_pt, height=alto_pt)
                except Exception:
                    try:
                        c.drawImage(etiqueta['imagen_bytes'], x, y - alto_pt, width=ancho_pt, height=alto_pt)
                    except Exception:
                        continue
                
                x += ancho_pt + 0.5 * cm
                
                if x + ancho_pt > ancho_pagina - 1 * cm:
                    x = 1 * cm
                    y -= alto_pt + 0.5 * cm
            
            c.save()
            print(f"✅ PDF creado: {output_pdf}")
            return True
        except Exception as e:
            print(f"❌ Error creando PDF: {e}")
            return False

def main():
    if not os.path.exists("data"):
        print("❌ Carpeta 'data' no encontrada")
        return
    
    generador = GeneradorEtiquetasDecathlon()
    codigos_a_procesar = ["4714062"]
    
    print("Generando etiquetas...")
    resultados = generador.generar_etiquetas_por_codigos(codigos_a_procesar)
    
    print(f"\n--- RESULTADOS ---")
    print(f"Total de etiquetas generadas: {len(resultados)}")
    for resultado in resultados:
        print(f"✓ {resultado['ean']} - Norma: {resultado['norma']}")
    
    if resultados:
        print("\nCreando PDF con etiquetas...")
        generador.crear_pdf_etiquetas(resultados, "etiquetas_decathlon.pdf")
    
    print("\n🎉 Proceso completado!")

    if resultados:
        generador.insertar_etiquetas_en_dictamen(
            dictamen_path="Dictamen_Lista_4_nan_007045_25_5.pdf",
            etiquetas=resultados,
            output_pdf="DICTAMEN_FINAL.pdf"
    )

if __name__ == "__main__":
    main()
