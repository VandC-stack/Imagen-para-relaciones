import json
import os
from PIL import Image, ImageDraw, ImageFont
import textwrap
import ast
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm

class GeneradorEtiquetasDecathlon:
    def __init__(self):
        # Rutas dentro de la carpeta data
        self.data_dir = "data"
        base_etiquetado_path = os.path.join(self.data_dir, "BASE_ETIQUETADO.json")
        tabla_relacion_path = os.path.join(self.data_dir, "TABLA_DE_RELACION.json")
        config_etiquetas_path = os.path.join(self.data_dir, "config_etiquetas.json")
        
        self.cargar_datos(base_etiquetado_path, tabla_relacion_path)
        self.configuraciones = self.cargar_configuraciones(config_etiquetas_path)
        self.mapeo_norma_uva = self.crear_mapeo_norma_uva()
        
    def cargar_datos(self, base_etiquetado_path, tabla_relacion_path):
        """Carga los datos de la base de etiquetado y tabla de relación"""
        try:
            with open(base_etiquetado_path, 'r', encoding='utf-8') as f:
                self.base_etiquetado = json.load(f)
            with open(tabla_relacion_path, 'r', encoding='utf-8') as f:
                self.tabla_relacion = json.load(f)
            print("✅ Archivos JSON cargados correctamente")
        except Exception as e:
            print(f"❌ Error cargando archivos: {e}")
            self.base_etiquetado = []
            self.tabla_relacion = []
    
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
        
        if campo == 'PAIS ORIGEN':
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
        encabezado = ['EAN', 'MARCA']
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
                "arial.ttf", 
                "Arial.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                "C:/Windows/Fonts/arial.ttf"
            ]
            font = None
            for font_path in font_paths:
                try:
                    area = ancho_cm * alto_cm
                    if area < 25:
                        font_size = 14
                    elif area < 35:
                        font_size = 16
                    else:
                        font_size = 20
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
        
        margin_x = 25
        margin_y = 25
        line_height = font.size + 14
        max_caracteres = 35 if ancho_cm < 5 else 45
        
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
                    draw.text((x_centered, y_actual), line, font=font, fill='black')
                    y_actual += line_height
        
        y_actual += 10
        
        # ============ CALCULAR ESPACIO PARA PIE ============
        lineas_pie = []
        for campo in campos_pie:
            valor = producto.get(campo, '')
            texto = self.formatear_dato(campo, valor)
            if texto:
                lines = textwrap.wrap(texto, width=max_caracteres)
                lineas_pie.extend(lines)
        
        altura_pie = len(lineas_pie) * line_height + margin_y
        
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
            if y_actual >= alto - altura_pie:
                break
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            draw.text((x_centered, y_actual), line, font=font, fill='black')
            y_actual += line_height + 5
        
        # ============ SECCIÓN PIE (TALLA o IMPORTADOR) ============
        y_pie = alto - margin_y
        
        # Dibujar desde abajo hacia arriba
        for line in reversed(lineas_pie):
            y_pie -= line_height
            
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            draw.text((x_centered, y_pie), line, font=font, fill='black')
        
        # Guardar imagen
        img.save(output_path, 'PNG', dpi=(300, 300))
        return True
    
    def generar_etiquetas_por_codigos(self, codigos, output_dir="etiquetas_generadas", guardar_en_disco=False):
        """Genera etiquetas para una lista de códigos EAN"""
        if guardar_en_disco and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        etiquetas_generadas = []
        
        for codigo in codigos:
            print(f"   🔍 Procesando código EAN: {codigo}")
            
            producto_relacionado = self.buscar_en_tabla_relacion(codigo)
            if not producto_relacionado:
                print(f"      ❌ EAN {codigo} no encontrado en tabla de relación")
                continue
            
            norma_uva = producto_relacionado.get('NORMA UVA')
            if norma_uva is None:
                print(f"      ❌ No se encontró NORMA UVA para EAN {codigo}")
                continue
            
            print(f"      📋 NORMA UVA encontrada: {norma_uva}")
            
            producto = self.buscar_producto_por_ean(codigo)
            if not producto:
                print(f"      ❌ Producto con EAN {codigo} no encontrado en base de etiquetado")
                continue
            
            norma = self.determinar_norma_por_uva(norma_uva, producto)
            if not norma:
                print(f"      ❌ No se pudo determinar la norma para NORMA UVA {norma_uva}")
                continue
            
            print(f"      🏷️ Norma determinada: {norma}")
            
            config = self.configuraciones.get(norma)
            if not config:
                print(f"      ❌ No hay configuración para la norma {norma}")
                continue
            
            nombre_archivo = f"{codigo}_{norma}.png"
            output_path = os.path.join(output_dir, nombre_archivo)
            
            try:
                success = self.crear_etiqueta(producto, config, output_path)
                if success:
                    etiquetas_generadas.append({
                        'codigo': codigo,
                        'ean': producto.get('EAN'),
                        'norma': norma,
                        'archivo': nombre_archivo,
                        'ruta': output_path,
                        'tamaño_cm': config['tamaño']
                    })
                    print(f"      ✅ Etiqueta generada: {nombre_archivo}")
                else:
                    print(f"      ❌ Error generando etiqueta para {codigo}")
            except Exception as e:
                print(f"      ❌ Error generando etiqueta para {codigo}: {e}")
                import traceback
                traceback.print_exc()
        
        return etiquetas_generadas
    
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
                
                c.drawImage(etiqueta['ruta'], x, y - alto_pt, width=ancho_pt, height=alto_pt)
                
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
    codigos_a_procesar = ["692071"]
    
    print("Generando etiquetas...")
    resultados = generador.generar_etiquetas_por_codigos(codigos_a_procesar)
    
    print(f"\n--- RESULTADOS ---")
    print(f"Total de etiquetas generadas: {len(resultados)}")
    for resultado in resultados:
        print(f"✓ {resultado['archivo']} - EAN: {resultado['ean']} - Norma: {resultado['norma']}")
    
    if resultados:
        print("\nCreando PDF con etiquetas...")
        generador.crear_pdf_etiquetas(resultados, "etiquetas_decathlon.pdf")
    
    print("\n🎉 Proceso completado!")

if __name__ == "__main__":
    main()
