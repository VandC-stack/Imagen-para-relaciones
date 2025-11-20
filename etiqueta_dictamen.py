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
        """Formatea los datos según el campo - SOLO VALORES, SIN ETIQUETAS"""
        if str(valor).upper() in ['N/A', 'NAN', ''] or not valor:
            return None
        
        # Para PAIS ORIGEN, mantener el formato "HECHO EN"
        if campo == 'PAIS ORIGEN':
            return f"HECHO EN {valor}"
        
        if campo == 'TALLA':
            return f"TALLA {valor}"
        
        # Para todos los demás campos, devolver solo el valor sin etiqueta
        return str(valor)
    
    def cm_a_pixeles(self, cm, dpi=300):
        """Convierte centímetros a píxeles"""
        return int(cm * dpi / 2.54)
    
    def crear_etiqueta(self, producto, config, output_path):
        """Crea una imagen de etiqueta con texto centrado"""
        # Convertir tamaño a píxeles
        ancho_cm, alto_cm = config['tamaño']
        ancho = self.cm_a_pixeles(ancho_cm)
        alto = self.cm_a_pixeles(alto_cm)
        
        # Crear imagen
        img = Image.new('RGB', (ancho, alto), 'white')
        draw = ImageDraw.Draw(img)
        
        # Configurar fuente - tamaño adaptable según el tamaño de la etiqueta
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
                    # Tamaño de fuente basado en el área de la etiqueta
                    area = ancho_cm * alto_cm
                    if area < 25:  # Etiquetas muy pequeñas
                        font_size = 15
                    elif area < 35:  # Etiquetas pequeñas
                        font_size = 15
                    else:  # Etiquetas normales
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
        
        # Posición inicial
        margin = 30
        y = margin
        line_height = font.size + 10
        espacio_entre_campos = 15
        
        # Recolectar todas las líneas que vamos a mostrar
        lineas_totales = []
        for campo in config['campos']:
            valor = producto.get(campo, '')
            texto_formateado = self.formatear_dato(campo, valor)
            
            if texto_formateado:
                # Dividir texto si es muy largo
                max_caracteres = 25 if ancho_cm < 5 else 35
                lines = textwrap.wrap(texto_formateado, width=max_caracteres)
                lineas_totales.extend(lines)
                lineas_totales.append("")  # Línea en blanco entre campos
        
        # Eliminar la última línea en blanco si existe
        if lineas_totales and lineas_totales[-1] == "":
            lineas_totales.pop()
        
        # Calcular la altura total del texto
        altura_total = len(lineas_totales) * line_height
        # Calcular la posición Y inicial para centrar verticalmente
        y_inicio = max(margin, (alto - altura_total) / 2)
        y = y_inicio
        
        # Dibujar todas las líneas centradas
        campos_mostrados = 0
        for line in lineas_totales:
            if line == "":
                y += espacio_entre_campos
                continue
                
            # Calcular ancho del texto para centrarlo
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                # Para versiones antiguas de PIL
                text_width = draw.textsize(line, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            
            # Verificar si hay espacio suficiente
            if y + line_height <= alto - margin:
                draw.text((x_centered, y), line, font=font, fill='black')
                y += line_height
                campos_mostrados += 1
        
        # Si no se pudo mostrar ningún campo, mostrar mensaje de error centrado
        if campos_mostrados == 0:
            error_msg = "No hay datos válidos"
            if hasattr(draw, 'textbbox'):
                bbox = draw.textbbox((0, 0), error_msg, font=font)
                text_width = bbox[2] - bbox[0]
            else:
                text_width = draw.textsize(error_msg, font=font)[0]
            
            x_centered = (ancho - text_width) / 2
            y_centered = alto / 2
            draw.text((x_centered, y_centered), error_msg, font=font, fill='red')
        
        # Guardar imagen con alta calidad
        img.save(output_path, 'PNG', dpi=(300, 300))
        return True
    
    def generar_etiquetas_por_codigos(self, codigos, output_dir="etiquetas_generadas"):
        """Genera etiquetas para una lista de códigos EAN"""
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        etiquetas_generadas = []
        
        for codigo in codigos:
            print(f"   🔍 Procesando código EAN: {codigo}")
            
            # Buscar en tabla de relación por EAN
            producto_relacionado = self.buscar_en_tabla_relacion(codigo)
            if not producto_relacionado:
                print(f"      ❌ EAN {codigo} no encontrado en tabla de relación")
                continue
            
            # Obtener NORMA UVA
            norma_uva = producto_relacionado.get('NORMA UVA')
            if norma_uva is None:
                print(f"      ❌ No se encontró NORMA UVA para EAN {codigo}")
                continue
            
            print(f"      📋 NORMA UVA encontrada: {norma_uva}")
            
            # Buscar producto en base de etiquetado por EAN
            producto = self.buscar_producto_por_ean(codigo)
            if not producto:
                print(f"      ❌ Producto con EAN {codigo} no encontrado en base de etiquetado")
                continue
            
            # Determinar norma específica basada en NORMA UVA
            norma = self.determinar_norma_por_uva(norma_uva, producto)
            if not norma:
                print(f"      ❌ No se pudo determinar la norma para NORMA UVA {norma_uva}")
                continue
            
            print(f"      🏷️ Norma determinada: {norma}")
            
            # Buscar configuración
            config = self.configuraciones.get(norma)
            if not config:
                print(f"      ❌ No hay configuración para la norma {norma}")
                continue
            
            # Generar etiqueta
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
            
            # Posición inicial en la página
            x = 1 * cm
            y = alto_pagina - 2 * cm
            max_y = 2 * cm
            
            for i, etiqueta in enumerate(etiquetas_generadas):
                # Verificar si necesitamos nueva página
                if y < max_y:
                    c.showPage()
                    y = alto_pagina - 2 * cm
                    x = 1 * cm
                
                # Obtener tamaño de la etiqueta en cm
                ancho_cm, alto_cm = etiqueta['tamaño_cm']
                
                # Convertir a puntos (1 cm = 28.35 puntos)
                ancho_pt = ancho_cm * 28.35
                alto_pt = alto_cm * 28.35
                
                # Insertar imagen de la etiqueta
                c.drawImage(etiqueta['ruta'], x, y - alto_pt, width=ancho_pt, height=alto_pt)
                
                # Mover posición para la siguiente etiqueta
                x += ancho_pt + 0.5 * cm
                
                # Si no cabe en el ancho, pasar a siguiente fila
                if x + ancho_pt > ancho_pagina - 1 * cm:
                    x = 1 * cm
                    y -= alto_pt + 0.5 * cm
            
            c.save()
            print(f"✅ PDF creado: {output_pdf}")
            return True
        except Exception as e:
            print(f"❌ Error creando PDF: {e}")
            return False

# Función principal de uso
def main():
    # Verificar que exista la carpeta data
    if not os.path.exists("data"):
        print("❌ Carpeta 'data' no encontrada")
        print("   Por favor, crea la carpeta 'data' y coloca allí los archivos:")
        print("   - BASE_ETIQUETADO.json")
        print("   - TABLA_DE_RELACION.json") 
        print("   - config_etiquetas.json")
        return
    
    # Inicializar generador (ahora lee automáticamente de la carpeta data)
    generador = GeneradorEtiquetasDecathlon()
    
    # Lista de códigos a procesar
    codigos_a_procesar = ["692071"]  # Ejemplos
    
    # Generar etiquetas
    print("Generando etiquetas...")
    resultados = generador.generar_etiquetas_por_codigos(codigos_a_procesar)
    
    # Mostrar resultados
    print(f"\n--- RESULTADOS ---")
    print(f"Total de etiquetas generadas: {len(resultados)}")
    for resultado in resultados:
        print(f"✓ {resultado['archivo']} - EAN: {resultado['ean']} - Norma: {resultado['norma']}")
    
    # Crear PDF con todas las etiquetas
    if resultados:
        print("\nCreando PDF con etiquetas...")
        generador.crear_pdf_etiquetas(resultados, "etiquetas_decathlon.pdf")
    
    print("\n🎉 Proceso completado!")

if __name__ == "__main__":
    main()
