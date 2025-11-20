# Generador de Dictámenes con Etiquetas Integradas

Sistema completo para generar dictámenes en PDF con etiquetas visuales automáticas.

## 📋 Características

- **Generación automática de etiquetas**: Crea imágenes PNG de etiquetas basándose en códigos EAN
- **Integración en PDF**: Inserta las etiquetas como imágenes en la segunda página del dictamen
- **Datos dinámicos**: Extrae información de múltiples fuentes JSON
- **Multi-familia**: Procesa múltiples dictámenes en lote

## 🗂️ Estructura del Proyecto

\`\`\`
proyecto/
├── data/                          # Carpeta con datos de entrada
│   ├── TABLA_DE_RELACION.json    # Códigos y productos
│   ├── BASE_ETIQUETADO.json      # Información de etiquetas por EAN
│   ├── config_etiquetas.json     # Configuración de tamaños y campos
│   ├── Normas.json               # Catálogo de normas oficiales
│   └── Clientes.json             # Información de clientes y RFC
├── img/
│   └── Fondo.jpeg                # Imagen de fondo para el PDF
├── etiquetas_generadas/          # Etiquetas PNG generadas (creada automáticamente)
├── dictamenes_generados/         # PDFs de salida (creada automáticamente)
│
├── etiqueta_dictamen.py          # Generador de imágenes de etiquetas
├── plantillaPDF.py               # Funciones de carga y preparación de datos
├── DictamenPDF.py                # Clase base para generación de PDF
├── PDFGeneradorConDatos.py       # Generador principal con datos reales
└── main.py                       # Script principal de ejecución
\`\`\`

## 🚀 Instalación

1. Instalar dependencias:

\`\`\`bash
pip install reportlab pandas pillow
\`\`\`

2. Crear la estructura de carpetas:

\`\`\`bash
mkdir -p data img etiquetas_generadas dictamenes_generados
\`\`\`

3. Colocar los archivos JSON en la carpeta `data/`
4. Colocar la imagen `Fondo.jpeg` en la carpeta `img/`

## 📝 Uso

### Ejecución Simple

\`\`\`bash
python main.py
\`\`\`

### Uso Programático

\`\`\`python
from PDFGeneradorConDatos import generar_dictamenes_completos

# Generar dictámenes
exito, mensaje, resultado = generar_dictamenes_completos("carpeta_salida")

if exito:
    print(f"✅ {mensaje}")
    print(f"Generados: {resultado['total_generados']} dictámenes")
\`\`\`

## 🏷️ Formato de Etiquetas

Las etiquetas se generan automáticamente en formato PNG con:
- Tamaño configurable por norma
- Texto centrado
- Borde negro
- Campos dinámicos (país, talla, composición, etc.)

### Configuración de Etiquetas (config_etiquetas.json)

\`\`\`json
{
  "NOM-024-SCFI-2013": {
    "tamaño_cm": "(5.0, 5.0)",
    "campos": ["MARCA", "PAIS ORIGEN", "TALLA", "COMPOSICION"]
  }
}
\`\`\`

## 📄 Estructura del Dictamen PDF

### Página 1
- Encabezado con código de identificación
- Fechas de inspección y emisión
- Cliente y RFC
- Texto legal del dictamen
- Tabla de productos
- Tamaño del lote
- Observaciones

### Página 2
- **Etiquetas del producto** (imágenes PNG insertadas)
- Imágenes del producto (placeholders)
- Firmas del inspector y responsable

## 🔧 Flujo de Procesamiento

1. **Carga de datos**: Lee archivos JSON de `data/`
2. **Procesamiento de familias**: Agrupa registros por NORMA UVA, FOLIO, SOLICITUD y LISTA
3. **Generación de etiquetas**: 
   - Busca códigos EAN en BASE_ETIQUETADO.json
   - Determina la norma aplicable
   - Genera imágenes PNG en `etiquetas_generadas/`
4. **Construcción del PDF**:
   - Primera página con datos del dictamen
   - Segunda página con etiquetas como imágenes
   - Fondo y marcas de agua
5. **Salida**: PDFs en `dictamenes_generados/`

## 🐛 Solución de Problemas

### "No se generaron etiquetas"

**Causa**: Los códigos EAN no se encuentran en BASE_ETIQUETADO.json

**Solución**: Verificar que los códigos en TABLA_DE_RELACION.json coincidan con los EAN en BASE_ETIQUETADO.json

### Las imágenes no aparecen en el PDF

**Causa**: Las rutas de las imágenes generadas no son correctas

**Solución**: Verificar que la carpeta `etiquetas_generadas/` tenga los archivos PNG

### Error al cargar normas

**Causa**: Formato incorrecto en Normas.json

**Solución**: Verificar que cada norma tenga los campos: NOM, NOMBRE, CAPITULO

## 📊 Ejemplo de Salida

\`\`\`
🚀 INICIANDO GENERACIÓN DE DICTÁMENES
============================================================
📂 Cargando datos...
✅ Tabla de relación cargada: 150 registros
✅ Normas cargadas correctamente: 10 mapeos
✅ Clientes cargados: 5

🛠️  Generando 3 dictámenes...

📄 Procesando familia LISTA 24_001_2025_1 (10 registros)...
Procesando código: 8123456789012
  ✅ Etiqueta generada: 8123456789012_NOM-024-SCFI-2013.png
   🏷️ Insertando 1 etiquetas en el PDF...
   ✅ Etiqueta cargada: 8123456789012_NOM-024-SCFI-2013.png
   ✅ Creado: Dictamen_Lista_24_001_2025_1.pdf

============================================================
✅ PROCESO COMPLETADO EXITOSAMENTE

📊 Resumen:
   • Dictámenes generados: 3
   • Total de familias: 3
   • Ubicación: dictamenes_generados/
\`\`\`

## 🤝 Contribuciones

Para agregar nuevas normas o campos de etiquetas, editar:
- `config_etiquetas.json` - Configuración de campos por norma
- `etiqueta_dictamen.py` - Método `crear_mapeo_norma_uva()` para nuevas normas

## 📞 Soporte

Si el mensaje "No se generaron etiquetas" persiste:
1. Verificar que los códigos EAN existan en BASE_ETIQUETADO.json
2. Revisar que NORMA UVA esté en el mapeo de normas
3. Comprobar que config_etiquetas.json tenga la configuración de la norma
