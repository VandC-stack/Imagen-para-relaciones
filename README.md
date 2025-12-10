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

## 🧭 Documentación del Código (desarrolladores)

Esta sección documenta los archivos principales, responsabilidades y puntos de extensión para que cualquier desarrollador pueda entender y modificar el proyecto.

- **`app.py`**: Interfaz gráfica (CustomTkinter) y orquestador principal.
   - Gestor de UI: pestañas *Principal* y *Historial*.
   - Funcionalidades clave: carga de clientes, preparación de visita, generación de dictámenes (dispara `generador_dictamen.py`), registro y sincronización del `historial_visitas.json`.
   - Módulos importantes: métodos `_cargar_historial`, `_guardar_historial`, `_poblar_historial_ui`, `hist_create_visita`, `hist_eliminar_registro`, `registrar_visita_automatica`.
   - Notas: la UI ya no contiene campo `Supervisor` manual; el inspector se determina desde `data/tabla_de_relacion.json` y `data/Firmas.json` cuando se generan dictámenes.

- **`generador_dictamen.py`**: Lógica que procesa los datos y genera los PDFs (usa ReportLab y plantillas).
   - Provee `generar_dictamenes_gui` y funciones auxiliares para construir tablas, calcular páginas y crear contenido dinámico.
   - Integra `plantillaPDF.py`, `DictamenPDF.py` y `etiqueta_dictamen.py` para componer documentos completos.

- **`plantillaPDF.py`**: Funciones de carga y preparación de datos.
   - Lectura de `data/tabla_de_relacion.json`, `data/Normas.json`, `data/Clientes.json`, `data/Firmas.json`.
   - Funciones: `cargar_tabla_relacion`, `cargar_normas`, `cargar_clientes`, `cargar_firmas`, `preparar_datos_familia`.
   - Normaliza y transforma los registros para que el generador tenga la estructura esperada.

- **`DictamenPDF.py`**: Clase base para generación de PDF con ReportLab.
   - Define estilos, layout y utilidades para encabezados, pies de página y paginación.
   - Se extiende desde `PDFGeneratorConDatos` en `generador_dictamen.py` para adaptarse a datos reales.

- **`etiqueta_dictamen.py`**: Generador de imágenes de etiquetas (Pillow).
   - Encargado de renderizar etiquetas PNG a partir de `BASE_ETIQUETADO.json` y `config_etiquetas.json`.
   - Métodos clave: `crear_mapeo_norma_uva`, `crear_etiqueta`, `generar_etiquetas_por_codigos`.

- **`data/`**: Carpeta con los JSON que alimentan el sistema.
   - `tabla_de_relacion.json`: tabla principal con filas para cada folio/solicitud (entradas usadas para generar dictámenes).
   - `Firmas.json`: mapeo FIRMA → NOMBRE DE INSPECTOR (usado para mostrar el inspector detectado en el historial).
   - `historial_visitas.json`: historial persistente de visitas (creado y mantenido por `app.py`).
   - `folios_visitas/`: archivos `folios_{CPxxxxx}.json` con listado de folios asociados a una visita; usados para eliminar persistencia por visita.

- **`Pegado de Evidenvia Fotografica/`**: utilidades para procesamiento de documentos e inserción de imágenes (dividido en `interfaz.py`, `main.py`, `pegado_*` y `registro_fallos.py`).
   - `interfaz.py`: UI para el módulo de imágenes.
   - `main.py`: utilidades centrales (indexado de imágenes, extracción de códigos, helpers para DOCX/PDF).

- **Otros**:
   - `DictamenMachote.py`, `Armado.py`, `DictamenPDF.py` (plantillas y utilidades históricas/auxiliares).
   - `requirements.txt`: dependencias mínimas.

### Flujo interno (resumen técnico)

1. El usuario carga una `tabla_de_relacion` (Excel → JSON) y selecciona un cliente.
2. `generador_dictamen.py` procesa familias, genera etiquetas PNG y construye PDFs mediante `DictamenPDF`.
3. Cuando se generan dictámenes, `app.py` recibe resultados y ejecuta `registrar_visita_automatica` para crear una entrada en `historial_visitas.json`.
4. `hist_eliminar_registro` borra solo la fila seleccionada, elimina `data/folios_visitas/folios_{folio}.json`, hace backup y limpia coincidencias en `data/tabla_de_relacion.json`.

### Puntos de extensión / cómo añadir nuevas normas

- Para agregar una norma nueva que afecte etiquetas:
   1. Añadir la entrada en `data/Normas.json` y en `data/Firmas.json` si aplica.
   2. Actualizar `config_etiquetas.json` con los campos y tamaños de la norma.
   3. Si la lógica es muy específica, extender `etiqueta_dictamen.py::crear_mapeo_norma_uva`.

### Desarrollo y pruebas rápidas

- Instalar dependencias:

```bash
pip install -r requirements.txt
```

- Ejecutar la app (GUI):

```bash
python app.py
```

- Para pruebas unitarias simples (no incluidas en el repo):
   - Puedes escribir scripts que llamen `plantillaPDF.cargar_tabla_relacion()` o `generador_dictamen.generar_dictamenes_completos(...)` con muestras de `data/`.

### Notas de mantenimiento

- Respaldos: antes de modificar `data/tabla_de_relacion.json` el sistema crea copias en `data/tabla_relacion_backups/`.
- Concurrencia: las actualizaciones del UI desde procesos en segundo plano usan `self.after(...)` para evitar problemas con Tkinter.
- Para registrar una operación (audit): consultar `data/operaciones_log.json` (método `_registrar_operacion` en `app.py`).

