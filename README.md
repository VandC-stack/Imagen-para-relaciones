# Sistema Generador de Documentos V&C

Este repositorio contiene el sistema "Sistema generador de documentos V&C", una aplicación en Python para generar documentos oficiales (dictámenes, constancias, oficios, etiquetas, etc.) a partir de plantillas y datos JSON.

## Objetivo

Proveer a un usuario externo (desarrollador o auditor técnico) de una visión clara del funcionamiento del sistema, los requisitos, y cómo se comunican los archivos entre sí, con diagramas y ejemplos de uso.

## Vista rápida

- **Entrada principal:** [app.py](app.py) (interfaz / punto de arranque)
- **Generación de documentos:** [generador_dictamen.py](generador_dictamen.py), [DictamenPDF.py](DictamenPDF.py), [plantillaPDF.py](plantillaPDF.py)
- **Gestión de folios:** [folio_manager.py](folio_manager.py)
- **Etiquetas:** [etiqueta_dictamen.py](etiqueta_dictamen.py)
- **Data y configuraciones:** carpeta `data/` (JSONs: `Clientes.json`, `Firmas.json`, `folio_counter.json`, etc.)

## Requisitos

- Python 3.11+ (probado con Python 3.13 en este entorno)
- Instalar dependencias:

```bash
python -m pip install -r requirements.txt
```

- Herramientas auxiliares:
  - `build_exe.bat` / `build_exe.ps1` para crear ejecutables (pyinstaller/auto). Ver [Sistema_Generador_Documentos_VC.spec](Sistema_Generador_Documentos_VC.spec).

Comprueba el archivo `requirements.txt` para versiones detalladas de librerías necesarias.

## Estructura principal y responsabilidades de archivos

- `app.py` — Punto de arranque. Inicializa la aplicación (CLI/GUI según implementación) y coordina la ejecución.
- `generador_dictamen.py` — Lógica de composición de contenido (reúne datos, estructura el documento antes de renderizar en PDF).
- `DictamenPDF.py` — Funciones y utilidades para producir el PDF final, llamadas a `plantillaPDF.py` y a librerías de PDF.
- `plantillaPDF.py` — Plantillas y layout (coloca texto, imágenes, firmas, tablas).
- `folio_manager.py` — Asigna y registra folios (lee/escribe `data/folio_counter.json` y `data/pending_folios.json`).
- `etiqueta_dictamen.py` — Genera etiquetas/pegatinas para documentación impresa o pegado de evidencia.
- Carpeta `Documentos Inspeccion/` — Tipos específicos de documentos (Acta, Constancia, Negación, Formatos de supervisión). Cada archivo implementa una variante del documento.
- Carpeta `Pegado de Evidenvia Fotografica/` — Herramientas para pegar imágenes e índices en documentos.

## Archivos de datos importantes (carpeta `data/`)

- `Clientes.json` — Datos de clientes.
- `Firmas.json` — Plantillas/archivos de firmas digitales o referencias de imagen.
- `folio_counter.json` — Contador central de folios usado por `folio_manager.py`.
- `historial_visitas.json` — Registro histórico de operaciones/visitas.
- `pending_folios.json` — Folios pendientes por procesar.

Los módulos leen y escriben estas fuentes JSON para persistir estado y configuraciones.

## Flujo general (resumen)

1. El usuario ejecuta [app.py](app.py) o un script específico.
2. La UI/CLI solicita datos (o lee un JSON de entrada) y selecciona el tipo de documento.
3. `folio_manager.py` asigna un folio disponible y actualiza `data/folio_counter.json`.
4. `generador_dictamen.py` compone el contenido del documento usando los datos de `data/` y las plantillas.
5. `DictamenPDF.py` y `plantillaPDF.py` renderizan el PDF final y lo guardan en `data/Dictamenes/` o en la carpeta configurada.
6. Si corresponde, `etiqueta_dictamen.py` produce una etiqueta y la guarda en `etiquetas_generadas/`.

## Diagrama de alto nivel

Diagrama ASCII (módulos y flujo de datos):

```
     [Usuario]
         |
         v
      [app.py]
         |
   +-----+-----+
   |           |
   v           v
 [folio_manager]  [generador_dictamen]
   |                |
   |                v
   |           [plantillaPDF.py]
   |                |
   v                v
 [data/folio_counter.json]  [DictamenPDF.py] ---> data/Dictamenes/*.pdf
                         |
                         v
                 [etiqueta_dictamen.py] -> etiquetas_generadas/
```

Secuencia de lectura/escritura con los JSON:

- `folio_manager.py`: lee/escribe `data/folio_counter.json`, actualiza `pending_folios.json`.
- `generador_dictamen.py`: lee `Clientes.json`, `Firmas.json`, `excel_export_data.json` (si aplica) y fusiona datos.
- `DictamenPDF.py`/`plantillaPDF.py`: consumen la estructura final y generan archivos PDF.

## Diagrama de componentes (texto)

- Interfaz (app.py)
  - Controlador: decide qué generador llamar
- Servicios
  - Folio (folio_manager)
  - Generación (generador_dictamen)
  - Plantillas/PDF (plantillaPDF, DictamenPDF)
  - Etiquetas (etiqueta_dictamen)
  - Pegado de evidencia (Pegado de Evidenvia Fotografica/)
- Persistencia
  - JSONs en `data/`
  - Carpetas de salida: `data/Dictamenes/`, `etiquetas_generadas/`, `etiquetas_generadas/`

## Ejecución rápida

1. Crear entorno virtual (recomendado):

```bash
python -m venv .venv
source .venv/Scripts/activate   # Windows: .venv\\Scripts\\activate
python -m pip install -r requirements.txt
```

2. Ejecutar la aplicación (modo desarrollo):

```bash
python app.py
```

3. Para construir un ejecutable (Windows):

```powershell
.\\build_exe.bat
```

## Ejemplo de caso de uso

- Generar un dictamen nuevo:
  1. Ejecutar `app.py`.
  2. Ingresar o seleccionar cliente.
  3. Seleccionar tipo de documento (ej. Dictamen).
  4. El sistema solicita/valida datos, asigna folio y genera el PDF en `data/Dictamenes/`.

## Notas de integración / cómo se comunican los archivos

- Comunicación entre módulos se realiza por llamadas a funciones (imports locales) y por persistencia en JSON para conservar estado entre ejecuciones.
- Para añadir un nuevo tipo de documento, crear un archivo en `Documentos Inspeccion/` que exponga una función de generación que acepte los datos requeridos y devuelva la estructura que `DictamenPDF.py` pueda renderizar.

## Buenas prácticas para mantener el sistema

- Respaldar `data/folio_counter.json` antes de operaciones masivas.
- Versionar los `Clientes.json` y `Firmas.json` si se realizan cambios manuales.
- Mantener `requirements.txt` actualizado.

## Dónde leer el código relevante

- Punto de entrada: [app.py](app.py)
- Generación: [generador_dictamen.py](generador_dictamen.py)
- PDF / Plantilla: [DictamenPDF.py](DictamenPDF.py), [plantillaPDF.py](plantillaPDF.py)
- Folios: [folio_manager.py](folio_manager.py)

## Próximos pasos sugeridos

- Documento de API interna (opcional): describir funciones públicas de `generador_dictamen.py` y `DictamenPDF.py` con firmas.
- Añadir diagramas gráficos (PlantUML/Mermaid) en la documentación si se desea visualización más rica.

---

Si quieres, puedo:

- Ajustar el README con diagramas PlantUML/Mermaid.
- Extraer y documentar las funciones públicas de los módulos clave.
- Generar un diagrama en formato PNG/SVG para incluir en la documentación.
# Generador de Dictámenes con Etiquetas Integradas

Sistema completo para generar documentos en PDF con etiquetas visuales automáticas.

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

## 🧩 Empaquetado a .exe (Windows)

Se incluye un `app.spec` configurado y un script `build_exe.bat` para generar un ejecutable con PyInstaller.

Pasos rápidos:

1. Crear un entorno virtual y activar:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instalar dependencias (incluye PyInstaller):

```powershell
pip install -r requirements.txt
```

3. Ejecutar el build:

```powershell
.\build_exe.bat
```

Notas importantes:
- `app.spec` incluye las carpetas de datos necesarias (`data`, `Documentos Inspeccion`, `Pegado de Evidenvia Fotografica`, `Firmas`, `img`, `Plantillas PDF`, `etiquetas_generadas`). Si añades otras carpetas con recursos, añádelas a `datas` en `app.spec`.
- Si usas archivos `.xlsb` en Excel necesitarás `pyxlsb` instalado en el entorno de destino.
- El código ya usa `sys._MEIPASS` mediante `plantillaPDF.obtener_ruta_recurso()` para localizar recursos cuando está empacado con PyInstaller.
- Para problemas de importación dinámica (módulos cargados por ruta), PyInstaller puede requerir `hiddenimports` — si al ejecutar el exe aparece un ImportError, añádelo a `hiddenimports` en `app.spec`.

Si quieres, puedo ejecutar el build aquí o ajustar `app.spec` para incluir/excluir archivos concretos según tus preferencias.

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

