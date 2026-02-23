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


# Sistema generador de Dictámenes con Etiquetas Integradas

Bienvenido: este repositorio genera dictámenes en PDF con etiquetas visuales (PNG) integradas. Está pensado para equipos que procesan lotes de productos, aplican normas y requieren la impresión o archivado de dictámenes con sus etiquetas correspondientes.

**Mantenedor:** EFRAIN MORALES ZAMARRON

**Resumen rápido:**
- **Genera** etiquetas PNG a partir de códigos EAN y plantillas de norma.
- **Inserta** dichas etiquetas en la segunda página de los dictámenes PDF.
- **Lee** datos desde la carpeta `data/` (JSON) y permite ejecución por GUI o por script.

## Contenido principal

- **`app.py`**: Interfaz gráfica y orquestador (CustomTkinter).
- **`generador_dictamen.py`**: Lógica principal para procesar familias y crear dictámenes.
- **`etiqueta_dictamen.py`**: Generador de imágenes de etiquetas (Pillow).
- **`plantillaPDF.py`**: Funciones para cargar y preparar datos desde `data/`.
- **`DictamenPDF.py`**: Clase base y utilidades para crear PDFs con ReportLab.
- **`data/`**: JSONs de entrada (tablas, normas, clientes, firmas, folios).
- **`etiquetas_generadas/`**: Salida automática de PNGs.
- **`dictamenes_generados/`**: PDFs resultantes.

## Instalación (rápida)

1. Crear y activar entorno virtual (Windows PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

3. Crear carpetas necesarias si no existen:

```powershell
mkdir data img etiquetas_generadas dictamenes_generados
```

4. Colocar los JSONs y recursos en `data/` y la imagen de fondo en `img/`.

## Uso

- Ejecución desde GUI:

```powershell
python app.py
```

- Ejecución por script (ejemplo):

```python
from generador_dictamen import generar_dictamenes_completos
exito, mensaje, resultado = generar_dictamenes_completos("dictamenes_generados")
```

## Formato y configuración de etiquetas

Las etiquetas se generan según la norma detectada y la configuración en `data/config_etiquetas.json`. Cada norma define tamaño y campos (marca, país, talla, composición, etc.). Las imágenes se guardan en `etiquetas_generadas/` y se insertan en la segunda página del PDF.

Ejemplo de entrada en `config_etiquetas.json`:

```json
{
  "NOM-024-SCFI-2013": {
    "tamaño_cm": "(5.0, 5.0)",
    "campos": ["MARCA", "PAIS ORIGEN", "TALLA", "COMPOSICION"]
  }
}
```

## Flujo de trabajo interno

1. Cargar datos: `data/tabla_de_relacion.json`, `data/Normas.json`, `data/Clientes.json`, `data/Firmas.json`.
2. Agrupar registros por familia/norma/folio para procesar lotes.
3. Para cada código EAN buscar la definición en `BASE_ETIQUETADO.json` y generar PNG.
4. Construir PDF: página 1 (dictamen), página 2 (etiquetas e imágenes), insertar firmas y fondo.

## Empaquetado a .exe (Windows)

Se incluye `build_exe.bat` y `Sistema_Generador_Documentos_VC.spec` para PyInstaller.

Pasos básicos:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\build_exe.bat
```

Nota: si usas archivos Excel `.xlsb` instala `pyxlsb` en el entorno destino y añade `hiddenimports` si PyInstaller reporta ImportError.

## Solución de problemas comunes

- "No se generaron etiquetas": verificar que los EAN estén en `BASE_ETIQUETADO.json` y que `TABLA_DE_RELACION.json` use los mismos códigos.
- "Imágenes no aparecen en el PDF": comprobar que `etiquetas_generadas/` contiene los PNG y que las rutas relativas en el proceso de inserción son correctas.
- Error al cargar normas: validar formato de `data/Normas.json` (campos `NOM`, `NOMBRE`, `CAPITULO`).

## Desarrollo y pruebas

- Ejecutar funciones directamente para pruebas unitarias: `plantillaPDF.cargar_tabla_relacion()` o `generador_dictamen.generar_dictamenes_completos(...)` con muestras en `data/`.
- Mantener respaldos automáticos: antes de editar `data/tabla_de_relacion.json` el sistema crea copias en `data/tabla_relacion_backups/`.

## Cómo contribuir o extender

- Añadir una nueva norma: editar `data/Normas.json` y `data/config_etiquetas.json`; si la norma requiere lógica especial, extender `etiqueta_dictamen.py::crear_mapeo_norma_uva`.
- Para agregar recursos al empaquetado con PyInstaller, editar `Sistema_Generador_Documentos_VC.spec` y añadir rutas a `datas`.

---

Si quieres, puedo:
- Ejecutar una generación de prueba con datos de ejemplo.
- Ajustar o ampliar este README con instrucciones paso a paso más detalladas.

Contacto del mantenedor: EFRAIN MORALES ZAMARRON
## 🤝 Contribuciones

