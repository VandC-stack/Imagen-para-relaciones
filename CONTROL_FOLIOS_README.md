# Control de Folios Anual - Generador de Excel

Este script genera un archivo Excel con el control de folios anual a partir de los datos almacenados en archivos JSON del sistema de dictámenes.

## 📋 Descripción

El script `control_folios_anual.py` lee información de múltiples fuentes de datos JSON y genera un archivo Excel estructurado con toda la información necesaria para el control anual de folios.

## 🎯 Características

- **Generación automática**: Crea un archivo Excel con formato profesional
- **Filtrado por fechas**: Permite generar reportes por rangos de fechas específicos
- **Agrupación por dictamen**: Agrupa los registros por solicitud y folio
- **Información completa**: Incluye 19 columnas con toda la información requerida
- **Validación de datos**: Manejo robusto de valores faltantes o incorrectos
- **Formato profesional**: Encabezados con estilo, bordes y colores

## 📊 Columnas del Excel Generado

El archivo Excel contiene las siguientes 19 columnas:

1. **NÚMERO DE SOLICITUD** - Código de identificación del dictamen
2. **CLIENTE** - Nombre del cliente
3. **NÚMERO DE CONTRATO** - Número de contrato asociado
4. **RFC** - RFC del cliente
5. **CURP** - CURP (valor por defecto: N/A)
6. **PRODUCTO VERIFICADO** - Descripción de los productos
7. **MARCAS** - Marcas de los productos
8. **NOM** - Clasificación UVA (NOM)
9. **TIPO DE DOCUMENTO OFICIAL EMITIDO** - Siempre "D" (Dictamen)
10. **DOCUMENTO EMITIDO** - Número de solicitud
11. **FECHA DE DOCUMENTO EMITIDO** - Fecha de emisión
12. **VERIFICADOR** - Nombre del inspector
13. **PEDIMENTO DE IMPORTACION** - Número de pedimento
14. **FECHA DE DESADUANAMIENTO (CUANDO APLIQUE)** - Fecha de entrada
15. **FECHA DE VISITA (CUANDO APLIQUE)** - Fecha de verificación
16. **MODELOS** - Lista de códigos de modelos
17. **SOL EMA** - Últimos valores del número de solicitud
18. **FOLIO EMA** - Folio formateado a 6 dígitos
19. **INSP EMA** - Nombre completo del inspector

## 🗂️ Fuentes de Datos

El script lee información de los siguientes archivos JSON ubicados en el directorio `data/`:

1. **tabla_de_relacion.json** - Tabla principal con información de productos y folios
2. **Clientes.json** - Información de clientes (RFC, NÚMERO_DE_CONTRATO, nombre)
3. **Firmas.json** - Información de inspectores (FIRMA, NOMBRE DE INSPECTOR)

## 🚀 Uso

### Instalación de Dependencias

```bash
pip install openpyxl
```

### Uso Básico

Generar el control completo de folios (todos los registros):

```bash
python control_folios_anual.py
```

Esto generará un archivo llamado `Control_Folios_Anual.xlsx` en el directorio actual.

### Opciones Avanzadas

#### Especificar nombre del archivo de salida

```bash
python control_folios_anual.py --output Mi_Control_2024.xlsx
```

o usando la forma corta:

```bash
python control_folios_anual.py -o Mi_Control_2024.xlsx
```

#### Filtrar por rango de fechas

Generar reporte solo para noviembre 2025:

```bash
python control_folios_anual.py --fecha-inicio 2025-11-01 --fecha-fin 2025-11-30
```

o usando las formas cortas:

```bash
python control_folios_anual.py -fi 2025-11-01 -ff 2025-11-30
```

#### Especificar directorio de datos alternativo

```bash
python control_folios_anual.py --data-dir /ruta/a/datos
```

o:

```bash
python control_folios_anual.py -d /ruta/a/datos
```

#### Combinando opciones

```bash
python control_folios_anual.py \
  -o Control_Noviembre_2025.xlsx \
  -fi 2025-11-01 \
  -ff 2025-11-30 \
  -d data
```

### Ver ayuda

```bash
python control_folios_anual.py --help
```

## 📝 Ejemplos de Uso

### Ejemplo 1: Generar reporte anual completo

```bash
python control_folios_anual.py -o Control_Folios_2025.xlsx
```

### Ejemplo 2: Reporte mensual (Diciembre 2025)

```bash
python control_folios_anual.py \
  -o Control_Diciembre_2025.xlsx \
  -fi 2025-12-01 \
  -ff 2025-12-31
```

### Ejemplo 3: Reporte trimestral (Q4 2025)

```bash
python control_folios_anual.py \
  -o Control_Q4_2025.xlsx \
  -fi 2025-10-01 \
  -ff 2025-12-31
```

### Ejemplo 4: Reporte de un rango personalizado

```bash
python control_folios_anual.py \
  -o Control_Nov_15_a_Dic_15.xlsx \
  -fi 2025-11-15 \
  -ff 2025-12-15
```

## 🔧 Detalles Técnicos

### Agrupación de Datos

El script agrupa los registros por:
- **Número de Solicitud** (SOLICITUD)
- **Folio** (FOLIO)

Cada combinación única de solicitud y folio genera una fila en el Excel.

### Manejo de Múltiples Registros

Cuando un dictamen tiene múltiples registros (varios productos), el script:
- Combina las descripciones de productos separadas por comas
- Combina las marcas únicas
- Combina las clasificaciones NOM
- Lista todos los códigos de modelos

### Formato de Fechas

Las fechas se procesan en múltiples formatos:
- YYYY-MM-DD (recomendado)
- YYYY/MM/DD
- DD/MM/YYYY
- DD-MM-YYYY

### Formato del Excel

- **Encabezados**: Fondo azul (#366092), texto blanco, negrita
- **Celdas**: Bordes en todas las celdas, texto ajustado
- **Columnas**: Ancho automático según contenido
- **Primera fila**: Congelada para facilitar la navegación

## ⚠️ Manejo de Errores

El script maneja los siguientes casos:

1. **Archivos JSON faltantes**: Muestra mensaje de error específico
2. **JSON inválido**: Captura errores de decodificación
3. **Valores faltantes**: Reemplaza con "N/A"
4. **Fechas inválidas**: Incluye el registro sin filtrar
5. **Firmas no encontradas**: Retorna "N/A"

## 📊 Salida de Ejemplo

```
======================================================================
📊 GENERADOR DE CONTROL DE FOLIOS ANUAL
======================================================================

📂 Cargando datos desde archivos JSON...
✅ Clientes cargados: 99 registros
✅ Firmas cargadas: 18 registros
✅ Tabla de relación cargada: 224 registros

🚀 Generando archivo Excel...
📊 Dictámenes encontrados: 95
✅ Archivo Excel generado exitosamente: Control_Folios_Anual.xlsx
   📊 Total de registros: 95
   📅 Rango de fechas aplicado: 2025-11-01 a 2025-11-30

======================================================================
✅ PROCESO COMPLETADO
======================================================================
```

## 🐛 Solución de Problemas

### Error: "No se encontró Clientes.json"

**Causa**: El archivo no existe en el directorio `data/`

**Solución**: Verifica que el archivo exista o especifica un directorio diferente con `--data-dir`

### Error: "Error al decodificar JSON"

**Causa**: Uno de los archivos JSON está mal formateado

**Solución**: Verifica la sintaxis del JSON en el archivo indicado

### El Excel no contiene datos

**Causa**: El rango de fechas no coincide con ningún registro

**Solución**: 
- Verifica que las fechas estén en formato YYYY-MM-DD
- Comprueba que existan registros en el rango especificado
- Prueba sin filtro de fechas primero

### Columnas muy estrechas/anchas

**Causa**: El script usa anchos predefinidos

**Solución**: Después de generar el archivo, puedes ajustar manualmente las columnas en Excel o modificar el script en la sección de ajuste de anchos de columnas

## 🔄 Integración con el Sistema

Este script está diseñado para trabajar con la estructura de datos existente del sistema de generación de dictámenes. Los archivos JSON son generados y mantenidos por la aplicación principal (`app.py`).

## 📚 Referencias

- [Documentación de openpyxl](https://openpyxl.readthedocs.io/)
- Sistema de Generación de Dictámenes - README.md principal

## 🤝 Contribuciones

Para agregar nuevas funcionalidades o modificar el script:

1. **Agregar nuevas columnas**: Modifica la lista `encabezados` y el método `generar_fila_excel()`
2. **Cambiar formato**: Modifica la sección de estilos en `crear_excel()`
3. **Agregar validaciones**: Extiende el método `filtrar_por_fechas()` o agrega nuevos métodos

## 📞 Soporte

Si encuentras problemas o necesitas nuevas características:

1. Verifica que todos los archivos JSON existan y sean válidos
2. Confirma que las dependencias estén instaladas (`pip install openpyxl`)
3. Revisa los mensajes de error para identificar el problema específico
4. Ejecuta el script con `--help` para ver todas las opciones disponibles
