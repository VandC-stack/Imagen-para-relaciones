
# Reglas para el pegado de evidencias y creación de etiquetas

---


## 1. Flujo del pegado de evidencia (modo evidencia)

**¿Qué es?**
Es cuando todas las fotos se colocan juntas en una sección especial del documento, mostrando las evidencias de lo que se hizo o revisó, sin poner una etiqueta individual a cada imagen.

**Flujo paso a paso:**
1. El usuario selecciona las fotos que quiere pegar como evidencia.
2. El sistema agrupa todas las fotos en una carpeta (o varias, según el evento o área).
3. Las fotos se ordenan para que cuenten la historia de la visita o inspección.
4. El sistema pega todas las fotos en una hoja o sección especial del documento, una tras otra.
5. Se puede poner una sola descripción general para todas las fotos, o una breve nota para cada grupo si es necesario.

---
## 2. Flujo del pegado por etiqueta (modo etiqueta)

**¿Qué es?**
Es cuando cada codigo lleva una etiqueta especial, que explica exactamente qué muestra esa imagen.

**¿Cómo arma el sistema la etiqueta?**
El sistema crea cada etiqueta usando dos fuentes:
- El archivo base de etiquetado (base_etiquetado.json), que define el formato y los campos estándar de la etiqueta.
- La tabla de relación (tabla_de_relacion.json), que contiene los datos específicos de cada producto, evento o evidencia.

**Flujo paso a paso:**
1. El sistema lee la tabla de relación y el archivo base de etiquetado.
2. Para cada codigo, busca en la tabla de relación los datos que le corresponden archivo config_etiquetas (por ejemplo: código, marca, descripción, medidas, etc.).
3. Usa el formato del archivo base de etiquetado para armar la etiqueta, llenando los campos con los datos encontrados.
4. La etiqueta final para cada foto puede incluir:
   - Nombre del cliente
   - Fecha de la visita o inspección
   - Número de folio o identificador
   - Descripción breve de la evidencia
   - Código de la evidencia (ejemplo: ETIQUETA 1, ETIQUETA 2...)
   - Otros datos relevantes según el cliente o el tipo de producto
5. El sistema pega cada Etiqueta en el documento junto con su etiqueta, una por una, siguiendo el orden de la tabla de relación.

---

## 3. ¿Cómo se entrelazan el pegado de imagen y la creación de etiquetas?

El sistema decide el modo (evidencia o etiqueta) según las instrucciones del cliente o la configuración interna:

- Si es modo evidencia, todas las fotos se agrupan y se pegan juntas, con una descripción general.
- Si es modo etiqueta, cada codigo se pega con su propia etiqueta.

En ambos casos, el documento solo se considera completo si:
1. Todas las fotos requeridas están pegadas.
2. Todas los codigos tienen su etiqueta (si es modo etiqueta).
3. No falta ninguna evidencia importante.
4. Se sigue el procedimiento estándar.

---

## 4. Criterios para que el documento esté completo

- Todas las fotos necesarias están incluidas.
- No hay documentos sin etiqueta (en modo etiqueta).
- No hay carpetas vacías.
- Si hay imágenes con el mismo nombre en diferentes carpetas, todas se incluyen.
- Si una foto está repetida, solo se pega una vez (a menos que el cliente pida lo contrario).
- Se sigue el orden y agrupación que pide el cliente.

---
