=======================================================================================
PROYECTO PARA EL ARMASDO DE DICTAMENES V&C
=======================================================================================


========================================
📌esta compuesto por varios archivos:
========================================
    1️⃣. app.py
    2️⃣. generador_dictamen
    3️⃣. ArmadoDictamen.py
    4️⃣. DictamenPDF.py
          DictamenDOCX.py (Este se considera en caso de que se requiera mas a delante generar dictamenes en formato WORD).

dichos archivos funcionan de la siguiente manera 

=======================================================================================
📌ORDEN GERARGICO DEL PROYECTO: 
=======================================================================================
    1️⃣. app.py - contien la interfaz y la carga de datos al sistema convirtiendo los archivos en json y almacenandolos en la carpeta data. 
    2️⃣. generador_dictamen -  el archivo app.py se conecta al archivo principal para realizar el armado del dictamen.
    3️⃣. ArmadoDictamen - se conecta a la plantilla DictamenPDF.py.
    4️⃣. DictamentPDF.py - es la plantilla y de donde estan asignadas los PLACEHOLDERS para sustituirlos por informacion que viene en la tabla de relacion que se carga incialmente al sistema. 


        NOTA: para el caso DECATHLON se asiganara un boton para que el usuario suba la base de etiquetado y se generen las etiquetas correctamente en el dictamene teniendo de esta manera 0% de error al generar el dictamen correctamente.

=======================================================================================
📌CARPETAS: 
=======================================================================================


    ==========
    📌img: 
    ==========
        contiene el icono y la imagen de fondo para la plantilla del dictamen en pdf

    ==========
    📌data:
    ==========
        contiene los JSON con los que trabaja el sistema.

        Estos archivos son fijos: 

           1️⃣. Clientes.json
           2️⃣. Normas.json
           3️⃣. Firmas.json ----posteriormente se integrara un nuevo archivo llamado firmas el cual contendra las firmas que se imprimen dentro del dictamen.----


        estos archivos se generan cuando el usuario carga datos para generar dictamenes
           1️⃣. base_etiquetado.json
           2️⃣. tabla_de_relacion.json





