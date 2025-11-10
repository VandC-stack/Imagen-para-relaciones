import os
from datetime import datetime

# Archivo donde se guardará el registro
LOG_FILE = os.path.abspath("documentos_sin_imagenes.txt")

def registrar_fallo(nombre_doc):
    """
    Registra en un archivo de texto el nombre de un documento .docx
    al cual no se le pudo insertar al menos una imagen.
    """
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{nombre_doc}\n")
        print(f"Documento agregado al registro de fallos: {nombre_doc}")
    except Exception as e:
        print(f"Error al registrar el fallo de {nombre_doc}: {e}")

def limpiar_registro():
    """Elimina el contenido previo del archivo de registro."""
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
        print("Registro de fallos reiniciado correctamente.")

def mostrar_registro():
    """Imprime en consola el contenido actual del registro."""
    if not os.path.exists(LOG_FILE):
        print("No hay registro de fallos todavía.")
        return
    print("\n===== DOCUMENTOS SIN IMÁGENES =====")
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        print(f.read())
