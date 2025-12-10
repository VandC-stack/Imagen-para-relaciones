import os

# Directorio APPDATA local de la aplicación (sin depender de main.py)
APPDATA_DIR = os.path.join(os.getenv("APPDATA"), "ImagenesVC")
os.makedirs(APPDATA_DIR, exist_ok=True)

LOG_FILE = os.path.join(APPDATA_DIR, "documentos_sin_imagenes.txt")


def registrar_fallo(nombre_doc):
    """
    Registra el nombre de un documento que no recibió imágenes en el archivo de log.
    """
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{nombre_doc}\n")
        print(f"Documento agregado al registro de fallos: {nombre_doc}")
    except Exception as e:
        print(f"Error al registrar el fallo de {nombre_doc}: {e}")


def limpiar_registro():
    """
    Borra el archivo de log si existe.
    """
    try:
        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)
            print("Registro de fallos reiniciado correctamente.")
    except Exception as e:
        print(f"Error al limpiar el registro: {e}")


def mostrar_registro():
    """
    Imprime en consola el contenido del archivo de log si existe.
    """
    if not os.path.exists(LOG_FILE):
        print("No hay registro de fallos todavía.")
        return

    print("\n===== DOCUMENTOS SIN IMÁGENES =====")
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            print(f.read())
    except Exception as e:
        print(f"Error al leer el registro de fallos: {e}")
