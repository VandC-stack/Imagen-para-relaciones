import os

LOG_FILE = os.path.abspath("documentos_sin_imagenes.txt")

def registrar_fallo(nombre_doc):
    try:
        with open(LOG_FILE, "a", encoding="utf-8")as f:
            f.write(f"{nombre_doc}\n")
        print(f"Documento agregado al registro de fallos: {nombre_doc}")
    except Exception as e:
        print(f"Error al registrar el fallo de {nombre_doc}: {e}")

def limpiar_registro():
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
        print("Registro de fallos reiniciado correctamente.")

def mostrar_registro():
    if not os.path.exists(LOG_FILE):
        print("No hay registro de fallos todavía.")
        return
    print("\n===== DOCUMENTOS SIN IMÁGENES =====")
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        print(f.read())