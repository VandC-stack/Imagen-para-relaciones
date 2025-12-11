import os
import tempfile

# Archivo temporal para esta ejecución
_temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w+", encoding="utf-8")
LOG_FILE = _temp_file.name  # Ruta del archivo temporal

# Esta variable se asigna dinámicamente desde cada módulo
RUTA_DOCS_BASE = None  


def set_base_docs_path(path):
    """
    Se llama desde cada módulo antes de registrar fallos.
    Permite convertir rutas absolutas a rutas relativas.
    """
    global RUTA_DOCS_BASE
    RUTA_DOCS_BASE = path.replace("\\", "/").rstrip("/") + "/"


def ruta_relativa(ruta_completa):
    """
    Convierte una ruta absoluta a ruta relativa respecto a RUTA_DOCS_BASE.
    """
    if not RUTA_DOCS_BASE:
        return os.path.basename(ruta_completa)

    ruta = ruta_completa.replace("\\", "/")
    if ruta.startswith(RUTA_DOCS_BASE):
        return ruta[len(RUTA_DOCS_BASE):]

    return os.path.basename(ruta_completa)


def registrar_fallo(ruta_doc):
    """
    Registra la ruta relativa de un documento que no recibió imágenes.
    Evita duplicados en el archivo temporal.
    """
    try:
        relativa = ruta_relativa(ruta_doc)

        # Evitar duplicados
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            existentes = {line.strip() for line in f.readlines()}
        if relativa in existentes:
            return

        # Registrar el nuevo fallo
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{relativa}\n")

        print(f"Documento agregado al registro de fallos: {relativa}")

    except Exception as e:
        print(f"Error al registrar el fallo de {ruta_doc}: {e}")


def limpiar_registro():
    """
    Reinicia el archivo temporal borrando su contenido.
    """
    try:
        open(LOG_FILE, "w").close()
        print("Registro de fallos reiniciado correctamente.")
    except Exception as e:
        print(f"Error al limpiar el registro: {e}")


def mostrar_registro():
    """
    Muestra el contenido del archivo temporal.
    """
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            contenido = f.read()

        if not contenido.strip():
            print("No hay registro de fallos todavía.")
            return

        print("\n===== DOCUMENTOS SIN IMÁGENES =====")
        print(contenido)

    except Exception as e:
        print(f"Error al leer el registro de fallos: {e}")
