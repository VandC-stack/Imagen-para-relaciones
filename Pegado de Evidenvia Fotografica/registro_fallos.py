import os

# Directorio APPDATA local de la aplicación (sin depender de main.py)
APPDATA_DIR = os.path.join(os.getenv("APPDATA"), "ImagenesVC")
os.makedirs(APPDATA_DIR, exist_ok=True)

LOG_FILE = os.path.join(APPDATA_DIR, "documentos_sin_imagenes.txt")
# JSONL con detalles estructurados: cada línea es un JSON con keys: doc, reason, details, timestamp
LOG_JSON = os.path.join(APPDATA_DIR, "documentos_sin_imagenes.jsonl")

def registrar_fallo(nombre_doc, reason=None, details=None):
    """
    Registra un fallo de pegado para `nombre_doc`.
    - `reason` (opcional): cadena corta explicando la causa (p.ej. 'no_codes', 'not_in_index', 'not_in_path')
    - `details` (opcional): objeto/str con datos adicionales (se serializa como JSON en LOG_JSON).

    Mantiene la compatibilidad hacia atrás: si se llama solo con `nombre_doc` funcionará igual.
    """
    try:
        # Escribir versión legible
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            if reason:
                f.write(f"{nombre_doc} | reason={reason}\n")
            else:
                f.write(f"{nombre_doc}\n")

        # Escribir versión estructurada
        try:
            import json, datetime
            entry = {
                "doc": nombre_doc,
                "reason": reason,
                "details": details,
                "timestamp": datetime.datetime.now().isoformat()
            }
            with open(LOG_JSON, "a", encoding="utf-8") as jf:
                jf.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception:
            pass

        print(f"Documento agregado al registro de fallos: {nombre_doc} (reason={reason})")
    except Exception as e:
        print(f"Error al registrar el fallo de {nombre_doc}: {e}")

def limpiar_registro():
    """
    Borra el archivo de log si existe.
    """
    try:
        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)
        if os.path.exists(LOG_JSON):
            os.remove(LOG_JSON)
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

    print("\n===== DOCUMENTOS SIN IMÁGENES (legible) =====")
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            print(f.read())
    except Exception as e:
        print(f"Error al leer el registro de fallos: {e}")

    # Mostrar registros estructurados (JSONL) si existen
    if os.path.exists(LOG_JSON):
        print("\n===== DOCUMENTOS SIN IMÁGENES (detallado JSON) =====")
        try:
            with open(LOG_JSON, "r", encoding="utf-8") as jf:
                for line in jf:
                    try:
                        print(line.strip())
                    except Exception:
                        print(line)
        except Exception as e:
            print(f"Error al leer el registro JSON de fallos: {e}")
