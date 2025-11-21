# -- convertidorjson.py --
# Conversor de archivos Excel a JSON (para carpeta /data)
# Para archivos Normas, listado de clientes Firmas de inspectores

import os
import json
import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox, Tk


# CONFIGURACIÓN
DATA_DIR = os.path.join(os.getcwd(), "data")
os.makedirs(DATA_DIR, exist_ok=True)  # Crea carpeta /data si no existe


# FUNCIONES
def convertir_excel_a_json(file_path: str) -> str:
    """
    Convierte un archivo Excel (.xlsx o .xls) a JSON
    y lo guarda en la carpeta /data del proyecto.
    
    Retorna la ruta completa del archivo JSON generado.
    """
    try:
        # Leer Excel
        df = pd.read_excel(file_path)
        if df.empty:
            raise ValueError("El archivo Excel está vacío o sin datos.")

        # Convertir a lista de diccionarios
        records = df.to_dict(orient="records")

        # Nombre base del archivo
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d")
        json_filename = f"{base_name}_{timestamp}.json"

        # Ruta de salida
        output_path = os.path.join(DATA_DIR, json_filename)

        # Función para serializar objetos Timestamp
        def convertir_timestamp(obj):
            if isinstance(obj, pd.Timestamp):
                return obj.isoformat()
            raise TypeError(f"Tipo {type(obj)} no serializable")

        # Guardar JSON con formato legible
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(records, f, 
                     ensure_ascii=False, 
                     indent=2,
                     default=convertir_timestamp)  # Usar conversor personalizado

        print(f"✅ Archivo convertido y guardado en: {output_path}")
        return output_path

    except Exception as e:
        raise RuntimeError(f"Error al convertir Excel a JSON: {e}")


def abrir_y_convertir():
    """
    Abre un diálogo para seleccionar un archivo Excel y lo convierte automáticamente a JSON.
    Guarda el resultado en /data.
    """
    root = Tk()
    root.withdraw()  # Oculta ventana principal
    try:
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel para convertir a JSON",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            print("Operación cancelada.")
            return None

        json_path = convertir_excel_a_json(file_path)
        messagebox.showinfo("Conversión completada", f"JSON generado:\n{json_path}")
        return json_path

    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        root.destroy()

# EJECUCIÓN DIRECTA
if __name__ == "__main__":
    print("=== CONVERTIDOR JSON ===")
    print("Seleccione un archivo Excel para convertirlo en JSON...")
    abrir_y_convertir()