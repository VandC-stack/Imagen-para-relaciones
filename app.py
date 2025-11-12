# -- Conversor Excel a JSON (BOSCH) -- #
import os
import json
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# ---------- ESTILO VISUAL CORPORATIVO ---------- #
STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#27AE60",
    "advertencia": "#d57067",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#FFFFFF",
    "texto_oscuro": "#282828",
    "texto_claro": "#4b4b4b",
}

FONT_TITLE = ("Inter", 22, "bold")
FONT_LABEL = ("Inter", 13)
FONT_SMALL = ("Inter", 9)


class BoschExcelToJson(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración general
        self.title("Generador de Dictámenes BOSCH")
        self.geometry("700x500")
        ctk.set_appearance_mode("light")
        self.configure(fg_color=STYLE["fondo"])

        # ===== HEADER =====
        header = ctk.CTkFrame(self, fg_color=STYLE["fondo"], corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Generador de Dictámenes BOSCH",
            font=FONT_TITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=16)

        # ===== CONTENIDO PRINCIPAL =====
        main = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # Botón para subir archivo
        self.upload_button = ctk.CTkButton(
            main,
            text="📁 Cargar Tabla de Relación (Excel)",
            command=self.cargar_excel,
            font=("Inter", 16, "bold"),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=50,
            corner_radius=12
        )
        self.upload_button.pack(pady=(40, 20), fill="x")

        # Etiqueta para mostrar resultados
        self.info_label = ctk.CTkLabel(
            main,
            text="Aún no se ha cargado ningún archivo.",
            font=FONT_LABEL,
            text_color=STYLE["texto_claro"]
        )
        self.info_label.pack(pady=(10, 10))

        # NUEVA etiqueta de estado (repara el error)
        self.status_label = ctk.CTkLabel(
            main,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.status_label.pack(pady=(5, 10))

    # ================== FUNCIONES ================== #
    def cargar_excel(self):
        """Permite seleccionar un archivo Excel y lo convierte a JSON en la carpeta data."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            return

        try:
            self.status_label.configure(text="⏳ Procesando...", text_color=STYLE["advertencia"])
            self.update_idletasks()

            df = pd.read_excel(file_path)
            if df.empty:
                messagebox.showwarning("Archivo vacío", "El archivo seleccionado no contiene datos.")
                return

            # Convertir columnas de fecha a texto
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)

            # Convertir DataFrame a lista de diccionarios
            records = df.to_dict(orient="records")

            # Crear carpeta data
            data_folder = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_folder, exist_ok=True)

            # Guardar JSON
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            json_filename = f"{base_name}_convertido.json"
            output_path = os.path.join(data_folder, json_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(records, f, ensure_ascii=False, indent=2)

            # Mostrar resultados
            self.info_label.configure(text=f"✅ Archivo convertido: {json_filename}")
            self.status_label.configure(text=f"Guardado en: {data_folder}", text_color=STYLE["exito"])

            messagebox.showinfo("Conversión exitosa", f"El archivo fue convertido correctamente:\n\n{output_path}")

        except Exception as e:
            self.status_label.configure(text="❌ Error en la conversión", text_color=STYLE["peligro"])
            messagebox.showerror("Error", f"Ocurrió un error al convertir el archivo:\n\n{e}")


# ================== EJECUCIÓN ================== #
if __name__ == "__main__":
    app = BoschExcelToJson()
    app.mainloop()
