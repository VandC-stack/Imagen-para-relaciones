import customtkinter as ctk
from tkinter import messagebox, filedialog
import threading
from main import procesar_lote, cargar_config, guardar_config


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Imágenes")
        self.geometry("520x480")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(pady=(15, 10), padx=20, fill="x")

        title = ctk.CTkLabel(
            header,
            text="Imágenes",
            font=("Segoe UI", 20, "bold"),
            text_color="#FFFFFF"
        )
        title.pack(side="left")

        version = ctk.CTkLabel(
            header,
            text="v03.9",
            font=("Segoe UI", 11),
            text_color="#4b4b4b"
        )
        version.pack(side="left", padx=(8, 0))

        config = cargar_config()

        config_card = ctk.CTkFrame(self, fg_color="#282828", corner_radius=10)
        config_card.pack(padx=20, pady=10, fill="both", expand=True)

        self._crear_ruta_item(
            config_card,
            "Documentos",
            config.get("ruta_docs", "No seleccionada"),
            "documentos"
        )

        self._crear_ruta_item(
            config_card,
            "Imágenes",
            config.get("ruta_imagenes", "No seleccionada"),
            "imágenes"
        )

        modo_label = ctk.CTkLabel(
            config_card,
            text="Modo de pegado:",
            font=("Segoe UI", 13, "bold"),
            text_color="#FFFFFF"
        )
        modo_label.pack(anchor="w", padx=15, pady=(12, 4))

        self.modo_var = ctk.StringVar(value=config.get("modo_pegado", "simple"))
        botones_frame = ctk.CTkFrame(config_card, fg_color="transparent")
        botones_frame.pack(anchor="w", padx=15, pady=(5, 10))

        def seleccionar_modo(modo):
            self.modo_var.set(modo)
            config = cargar_config()
            config["modo_pegado"] = modo
            guardar_config(config)
            actualizar_colores()

        self.btn_simple = ctk.CTkButton(
            botones_frame,
            text="Carpeta única",
            width=130,
            height=34,
            fg_color="#4b4b4b",
            hover_color="#6c6c6c",
            text_color="#FFFFFF",
            command=lambda: seleccionar_modo("simple")
        )
        self.btn_simple.pack(side="left", padx=4)

        self.btn_carpetas = ctk.CTkButton(
            botones_frame,
            text="Múltiples carpetas",
            width=150,
            height=34,
            fg_color="#4b4b4b",
            hover_color="#6c6c6c",
            text_color="#FFFFFF",
            command=lambda: seleccionar_modo("carpetas")
        )
        self.btn_carpetas.pack(side="left", padx=4)

        self.btn_indice = ctk.CTkButton(
            botones_frame,
            text="Pegado por Índice",
            width=150,
            height=34,
            fg_color="#4b4b4b",
            hover_color="#6c6c6c",
            text_color="#FFFFFF",
            command=lambda: seleccionar_modo("indice")
        )
        self.btn_indice.pack(side="left", padx=4)

        def actualizar_colores():
            botones = {
                "simple": self.btn_simple,
                "carpetas": self.btn_carpetas,
                "indice": self.btn_indice
            }
            for m, btn in botones.items():
                if self.modo_var.get() == m:
                    btn.configure(fg_color="#ECD925", text_color="#000000", hover_color="#B8AA00")
                else:
                    btn.configure(fg_color="#4b4b4b", text_color="#FFFFFF", hover_color="#6c6c6c")

        actualizar_colores()

        self.btn_procesar = ctk.CTkButton(
            self,
            text="Iniciar Procesamiento",
            font=("Segoe UI", 16, "bold"),
            height=48,
            fg_color="#ECD925",
            hover_color="#B8AA00",
            text_color="#000000",
            corner_radius=8,
            command=self.iniciar_proceso
        )
        self.btn_procesar.pack(padx=20, pady=15, fill="x")

        self.status = ctk.CTkLabel(
            self,
            text="Listo para procesar",
            font=("Segoe UI", 10),
            text_color="#8c8c8c"
        )
        self.status.pack(pady=(0, 10))

        footer = ctk.CTkLabel(
            self,
            text="Desarrollado por Enrique Guzmán",
            font=("Segoe UI", 9),
            text_color="#666666"
        )
        footer.pack(side="bottom", pady=(0, 8))

    def _crear_ruta_item(self, parent, titulo, ruta, tipo):
        item = ctk.CTkFrame(parent, fg_color="transparent")
        item.pack(padx=15, pady=(14, 2), fill="x")

        label = ctk.CTkLabel(
            item,
            text=titulo,
            font=("Segoe UI", 13, "bold"),
            text_color="#ffffff",
            anchor="w"
        )
        label.pack(fill="x")

        ruta_label = ctk.CTkLabel(
            item,
            text=self._truncar_ruta(ruta),
            font=("Segoe UI", 10),
            text_color="#b3b3b3",
            anchor="w"
        )
        ruta_label.pack(fill="x", pady=(2, 5))

        btn = ctk.CTkButton(
            item,
            text="Cambiar carpeta",
            height=28,
            font=("Segoe UI", 11),
            fg_color="#4b4b4b",
            hover_color="#6c6c6c",
            text_color="#ffffff",
            corner_radius=6,
            command=lambda: self._cambiar_ruta(tipo, ruta_label)
        )
        btn.pack(fill="x")

    def _truncar_ruta(self, ruta, max_len=58):
        return ruta if len(ruta) <= max_len else "..." + ruta[-(max_len - 3):]

    def _cambiar_ruta(self, tipo, label):
        carpeta = filedialog.askdirectory(title=f"Selecciona carpeta de {tipo}")
        if carpeta:
            config = cargar_config()
            key = "ruta_docs" if tipo == "documentos" else "ruta_imagenes"
            config[key] = carpeta
            guardar_config(config)
            label.configure(text=self._truncar_ruta(carpeta))

    def iniciar_proceso(self):
        self.btn_procesar.configure(state="disabled", text="Procesando...")
        self.status.configure(text="Procesando archivos...", text_color="#ffa726")

        def run():
            try:
                procesar_lote()
                self.after(0, self._proceso_exitoso)
            except Exception as err:
                self.after(0, lambda err=err: self._proceso_error(err))

        threading.Thread(target=run, daemon=True).start()

    def _proceso_exitoso(self):
        self.btn_procesar.configure(state="normal", text="Iniciar Procesamiento")
        self.status.configure(text="Procesamiento completado exitosamente", text_color="#4caf50")
        messagebox.showinfo("Éxito", "Procesamiento completado correctamente")

    def _proceso_error(self, error):
        self.btn_procesar.configure(state="normal", text="Iniciar Procesamiento")
        self.status.configure(text="Error en el procesamiento", text_color="#f44336")
        messagebox.showerror("Error", f"Ocurrió un problema:\n{error}")


if __name__ == "__main__":
    # Firma interna – NO visible para el usuario final
    app = App()
    app.mainloop()  