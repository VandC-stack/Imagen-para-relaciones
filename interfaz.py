import customtkinter as ctk
from tkinter import messagebox, filedialog
import threading
from main import procesar_lote, cargar_config, guardar_config

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# =========================================================
# UTILIDAD: INTERPOLACIÓN SUAVE DE COLOR
# =========================================================
def _hex_to_rgb(h):
    h = h.replace("#", "")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def _rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)

def _interpolate_color(c1, c2, t):
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Imágenes")
        self.geometry("480x460")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")
        
        # Header
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
            text="v3.1", 
            font=("Segoe UI", 11), 
            text_color="#4b4b4b"
        )
        version.pack(side="left", padx=(8, 0))
        
        # Configuración inicial
        config = cargar_config()
        
        # Card de configuración
        config_card = ctk.CTkFrame(self, fg_color="#282828", corner_radius=10)
        config_card.pack(padx=20, pady=10, fill="both", expand=True)
        
        # Documentos
        self._crear_ruta_item(
            config_card,
            "Documentos",
            config.get("ruta_docs", "No seleccionada"),
            "documentos"
        )
        
        # Imágenes
        self._crear_ruta_item(
            config_card,
            "Imágenes",
            config.get("ruta_imagenes", "No seleccionada"),
            "imágenes"
        )

        # Selector de modo
        modo_label = ctk.CTkLabel(
            config_card,
            text="Modo de pegado:",
            font=("Segoe UI", 13, "bold"),
            text_color="#FFFFFF"
        )
        modo_label.pack(anchor="w", padx=15, pady=(10, 2))

        self.modo_var = ctk.StringVar(value=config.get("modo_pegado", "simple"))

        # Frame contenedor del switch y texto
        modo_frame = ctk.CTkFrame(config_card, fg_color="transparent")
        modo_frame.pack(anchor="w", padx=15, pady=(2, 10), fill="x")

        # Switch con colores base
        self.switch = ctk.CTkSwitch(
            modo_frame,
            text="",
            variable=self.modo_var,
            onvalue="carpetas",
            offvalue="simple",
            command=self._actualizar_modo,
            fg_color="#4b4b4b",
            progress_color="#4b4b4b",
            button_color="#2b2b2b",
            button_hover_color="#3c3c3c"
        )
        self.switch.pack(side="left")

        # Texto dinámico junto al switch
        self.modo_texto = ctk.CTkLabel(
            modo_frame,
            text="Carpeta única" if self.modo_var.get() == "simple" else "Multiples carpetas",
            font=("Segoe UI", 11),
            text_color="#FFFFFF"
        )
        self.modo_texto.pack(side="left", padx=10)

        # Ajuste inicial de colores según config
        if self.modo_var.get() == "carpetas":
            self.switch.configure(progress_color="#B8AA00", button_color="#ECD925")
        else:
            self.switch.configure(progress_color="#4b4b4b", button_color="#2b2b2b")

        # Botón procesar
        self.btn_procesar = ctk.CTkButton(
            self,
            text="Iniciar Procesamiento",
            font=("Segoe UI", 16, "bold"),
            height=45,
            fg_color="#ECD925",
            hover_color="#B8AA00",
            text_color="#000000",
            corner_radius=8,
            command=self.iniciar_proceso
        )
        self.btn_procesar.pack(padx=20, pady=15, fill="x")
        
        # Status bar
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


    # =========================================================
    # ANIMACIÓN SUAVE DEL SWITCH
    # =========================================================
    def _animate_switch(self, start_hex, end_hex, start_btn, end_btn, steps=12):
        start_rgb = _hex_to_rgb(start_hex)
        end_rgb   = _hex_to_rgb(end_hex)

        start_rb2 = _hex_to_rgb(start_btn)
        end_rb2   = _hex_to_rgb(end_btn)

        def step(i=0):
            t = i / steps
            color1 = _rgb_to_hex(_interpolate_color(start_rgb, end_rgb, t))
            color2 = _rgb_to_hex(_interpolate_color(start_rb2, end_rb2, t))

            self.switch.configure(progress_color=color1, button_color=color2)

            if i < steps:
                self.after(25, lambda: step(i+1))

        step()


    # =========================================================
    # CAMBIO DE MODO + ANIMACIÓN
    # =========================================================
    def _actualizar_modo(self):
        config = cargar_config()
        config["modo_pegado"] = self.modo_var.get()
        guardar_config(config)

        if self.modo_var.get() == "carpetas":
            self.modo_texto.configure(text="Multiples carpetas")

            # Animación hacia AMARILLO
            self._animate_switch(
                start_hex=self.switch.cget("progress_color"),
                end_hex="#ECD925",
                start_btn=self.switch.cget("button_color"),
                end_btn="#B8AA00"
            )

            # Hover azul solo cuando está encendido
            self.switch.configure(button_hover_color="#585400")

        else:
            self.modo_texto.configure(text="Carpeta única")

            # Animación hacia GRIS
            self._animate_switch(
                start_hex=self.switch.cget("progress_color"),
                end_hex="#4b4b4b",
                start_btn=self.switch.cget("button_color"),
                end_btn="#2b2b2b"
            )

            # Hover gris cuando está apagado
            self.switch.configure(button_hover_color="#3c3c3c")


    # =========================================================
    # RESTO DEL SISTEMA
    # =========================================================
    def _crear_ruta_item(self, parent, titulo, ruta, tipo):
        item = ctk.CTkFrame(parent, fg_color="transparent")
        item.pack(padx=15, pady=(15, 5), fill="x")
        
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
    
    def _truncar_ruta(self, ruta, max_len=50):
        return ruta if len(ruta) <= max_len else "..." + ruta[-(max_len-3):]
    
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
            except Exception as e:
                self.after(0, lambda: self._proceso_error(e))
        
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
    app = App()
    app.mainloop()