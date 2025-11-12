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
        self.geometry("480x380")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")
        
        # Header compacto
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(pady=(15, 10), padx=20, fill="x")
        
        title = ctk.CTkLabel(header, text="Imágenes", 
                            font=("Segoe UI", 20, "bold"),
                            text_color="#FFFFFF")
        title.pack(side="left")
        
        version = ctk.CTkLabel(header, text="v3.0", 
                              font=("Segoe UI", 11), 
                              text_color="#4b4b4b")
        version.pack(side="left", padx=(8, 0))
        
        # Configuración de rutas
        config = cargar_config()
        
        # Card de configuración
        config_card = ctk.CTkFrame(self, fg_color="#282828", corner_radius=10)
        config_card.pack(padx=20, pady=10, fill="both", expand=True)
        
        # Documentos
        self._crear_ruta_item(config_card, "📄 Documentos", 
                             config.get("ruta_docs", "No seleccionada"),
                             "documentos", 0)
        
        # Imágenes
        self._crear_ruta_item(config_card, "🖼️ Imágenes", 
                             config.get("ruta_imagenes", "No seleccionada"),
                             "imágenes", 1)
        
        # Botón de acción principal
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
        
        # Footer minimalista
        footer = ctk.CTkLabel(
            self, 
            text="Desarrollado por Enrique Guzmán",
            font=("Segoe UI", 9),
            text_color="#666666"
        )
        footer.pack(side="bottom", pady=(0, 8))
    
    def _crear_ruta_item(self, parent, titulo, ruta, tipo, row):
        # Container para cada ruta
        item = ctk.CTkFrame(parent, fg_color="transparent")
        item.pack(padx=15, pady=(15 if row == 0 else 8, 8), fill="x")
        
        # Label del tipo
        label = ctk.CTkLabel(item, text=titulo, 
                            font=("Segoe UI", 13, "bold"),
                            text_color="#ffffff",
                            anchor="w")
        label.pack(fill="x")
        
        # Ruta actual (truncada si es muy larga)
        ruta_display = self._truncar_ruta(ruta)
        ruta_label = ctk.CTkLabel(item, text=ruta_display,
                                  font=("Segoe UI", 10),
                                  text_color="#b3b3b3",
                                  anchor="w")
        ruta_label.pack(fill="x", pady=(2, 5))
        
        # Botón para cambiar
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
        if len(ruta) <= max_len:
            return ruta
        return "..." + ruta[-(max_len-3):]
    
    def _cambiar_ruta(self, tipo, label):
        carpeta = filedialog.askdirectory(title=f"Selecciona carpeta de {tipo}")
        if carpeta:
            config = cargar_config()
            key = "ruta_docs" if tipo == "documentos" else "ruta_imagenes"
            config[key] = carpeta
            guardar_config(config)
            label.configure(text=self._truncar_ruta(carpeta))
            self.status.configure(text=f"✓ Ruta de {tipo} actualizada", text_color="#4caf50")
    
    def iniciar_proceso(self):
        self.btn_procesar.configure(state="disabled", text="Procesando...")
        self.status.configure(text="⏳ Procesando archivos...", text_color="#ffa726")
        
        def run():
            try:
                procesar_lote()
                self.after(0, self._proceso_exitoso)
            except Exception as e:
                self.after(0, lambda: self._proceso_error(e))
        
        threading.Thread(target=run, daemon=True).start()
    
    def _proceso_exitoso(self):
        self.btn_procesar.configure(state="normal", text="Iniciar Procesamiento")
        self.status.configure(text="✓ Procesamiento completado exitosamente", text_color="#4caf50")
        messagebox.showinfo("Éxito", "Procesamiento completado correctamente")
    
    def _proceso_error(self, error):
        self.btn_procesar.configure(state="normal", text="Iniciar Procesamiento")
        self.status.configure(text="✗ Error en el procesamiento", text_color="#f44336")
        messagebox.showerror("Error", f"Ocurrió un problema:\n{error}")

if __name__ == "__main__":
    app = App()
    app.mainloop()