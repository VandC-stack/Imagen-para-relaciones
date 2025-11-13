# -- SISTEMA V&C - GENERADOR DE DICTÁMENES -- #
import os, sys
import json
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import subprocess
from datetime import datetime
import time
import platform

# ---------- ESTILO VISUAL V&C ---------- #
STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008D53",
    "advertencia": "#d57067",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#FFFFFF",
    "texto_oscuro": "#282828",
    "texto_claro": "#4b4b4b",
    "borde": "#DDDDDD"
}

FONT_TITLE = ("Inter", 17, "bold")
FONT_SUBTITLE = ("Inter", 14, "bold")
FONT_LABEL = ("Inter", 13)
FONT_SMALL = ("Inter", 12)


class SistemaDictamenesVC(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración general
        self.title("Generador de Dictámenes")
        self.geometry("460x500")
        self.minsize(460, 600)  # ⛔ Evita que la ventana sea más pequeña
        ctk.set_appearance_mode("light")
        self.configure(fg_color=STYLE["fondo"])

        # Variables de estado
        self.archivo_excel_cargado = None
        self.archivo_json_generado = None
        self.json_filename = None
        self.generando_dictamenes = False

        # ===== HEADER =====
        self.crear_header()

        # ===== CONTENIDO PRINCIPAL =====
        self.crear_contenido_principal()

        # ===== FOOTER =====
        self.crear_footer()

    def centerwindow(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        ancho_ventana = self.winfo_width()
        alto_ventana = self.winfo_height()
        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        self.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

    # -----------------------------------------------------------
    # SECCIONES VISUALES
    # -----------------------------------------------------------

    def crear_header(self):
        """Header mejorado con diseño más profesional"""
        header = ctk.CTkFrame(self, fg_color=STYLE["fondo"], corner_radius=0, height=50)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        # Contenedor principal del header
        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(expand=True, fill="both", padx=25, pady=15)

        # Título principal
        ctk.CTkLabel(
            header_content,
            text="Generador de Dictámenes",
            font=FONT_TITLE,
            text_color=STYLE["secundario"]
        ).pack(anchor="center", expand=True, fill="both", pady=(0, 5))

    def crear_contenido_principal(self):
        """Contenido principal reorganizado y mejorado"""
        main_container = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
        main_container.pack(fill="both", expand=True, padx=25, pady=20)

        # ===== TARJETA DE CARGA =====
        card_carga = ctk.CTkFrame(main_container, fg_color=STYLE["surface"], corner_radius=12)
        card_carga.pack(fill="x", pady=(0, 20))

        # Encabezado de la tarjeta
        ctk.CTkLabel(
            card_carga,
            text="📊 Cargar Tabla de Relación",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 5))

        # Información del archivo
        self.info_archivo = ctk.CTkLabel(
            card_carga,
            text="No se ha cargado ningún archivo",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.info_archivo.pack(anchor="w", padx=20, pady=(0, 15))

        # Botones de acción
        botones_frame = ctk.CTkFrame(card_carga, fg_color="transparent")
        botones_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.boton_cargar_excel = ctk.CTkButton(
            botones_frame,
            text="Subir archivo",
            command=self.cargar_excel,
            font=("Inter", 14, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=40,
            width=120,
            corner_radius=8
        )
        self.boton_cargar_excel.pack(side="left", padx=(0, 10))

        self.boton_limpiar = ctk.CTkButton(
            botones_frame,
            text="Limpiar",
            command=self.limpiar_archivo,
            font=("Inter", 14),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=40,
            width=70,
            corner_radius=8,
            state="disabled"
        )
        self.boton_limpiar.pack(side="left")

        # Estado de conversión
        estado_frame = ctk.CTkFrame(card_carga, fg_color="transparent")
        estado_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.etiqueta_estado = ctk.CTkLabel(
            estado_frame,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.etiqueta_estado.pack(side="left")

        self.check_label = ctk.CTkLabel(
            estado_frame,
            text="",
            font=("Inter", 16, "bold"),
            text_color=STYLE["exito"]
        )
        self.check_label.pack(side="right")

        # ===== TARJETA DE GENERACIÓN =====
        card_generacion = ctk.CTkFrame(main_container, fg_color=STYLE["surface"], corner_radius=12)
        card_generacion.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            card_generacion,
            text="🧾 Generar Dictámenes",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 5))

        # Información del archivo
        self.info_generacion = ctk.CTkLabel(
            card_generacion,
            text="Se generan dictámenes en formato pdf",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.info_generacion.pack(anchor="w", padx=20, pady=(0, 10))

        # Barra de progreso para generación de dictámenes (color primario #ECD925)
        self.barra_progreso = ctk.CTkProgressBar(
            card_generacion,
            progress_color=STYLE["primario"],  # Cambiado a color primario
            height=12,
            width=400,
            corner_radius=6
        )
        self.barra_progreso.pack(padx=20, pady=(10, 10))
        self.barra_progreso.set(0)

        # Etiqueta de progreso
        self.etiqueta_progreso = ctk.CTkLabel(
            card_generacion,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.etiqueta_progreso.pack(padx=20, pady=(0, 10))

        # Botón de generación
        self.boton_generar_dictamen = ctk.CTkButton(
            card_generacion,
            text="Generar Dictámenes",
            command=self.generar_dictamenes,
            font=("Inter", 15, "bold"),
            fg_color=STYLE["exito"],
            hover_color="#1f8c4d",
            text_color=STYLE["surface"],
            height=40,
            width=120,
            corner_radius=8,
            state="disabled"
        )
        self.boton_generar_dictamen.pack(padx=20, pady=(0, 20))

    def crear_footer(self):
        """Footer mejorado"""
        footer = ctk.CTkFrame(self, fg_color=STYLE["fondo"], corner_radius=0, height=40)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(expand=True, fill="both", padx=25, pady=10)

        ctk.CTkLabel(
            footer_content,
            text="Sistema V&C - Generador de Dictámenes de Cumplimiento",
            font=("Inter", 10),
            text_color=STYLE["secundario"]
        ).pack(side="left")

    # -----------------------------------------------------------
    # FUNCIONALIDAD PRINCIPAL
    # -----------------------------------------------------------

    def cargar_excel(self):
        """Selecciona el Excel y lo convierte automáticamente a JSON"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            return

        self.archivo_excel_cargado = file_path
        nombre_archivo = os.path.basename(file_path)
        
        self.info_archivo.configure(
            text=f"📄 {nombre_archivo}",
            text_color=STYLE["exito"]
        )
        
        self.boton_limpiar.configure(state="normal")
        self.etiqueta_estado.configure(
            text="⏳ Convirtiendo a JSON...", 
            text_color=STYLE["advertencia"]
        )
        self.check_label.configure(text="")
        self.update_idletasks()

        # Procesar conversión en segundo plano
        thread = threading.Thread(target=self.convertir_a_json, args=(file_path,))
        thread.daemon = True
        thread.start()

    def convertir_a_json(self, file_path):
        """Convierte el Excel a JSON directamente"""
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                self.mostrar_error("El archivo seleccionado no contiene datos.")
                return

            # Convertir fechas a texto
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)

            records = df.to_dict(orient="records")

            # Guardar JSON con nombre fijo
            data_folder = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_folder, exist_ok=True)

            # 🔹 Nombre fijo del archivo JSON
            self.json_filename = "tabla_de_relacion.json"
            output_path = os.path.join(data_folder, self.json_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(records, f, ensure_ascii=False, indent=2)

            # Actualizar UI en el hilo principal
            self.after(0, self._actualizar_ui_conversion_exitosa, output_path, len(records))

        except Exception as e:
            self.after(0, self.mostrar_error, f"Error al convertir el archivo:\n{e}")

    def _actualizar_ui_conversion_exitosa(self, output_path, num_registros):
        """Actualiza la UI cuando la conversión es exitosa"""
        self.archivo_json_generado = output_path
        self.etiqueta_estado.configure(
            text=f"✅ Convertido - {num_registros} registros", 
            text_color=STYLE["exito"]
        )
        self.check_label.configure(text="✓")
        self.boton_generar_dictamen.configure(state="normal")
        
        messagebox.showinfo(
            "Conversión exitosa",
            f"Archivo convertido correctamente.\n\n"
            f"Ubicación: {output_path}\n"
            f"Total de registros: {num_registros}"
        )

    def limpiar_archivo(self):
        """Limpia el estado actual y elimina archivos generados"""
        try:
            # Eliminar archivo JSON si existe
            if self.json_filename:
                data_folder = os.path.join(os.path.dirname(__file__), "data")
                json_path = os.path.join(data_folder, self.json_filename)
                if os.path.exists(json_path):
                    os.remove(json_path)
                    print(f"Archivo eliminado: {json_path}")
        except Exception as e:
            print(f"Error al eliminar archivo: {e}")

        # Resetear estado
        self.archivo_excel_cargado = None
        self.archivo_json_generado = None
        self.json_filename = None
        
        # Resetear UI
        self.info_archivo.configure(
            text="No se ha cargado ningún archivo", 
            text_color=STYLE["texto_claro"]
        )
        self.etiqueta_estado.configure(text="")
        self.check_label.configure(text="")
        self.boton_limpiar.configure(state="disabled")
        self.boton_generar_dictamen.configure(state="disabled")
        self.barra_progreso.set(0)
        self.etiqueta_progreso.configure(text="")

        messagebox.showinfo("Limpieza completada", "Todos los archivos y estados han sido limpiados.")

    def generar_dictamenes(self):
        """Ejecuta el generador de dictámenes PDF con barra de progreso"""
        if not self.archivo_json_generado:
            messagebox.showwarning("Sin datos", "No hay archivo JSON disponible para generar dictámenes.")
            return

        try:
            # Mostrar confirmación
            confirmacion = messagebox.askyesno(
                "Generar Dictámenes",
                "¿Está seguro de que desea generar los dictámenes PDF?\n\n"
                f"Archivo: {os.path.basename(self.archivo_json_generado)}"
            )
            
            if not confirmacion:
                return

            # Configurar UI para generación
            self.generando_dictamenes = True
            self.boton_generar_dictamen.configure(state="disabled")
            self.barra_progreso.set(0)
            self.etiqueta_progreso.configure(
                text="⏳ Iniciando generación de dictámenes...",
                text_color=STYLE["advertencia"]
            )
            self.update_idletasks()

            # Ejecutar el generador en un hilo separado
            thread = threading.Thread(target=self._ejecutar_generador_con_progreso)
            thread.daemon = True
            thread.start()

        except Exception as e:
            self.mostrar_error(f"No se pudo iniciar el generador:\n{e}")

    def _ejecutar_generador_con_progreso(self):
        """Ejecuta el generador de dictámenes en segundo plano"""
        try:
            # Importar el generador
            sys.path.append(os.path.dirname(__file__))
            from generador_dictamen import generar_dictamenes_gui
            
            # Función para actualizar progreso
            def actualizar_progreso(porcentaje, mensaje):
                self.actualizar_progreso(porcentaje, mensaje)
            
            # Función para cuando finalice
            def finalizado(exito, mensaje, resultado):
                if exito and resultado:
                    # Mostrar resultados
                    directorio = resultado['directorio']
                    total_gen = resultado['total_generados']
                    total_fam = resultado['total_familias']
                    
                    # Verificar que los archivos existen
                    archivos_existentes = []
                    if os.path.exists(directorio):
                        archivos_existentes = [f for f in os.listdir(directorio) if f.endswith('.pdf')]
                    
                    mensaje_final = f"✅ {mensaje}\n\n📁 Ubicación: {directorio}"
                    
                    if archivos_existentes:
                        mensaje_final += f"\n📄 Archivos creados: {len(archivos_existentes)}"
                    else:
                        mensaje_final += "\n⚠️  No se encontraron archivos PDF en la carpeta"
                    
                    # Mostrar mensaje
                    self.after(0, lambda: messagebox.showinfo("Generación Completada", mensaje_final))
                    
                    # Abrir carpeta si hay archivos
                    if archivos_existentes:
                        self.after(1000, lambda: self._abrir_carpeta(directorio))
                    
                else:
                    self.after(0, lambda: self.mostrar_error(mensaje))
            
            # Ejecutar generación
            generar_dictamenes_gui(
                callback_progreso=actualizar_progreso,
                callback_finalizado=finalizado
            )
            
        except Exception as e:
            self.after(0, lambda: self.mostrar_error(f"Error iniciando generador: {str(e)}"))
        finally:
            self.after(0, self._finalizar_generacion)

    def _abrir_carpeta(self, directorio):
        """Abre la carpeta en el explorador"""
        try:
            if os.path.exists(directorio):
                if os.name == 'nt':  # Windows
                    os.startfile(directorio)
                elif os.name == 'posix':  # macOS o Linux
                    os.system(f'open "{directorio}"' if sys.platform == 'darwin' else f'xdg-open "{directorio}"')
        except Exception as e:
            print(f"Error abriendo carpeta: {e}")

    def actualizar_progreso(self, porcentaje, mensaje):
        """Actualiza la barra de progreso y el mensaje (se puede llamar desde hilos)"""
        def _actualizar():
            self.barra_progreso.set(porcentaje / 100.0)
            self.etiqueta_progreso.configure(text=f"⏳ {mensaje}")
            self.update_idletasks()
        
        # Usar after para ejecutar en el hilo principal de TKinter
        self.after(0, _actualizar)

    def _finalizar_generacion(self):
        """Restaura el estado de la UI después de la generación"""
        self.generando_dictamenes = False
        self.boton_generar_dictamen.configure(state="normal")

    def _obtener_carpeta_dictamenes(self):
        """Determina la carpeta donde se guardan los dictámenes"""
        # Primero intenta encontrar la carpeta en el directorio actual
        posibles_carpetas = [
            "Dictamenes",
            "dictamenes", 
            "output",
            "Dictamenes_PDF",
            os.path.join(os.path.dirname(__file__), "Dictamenes"),
            os.path.join(os.path.dirname(__file__), "dictamenes"),
            os.path.join(os.path.dirname(__file__), "output")
        ]
        
        for carpeta in posibles_carpetas:
            if os.path.exists(carpeta) and os.path.isdir(carpeta):
                return carpeta
        
        # Si no encuentra ninguna carpeta específica, usa el directorio actual
        return os.path.dirname(__file__)

    def _abrir_explorador_archivos(self, ruta):
        """Abre el explorador de archivos en la ruta especificada"""
        try:
            sistema_operativo = platform.system()
            
            if sistema_operativo == "Windows":
                os.startfile(ruta)
            elif sistema_operativo == "Darwin":  # macOS
                subprocess.Popen(["open", ruta])
            else:  # Linux y otros
                subprocess.Popen(["xdg-open", ruta])
                
            return True
        except Exception as e:
            print(f"Error al abrir explorador: {e}")
            return False

    def _actualizar_progreso(self, valor, mensaje):
        """Actualiza la barra de progreso y el mensaje"""
        self.barra_progreso.set(valor)
        self.etiqueta_progreso.configure(text=mensaje)
        
        # Cambiar color del texto según el progreso
        if valor >= 1.0:
            self.etiqueta_progreso.configure(text_color=STYLE["exito"])
        else:
            self.etiqueta_progreso.configure(text_color=STYLE["advertencia"])

    def _generacion_completada(self, exito, mensaje, carpeta_dictamenes=None):
        """Maneja la finalización de la generación"""
        self.generando_dictamenes = False
        self.boton_generar_dictamen.configure(state="normal")
        
        if exito:
            # Mostrar mensaje de éxito
            respuesta = messagebox.askyesno(
                "Proceso Finalizado", 
                f"✅ {mensaje}\n\n¿Desea abrir la carpeta donde se guardaron los dictámenes?"
            )
            
            if respuesta and carpeta_dictamenes:
                if not self._abrir_explorador_archivos(carpeta_dictamenes):
                    messagebox.showinfo(
                        "Carpeta no encontrada", 
                        f"No se pudo abrir la carpeta automáticamente.\n\nPuede encontrar los dictámenes en:\n{carpeta_dictamenes}"
                    )
        else:
            self.mostrar_error(mensaje)
            # Resetear barra de progreso en caso de error
            self.barra_progreso.set(0)
            self.etiqueta_progreso.configure(text="")

    # -----------------------------------------------------------
    # AUXILIARES
    # -----------------------------------------------------------

    def mostrar_error(self, mensaje):
        """Muestra un error en la interfaz"""
        self.etiqueta_estado.configure(
            text="❌ Error en el proceso", 
            text_color=STYLE["peligro"]
        )
        self.check_label.configure(text="")
        messagebox.showerror("Error", mensaje)

# ================== EJECUCIÓN ================== #
if __name__ == "__main__":
    app = SistemaDictamenesVC()
    app.mainloop()