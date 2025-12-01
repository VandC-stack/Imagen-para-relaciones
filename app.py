# -- SISTEMA V&C - GENERADOR DE DICTÁMENES -- #
import os, sys, uuid, shutil
import json
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import subprocess
from datetime import datetime
import time
import platform
from datetime import datetime

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
    "texto_claro": "#ffffff",
    "borde": "#DDDDDD"
}

FONT_TITLE = ("Inter", 22, "bold")
FONT_SUBTITLE = ("Inter", 17, "bold")
FONT_LABEL = ("Inter", 13)
FONT_SMALL = ("Inter", 12)

class SistemaDictamenesVC(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración general
        self.title("Generador de Dictámenes")
        self.geometry("1275x600")
        self.minsize(1275, 600)
        ctk.set_appearance_mode("light")
        self.configure(fg_color=STYLE["fondo"])

        # Variables de estado
        self.archivo_excel_cargado = None
        self.archivo_json_generado = None
        self.json_filename = None
        self.generando_dictamenes = False
        self.clientes_data = []
        self.cliente_seleccionado = None
        self.archivo_etiquetado_json = None

        # Variables para nueva visita
        self.current_folio = "0001"

        # ===== NUEVAS VARIABLES PARA HISTORIAL =====
        self.historial_data = []
        self.historial_data_original = []
        self.historial_path = os.path.join(os.path.dirname(__file__), "data", "historial_visitas.json")
        
        # INICIALIZAR self.historial COMO DICCIONARIO
        self.historial = {"visitas": []}

        # ===== NUEVA VARIABLE PARA FOLIOS POR VISITA =====
        # Crear directorios necesarios
        data_dir = os.path.join(os.path.dirname(__file__), "data")
        os.makedirs(data_dir, exist_ok=True)
        
        self.folios_visita_path = os.path.join(data_dir, "folios_visitas")
        os.makedirs(self.folios_visita_path, exist_ok=True)

        # ===== NUEVA ESTRUCTURA DE NAVEGACIÓN =====
        self.crear_navegacion()
        self.crear_area_contenido()

        # ===== FOOTER =====
        self.crear_footer()

        # Cargar clientes al iniciar
        self.cargar_clientes_desde_json()
        self.cargar_ultimo_folio()

        # --------------------------- ICONO ---------------------------- #
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        try:
            icon_path = resource_path("img/icono.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
                print(f"🟡 Icono cargado: {icon_path}")
            else:
                print("⚠ No se encontró icono.ico")
        except Exception as e:
            print(f"⚠ Error cargando icono.ico: {e}")

    def centerwindow(self):
        self.update_idletasks()
        ancho_ventana = self.winfo_width()
        alto_ventana = self.winfo_height()
        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        self.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

    def crear_navegacion(self):
        """Crea la barra de navegación con botones mejorados"""
        nav_frame = ctk.CTkFrame(self, fg_color=STYLE["surface"], height=60)
        nav_frame.pack(fill="x", padx=20, pady=(15, 0))
        nav_frame.pack_propagate(False)
        
        # Contenedor para los botones
        botones_frame = ctk.CTkFrame(nav_frame, fg_color="transparent")
        botones_frame.pack(expand=True, fill="both", padx=20, pady=12)
        
        # Botón Principal con estilo mejorado
        self.btn_principal = ctk.CTkButton(
            botones_frame,
            text="🏠 Principal",
            command=self.mostrar_principal,
            font=("Inter", 14, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=38,
            width=130,
            corner_radius=10,
            border_width=2,
            border_color=STYLE["secundario"]
        )
        self.btn_principal.pack(side="left", padx=(0, 10))
        
        # Botón Historial con estilo mejorado
        self.btn_historial = ctk.CTkButton(
            botones_frame,
            text="📊 Historial",
            command=self.mostrar_historial,
            font=("Inter", 14, "bold"),
            fg_color=STYLE["surface"],
            hover_color=STYLE["primario"],
            text_color=STYLE["secundario"],
            height=38,
            width=130,
            corner_radius=10,
            border_width=2,
            border_color=STYLE["secundario"]
        )
        self.btn_historial.pack(side="left", padx=(0, 10))
        
        # Espacio flexible
        ctk.CTkLabel(botones_frame, text="", fg_color="transparent").pack(side="left", expand=True)
        
        # Información del sistema
        self.lbl_info_sistema = ctk.CTkLabel(
            botones_frame,
            text="Sistema de Dictámenes - V&C",
            font=("Inter", 12),
            text_color=STYLE["texto_claro"]
        )
        self.lbl_info_sistema.pack(side="right")

    def crear_area_contenido(self):
        """Crea el área de contenido donde se muestran las secciones"""
        # Frame contenedor del contenido
        self.contenido_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.contenido_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        
        # Frame para el contenido principal
        self.frame_principal = ctk.CTkFrame(self.contenido_frame, fg_color="transparent")
        
        # Frame para el historial
        self.frame_historial = ctk.CTkFrame(self.contenido_frame, fg_color="transparent")
        
        # Construir el contenido de cada sección
        self._construir_tab_principal(self.frame_principal)
        self._construir_tab_historial(self.frame_historial)
        
        # Mostrar la sección principal por defecto
        self.mostrar_principal()

    def mostrar_principal(self):
        """Muestra la sección principal y oculta las demás"""
        # Ocultar todos los frames primero
        self.frame_principal.pack_forget()
        self.frame_historial.pack_forget()
        
        # Mostrar el frame principal
        self.frame_principal.pack(fill="both", expand=True)
        
        # Actualizar estado de los botones con mejor contraste
        self.btn_principal.configure(
            fg_color=STYLE["primario"],
            text_color=STYLE["secundario"],
            border_color=STYLE["primario"]
        )
        self.btn_historial.configure(
            fg_color=STYLE["surface"],
            text_color=STYLE["secundario"],
            border_color=STYLE["secundario"]
        )

    def mostrar_historial(self):
            """Muestra la sección de historial y oculta las demás"""
            # Ocultar todos los frames primero
            self.frame_principal.pack_forget()
            self.frame_historial.pack_forget()
            
            # Mostrar el frame de historial
            self.frame_historial.pack(fill="both", expand=True)
            
            # Actualizar estado de los botones con mejor contraste
            self.btn_principal.configure(
                fg_color=STYLE["surface"],
                text_color=STYLE["secundario"],
                border_color=STYLE["secundario"]
            )
            self.btn_historial.configure(
                fg_color=STYLE["primario"],
                text_color=STYLE["secundario"],
                border_color=STYLE["primario"]
            )
            
            # Verificar y reparar datos existentes al mostrar historial
            self.verificar_datos_folios_existentes()
            
            # Refrescar el historial si es necesario
            self._cargar_historial()
            self._poblar_historial_ui()

    def _construir_tab_principal(self, parent):
        """Construye la interfaz principal con dos tarjetas en proporción 30%/70%"""
        # ===== CONTENEDOR PRINCIPAL CON 2 COLUMNAS =====
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True)

        # Configurar grid para 2 columnas con proporción 30%/70%
        main_frame.grid_columnconfigure(0, weight=3)  # 30%
        main_frame.grid_columnconfigure(1, weight=7)  # 70%
        main_frame.grid_rowconfigure(0, weight=1)

        # ===== TARJETA INFORMACIÓN DE VISITA (IZQUIERDA) - 30% =====
        card_visita = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12)
        card_visita.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="nsew")

        ctk.CTkLabel(
            card_visita,
            text="📋 Información de Visita",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 15))

        visita_frame = ctk.CTkFrame(card_visita, fg_color="transparent")
        visita_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor scrollable con scrollbar personalizada
        scroll_form = ctk.CTkScrollableFrame(
            visita_frame, 
            fg_color="transparent",
            scrollbar_button_color=STYLE["primario"],  # Color #ecd925
            scrollbar_button_hover_color=STYLE["primario"]
        )
        scroll_form.pack(fill="both", expand=True)

        # Folio de visita (automático)
        folio_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        folio_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            folio_frame,
            text="Folio Visita:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_folio_visita = ctk.CTkEntry(
            folio_frame,
            placeholder_text="0001",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_folio_visita.pack(fill="x", pady=(0, 5))
        self.entry_folio_visita.insert(0, self.current_folio)
        self.entry_folio_visita.configure(state="readonly")

        # Folio de acta (automático - AC + folio visita)
        acta_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        acta_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            acta_frame,
            text="Folio Acta:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_folio_acta = ctk.CTkEntry(
            acta_frame,
            placeholder_text="AC0001",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_folio_acta.pack(fill="x", pady=(0, 5))
        self.entry_folio_acta.insert(0, f"AC{self.current_folio}")
        self.entry_folio_acta.configure(state="readonly")

        # Fecha Inicio
        fecha_inicio_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        fecha_inicio_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            fecha_inicio_frame,
            text="Fecha Inicio (dd/mm/yyyy):",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_fecha_inicio = ctk.CTkEntry(
            fecha_inicio_frame,
            placeholder_text="dd/mm/yyyy",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_fecha_inicio.pack(fill="x", pady=(0, 5))
        self.entry_fecha_inicio.insert(0, datetime.now().strftime("%d/%m/%Y"))

        # Hora Inicio
        hora_inicio_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        hora_inicio_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            hora_inicio_frame,
            text="Hora Inicio (HH:MM):",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_hora_inicio = ctk.CTkEntry(
            hora_inicio_frame,
            placeholder_text="HH:MM",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_hora_inicio.pack(fill="x", pady=(0, 5))
        self.entry_hora_inicio.insert(0, datetime.now().strftime("%H:%M"))

        # Fecha Termino
        fecha_termino_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        fecha_termino_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            fecha_termino_frame,
            text="Fecha Termino (dd/mm/yyyy):",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_fecha_termino = ctk.CTkEntry(
            fecha_termino_frame,
            placeholder_text="dd/mm/yyyy",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_fecha_termino.pack(fill="x", pady=(0, 5))

        # Hora Termino
        hora_termino_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        hora_termino_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            hora_termino_frame,
            text="Hora Termino (HH:MM):",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_hora_termino = ctk.CTkEntry(
            hora_termino_frame,
            placeholder_text="HH:MM",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_hora_termino.pack(fill="x", pady=(0, 5))

        # Nombre Supervisor
        supervisor_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        supervisor_frame.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            supervisor_frame,
            text="Nombre Supervisor:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.entry_supervisor = ctk.CTkEntry(
            supervisor_frame,
            placeholder_text="Nombre del supervisor...",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_supervisor.pack(fill="x", pady=(0, 5))

        # ===== TARJETA GENERACIÓN (DERECHA) - 70% =====
        card_generacion = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12)
        card_generacion.grid(row=0, column=1, padx=(10, 0), pady=0, sticky="nsew")

        ctk.CTkLabel(
            card_generacion,
            text="🚀 Generación de Dictámenes",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 15))

        generacion_frame = ctk.CTkFrame(card_generacion, fg_color="transparent")
        generacion_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor scrollable con scrollbar personalizada
        scroll_generacion = ctk.CTkScrollableFrame(
            generacion_frame, 
            fg_color="transparent",
            scrollbar_button_color=STYLE["primario"],  # Color #ecd925
            scrollbar_button_hover_color=STYLE["primario"]
        )
        scroll_generacion.pack(fill="both", expand=True)

        # --- SELECCIONAR CLIENTE ---
        cliente_section = ctk.CTkFrame(scroll_generacion, fg_color="transparent")
        cliente_section.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            cliente_section,
            text="👤 Seleccionar Cliente",
            font=FONT_LABEL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 10))

        cliente_controls_frame = ctk.CTkFrame(cliente_section, fg_color="transparent")
        cliente_controls_frame.pack(fill="x", pady=(0, 10))

        self.combo_cliente = ctk.CTkComboBox(
            cliente_controls_frame,
            values=["Seleccione un cliente..."],
            font=FONT_SMALL,
            dropdown_font=FONT_SMALL,
            state="readonly",
            height=35,
            corner_radius=8,
            command=self.actualizar_cliente_seleccionado
        )
        self.combo_cliente.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.boton_limpiar_cliente = ctk.CTkButton(
            cliente_controls_frame,
            text="✕",
            command=self.limpiar_cliente,
            font=("Inter", 14, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=35,
            width=35,
            corner_radius=8,
            state="disabled"
        )
        self.boton_limpiar_cliente.pack(side="left")

        self.info_cliente = ctk.CTkLabel(
            cliente_section,
            text="No se ha seleccionado ningún cliente",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"],
            wraplength=350
        )
        self.info_cliente.pack(anchor="w", fill="x")

        # --- CARGAR TABLA DE RELACIÓN ---
        carga_section = ctk.CTkFrame(scroll_generacion, fg_color="transparent")
        carga_section.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            carga_section,
            text="📊 Cargar Tabla de Relación",
            font=FONT_LABEL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 10))

        self.info_archivo = ctk.CTkLabel(
            carga_section,
            text="No se ha cargado ningún archivo",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"],
            wraplength=350
        )
        self.info_archivo.pack(anchor="w", pady=(0, 10))

        botones_carga_frame = ctk.CTkFrame(carga_section, fg_color="transparent")
        botones_carga_frame.pack(fill="x", pady=(0, 10))

        botones_fila1 = ctk.CTkFrame(botones_carga_frame, fg_color="transparent")
        botones_fila1.pack(fill="x", pady=(0, 8))

        self.boton_cargar_excel = ctk.CTkButton(
            botones_fila1,
            text="Subir archivo",
            command=self.cargar_excel,
            font=("Inter", 13, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=35,
            width=110,
            corner_radius=8
        )
        self.boton_cargar_excel.pack(side="left", padx=(0, 8))

        # Botón Verificar Datos movido aquí
        ctk.CTkButton(
            botones_fila1,
            text="🔍 Verificar Datos",
            command=self.verificar_integridad_datos,
            font=("Inter", 11),
            fg_color=STYLE["advertencia"],
            hover_color="#b85a52",
            text_color=STYLE["surface"],
            height=35,
            width=100,
            corner_radius=8
        ).pack(side="left", padx=(0, 8))

        self.boton_limpiar = ctk.CTkButton(
            botones_fila1,
            text="Limpiar",
            command=self.limpiar_archivo,
            font=("Inter", 13),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=35,
            width=70,
            corner_radius=8,
            state="disabled"
        )
        self.boton_limpiar.pack(side="left", padx=(0, 8))

        # Botón de etiquetado DECATHLON (inicialmente oculto)
        self.boton_subir_etiquetado = ctk.CTkButton(
            botones_fila1,
            text="📦 Etiquetado DECATHLON",
            command=self.cargar_base_etiquetado,
            font=("Inter", 12, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=35,
            width=160,
            corner_radius=8
        )
        # Inicialmente no se muestra
        self.boton_subir_etiquetado.pack(side="left", padx=(0, 8))
        self.boton_subir_etiquetado.pack_forget()  # Ocultar inicialmente

        self.info_etiquetado = ctk.CTkLabel(
            botones_carga_frame,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"],
            wraplength=350
        )
        self.info_etiquetado.pack(anchor="w", pady=(5, 0))

        estado_carga_frame = ctk.CTkFrame(carga_section, fg_color="transparent")
        estado_carga_frame.pack(fill="x", pady=(0, 15))

        self.etiqueta_estado = ctk.CTkLabel(
            estado_carga_frame,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.etiqueta_estado.pack(side="left")

        self.check_label = ctk.CTkLabel(
            estado_carga_frame,
            text="",
            font=("Inter", 16, "bold"),
            text_color=STYLE["exito"]
        )
        self.check_label.pack(side="right")

        # --- GENERAR DICTÁMENES ---
        generar_section = ctk.CTkFrame(scroll_generacion, fg_color="transparent")
        generar_section.pack(fill="x", pady=(0, 0))

        ctk.CTkLabel(
            generar_section,
            text="🧾 Generar Dictámenes PDF",
            font=FONT_LABEL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 10))

        self.info_generacion = ctk.CTkLabel(
            generar_section,
            text="Seleccione un cliente y cargue la tabla para habilitar",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"]
        )
        self.info_generacion.pack(anchor="w", pady=(0, 10))

        # Barra de progreso
        self.barra_progreso = ctk.CTkProgressBar(
            generar_section,
            progress_color=STYLE["primario"],
            height=10,
            corner_radius=5
        )
        self.barra_progreso.pack(fill="x", pady=(5, 8))
        self.barra_progreso.set(0)

        self.etiqueta_progreso = ctk.CTkLabel(
            generar_section,
            text="",
            font=("Inter", 11),
            text_color=STYLE["texto_claro"]
        )
        self.etiqueta_progreso.pack(pady=(0, 8))

        # Botón de generación
        self.boton_generar_dictamen = ctk.CTkButton(
            generar_section,
            text="Generar Dictámenes",
            command=self.generar_dictamenes,
            font=("Inter", 13, "bold"),
            fg_color=STYLE["exito"],
            hover_color="#1f8c4d",
            text_color=STYLE["surface"],
            height=38,
            corner_radius=8,
            state="disabled"
        )
        self.boton_generar_dictamen.pack(pady=(0, 5))

    def _construir_tab_historial(self, parent):
        """Construye la pestaña de historial con columnas mejor organizadas"""
        cont = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=8)
        cont.pack(fill="both", expand=True, padx=0, pady=0)

        # ===== BARRA SUPERIOR CON BUSCADORES EN LÍNEA =====
        barra_superior_historial = ctk.CTkFrame(cont, fg_color="transparent", height=60)
        barra_superior_historial.pack(fill="x", pady=(0, 0))
        barra_superior_historial.pack_propagate(False)

        # Frame para los buscadores en horizontal
        buscadores_frame = ctk.CTkFrame(barra_superior_historial, fg_color="transparent")
        buscadores_frame.pack(side="left", fill="x", expand=True, padx=0, pady=0)

        # Buscador por folio (primero)
        busqueda_folio_frame = ctk.CTkFrame(buscadores_frame, fg_color="transparent")
        busqueda_folio_frame.pack(side="left", padx=(0, 20))

        ctk.CTkLabel(
            busqueda_folio_frame,
            text="Buscar por folio visita:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(0, 8))

        self.entry_buscar_folio = ctk.CTkEntry(
            busqueda_folio_frame,
            placeholder_text="Ej: 0001",
            width=120,
            height=32,
            corner_radius=8
        )
        self.entry_buscar_folio.pack(side="left", padx=(0, 8))

        btn_buscar_folio = ctk.CTkButton(
            busqueda_folio_frame,
            text="Buscar",
            command=self.hist_buscar_por_folio,
            font=("Inter", 11, "bold"),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=32,
            width=80,
            corner_radius=8
        )
        btn_buscar_folio.pack(side="left")

        # Buscador general (segundo, al lado del primero)
        busqueda_general_frame = ctk.CTkFrame(buscadores_frame, fg_color="transparent")
        busqueda_general_frame.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            busqueda_general_frame,
            text="Búsqueda general:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(0, 8))

        self.entry_buscar_general = ctk.CTkEntry(
            busqueda_general_frame,
            placeholder_text="Buscar por cliente, folio, fecha...",
            width=220,
            height=32,
            corner_radius=8
        )
        self.entry_buscar_general.pack(side="left", padx=(0, 8))
        self.entry_buscar_general.bind("<KeyRelease>", self.hist_buscar_general)

        btn_limpiar_busqueda = ctk.CTkButton(
            busqueda_general_frame,
            text="Limpiar",
            command=self.hist_limpiar_busqueda,
            font=("Inter", 11),
            fg_color=STYLE["advertencia"],
            hover_color="#b85a52",
            text_color=STYLE["surface"],
            height=32,
            width=70,
            corner_radius=8
        )
        btn_limpiar_busqueda.pack(side="left")

        # ===== TABLA MEJORADA CON NUEVAS COLUMNAS =====
        tabla_container = ctk.CTkFrame(cont, fg_color=STYLE["fondo"], corner_radius=8)
        tabla_container.pack(fill="both", expand=True, pady=(0, 10))

        # Encabezados de la tabla (agregada columna Supervisor)
        header_frame = ctk.CTkFrame(tabla_container, fg_color=STYLE["secundario"], height=35)
        header_frame.pack(fill="x", padx=0, pady=(0, 1))
        header_frame.pack_propagate(False)

        # Configuración de anchos fijos para cada columna (incluyendo Supervisor)
        column_widths = [90, 90, 100, 100, 90, 90, 180, 120, 100, 120, 110]

        headers = [
            "Folio Visita",
            "Folio Acta", 
            "Fecha Inicio",
            "Fecha Termino",
            "Hora Inicio",
            "Hora Termino",
            "Cliente",
            "Supervisor",
            "Estatus",
            "Folios Usados",
            "Acciones"
        ]

        # Crear headers
        for i, header_text in enumerate(headers):
            lbl = ctk.CTkLabel(
                header_frame, 
                text=header_text, 
                font=("Inter", 12, "bold"),
                text_color=STYLE["surface"],
                width=column_widths[i],
                anchor="center"
            )
            lbl.pack(side="left", padx=1)

        # Área scrollable para los registros
        self.hist_scroll = ctk.CTkScrollableFrame(
            tabla_container, 
            fg_color=STYLE["fondo"],
            scrollbar_button_color=STYLE["primario"],
            scrollbar_button_hover_color=STYLE["primario"]
        )
        self.hist_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # ===== PIE DE PÁGINA VISIBLE =====
        footer_frame = ctk.CTkFrame(cont, fg_color="transparent", height=40)
        footer_frame.pack(fill="x", side="bottom", pady=(5, 0))
        footer_frame.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer_frame, fg_color="transparent")
        footer_content.pack(expand=True, fill="both", padx=0, pady=5)

        self.hist_info_label = ctk.CTkLabel(
            footer_content, 
            text="Sistema de historial de visitas - V&C", 
            font=FONT_SMALL, 
            text_color=STYLE["texto_claro"]
        )
        self.hist_info_label.pack(side="left")
        
        ctk.CTkButton(
            footer_content, 
            text="📁 Respaldar historial", 
            command=self.hist_hacer_backup,
            font=("Inter", 11, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=30,
            width=120,
            corner_radius=6
        ).pack(side="right", padx=(5, 0))

        # Cargar datos del historial
        self.historial_path = os.path.join(os.path.dirname(__file__), "data", "historial_visitas.json")
        self._cargar_historial()
        self._poblar_historial_ui()





    def _formatear_hora_12h(self, hora_str):
        """Convierte hora de formato 24h a formato 12h con AM/PM de forma consistente"""
        if not hora_str or hora_str.strip() == "":
            return ""
        
        try:
            # Limpiar y estandarizar la cadena
            hora_str = str(hora_str).strip()
            
            # Si ya contiene AM/PM, devolver tal cual (pero limpiando espacios)
            hora_str_upper = hora_str.upper()
            if "AM" in hora_str_upper or "PM" in hora_str_upper:
                # Ya está en formato 12h, solo limpiar
                # Asegurar que AM/PM estén separados correctamente
                if "AM" in hora_str_upper:
                    hora_str = hora_str_upper.replace("AM", " AM")
                elif "PM" in hora_str_upper:
                    hora_str = hora_str_upper.replace("PM", " PM")
                return hora_str.strip()
            
            # Reemplazar punto por dos puntos (por si viene como "17.25")
            hora_str = hora_str.replace(".", ":")
            
            # Parsear la hora
            if ":" in hora_str:
                partes = hora_str.split(":")
                hora = int(partes[0].strip())
                minutos = partes[1].strip()[:2]  # Tomar solo los primeros 2 dígitos
                
                # Formatear minutos a 2 dígitos
                if len(minutos) == 1:
                    minutos = f"0{minutos}"
                
                # Determinar AM/PM
                if hora == 0:
                    return f"12:{minutos} AM"
                elif hora < 12:
                    return f"{hora}:{minutos} AM"
                elif hora == 12:
                    return f"12:{minutos} PM"
                else:
                    return f"{hora-12}:{minutos} PM"
            else:
                # Si no tiene formato de hora, devolver tal cual
                return hora_str
        except Exception as e:
            print(f"⚠️ Error formateando hora {hora_str}: {e}")
            return hora_str

    def crear_footer(self):
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
    # MÉTODOS PARA GESTIÓN DE CLIENTES
    # -----------------------------------------------------------
    def cargar_clientes_desde_json(self):
        try:
            posibles_rutas = [
                'data/Clientes.json',
                'Clientes.json',
                '../data/Clientes.json'
            ]
            
            archivo_encontrado = None
            for ruta in posibles_rutas:
                if os.path.exists(ruta):
                    archivo_encontrado = ruta
                    break
            
            if not archivo_encontrado:
                print("⚠️  No se encontró el archivo Clientes.json")
                return
            
            with open(archivo_encontrado, 'r', encoding='utf-8') as f:
                self.clientes_data = json.load(f)
            
            self.clientes_data.sort(key=lambda x: x['CLIENTE'])
            nombres_clientes = [cliente['CLIENTE'] for cliente in self.clientes_data]
            self.combo_cliente.configure(values=nombres_clientes)
            print(f"✅ Clientes cargados: {len(nombres_clientes)} clientes")
            
        except Exception as e:
            print(f"❌ Error al cargar clientes: {e}")
            messagebox.showerror("Error", f"No se pudieron cargar los clientes:\n{e}")

    def actualizar_cliente_seleccionado(self, cliente_nombre):
        if cliente_nombre == "Seleccione un cliente...":
            self.cliente_seleccionado = None
            self.info_cliente.configure(
                text="No se ha seleccionado ningún cliente",
                text_color=STYLE["texto_claro"]
            )
            self.boton_limpiar_cliente.configure(state="disabled")
            self.boton_subir_etiquetado.pack_forget()
            self.info_etiquetado.pack_forget()
            return

        for cliente in self.clientes_data:
            if cliente['CLIENTE'] == cliente_nombre:
                self.cliente_seleccionado = cliente
                rfc = cliente.get('RFC', 'No disponible')

                self.info_cliente.configure(
                    text=f"✅ {cliente_nombre}\n📋 RFC: {rfc}",
                    text_color=STYLE["exito"]
                )
                self.boton_limpiar_cliente.configure(state="normal")

                if cliente_nombre == "ARTICULOS DEPORTIVOS DECATHLON SA DE CV":
                    self.boton_subir_etiquetado.pack(side="left", padx=(0, 8))
                    if self.archivo_etiquetado_json:
                        self.info_etiquetado.pack(anchor="w", fill="x", pady=(5, 0))
                else:
                    self.boton_subir_etiquetado.pack_forget()
                    self.info_etiquetado.pack_forget()

                if self.archivo_json_generado:
                    self.boton_generar_dictamen.configure(state="normal")
                break

    def cargar_base_etiquetado(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Base de Etiquetado DECATHLON",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            if df.empty:
                messagebox.showwarning("Archivo vacío", "El archivo de etiquetado no contiene datos.")
                return

            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)

            registros = df.to_dict(orient="records")

            data_dir = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_dir, exist_ok=True)

            output_json = os.path.join(data_dir, "base_etiquetado.json")

            with open(output_json, "w", encoding="utf-8") as f:
                json.dump(registros, f, ensure_ascii=False, indent=2)

            self.archivo_etiquetado_json = output_json

            self.info_etiquetado.configure(
                text=f"📄 Base de etiquetado cargada ({len(registros)} registros)",
                text_color=STYLE["exito"]
            )
            self.info_etiquetado.pack(anchor="w", fill="x", pady=(5, 0))

            messagebox.showinfo(
                "Base cargada",
                f"Base de etiquetado convertida exitosamente.\n\nGuardado en:\n{output_json}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar la base de etiquetado:\n{e}")

    def limpiar_cliente(self):
        self.combo_cliente.set("Seleccione un cliente...")
        self.cliente_seleccionado = None
        self.info_cliente.configure(
            text="No se ha seleccionado ningún cliente",
            text_color=STYLE["texto_claro"]
        )
        self.boton_limpiar_cliente.configure(state="disabled")
        self.boton_generar_dictamen.configure(state="disabled")
        self.boton_subir_etiquetado.pack_forget()
        self.info_etiquetado.pack_forget()

    # -----------------------------------------------------------
    # MÉTODOS MEJORADOS PARA GESTIÓN DE FOLIOS
    # -----------------------------------------------------------
    def cargar_ultimo_folio(self):
        """Carga el último folio utilizado y determina el siguiente disponible"""
        try:
            if os.path.exists(self.historial_path):
                with open(self.historial_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                visitas = data.get("visitas", [])
                if visitas:
                    # Obtener todos los folios existentes
                    folios_existentes = set()
                    for visita in visitas:
                        folio = visita.get("folio_visita", "0")
                        if folio.isdigit():
                            folios_existentes.add(int(folio))
                    
                    # Encontrar el primer folio disponible
                    folio_disponible = 1
                    while folio_disponible in folios_existentes:
                        folio_disponible += 1
                    
                    self.current_folio = f"{folio_disponible:04d}"
                else:
                    self.current_folio = "0001"
                    
                # Actualizar el campo en la interfaz
                if hasattr(self, 'entry_folio_visita'):
                    self.entry_folio_visita.configure(state="normal")
                    self.entry_folio_visita.delete(0, "end")
                    self.entry_folio_visita.insert(0, self.current_folio)
                    self.entry_folio_visita.configure(state="readonly")
                    
                    # Actualizar también el folio del acta
                    self.entry_folio_acta.configure(state="normal")
                    self.entry_folio_acta.delete(0, "end")
                    self.entry_folio_acta.insert(0, f"AC{self.current_folio}")
                    self.entry_folio_acta.configure(state="readonly")
                    
        except Exception as e:
            print(f"❌ Error cargando último folio: {e}")
            self.current_folio = "0001"

    def crear_nueva_visita(self):
        """Prepara el formulario para una nueva visita"""
        try:
            # Obtener el siguiente folio disponible
            self.cargar_ultimo_folio()
            
            # Actualizar campos
            self.entry_folio_visita.configure(state="normal")
            self.entry_folio_visita.delete(0, "end")
            self.entry_folio_visita.insert(0, self.current_folio)
            self.entry_folio_visita.configure(state="readonly")
            
            # Actualizar folio acta automáticamente
            self.entry_folio_acta.configure(state="normal")
            self.entry_folio_acta.delete(0, "end")
            self.entry_folio_acta.insert(0, f"AC{self.current_folio}")
            self.entry_folio_acta.configure(state="readonly")
            
            # Limpiar otros campos
            self.entry_fecha_inicio.delete(0, "end")
            self.entry_fecha_inicio.insert(0, datetime.now().strftime("%d/%m/%Y"))
            self.entry_hora_inicio.delete(0, "end")
            self.entry_hora_inicio.insert(0, datetime.now().strftime("%H:%M"))
            self.entry_fecha_termino.delete(0, "end")
            self.entry_hora_termino.delete(0, "end")
            self.entry_supervisor.delete(0, "end")
            
            messagebox.showinfo("Nueva Visita", "Formulario listo para nueva visita")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear nueva visita:\n{e}")

    def guardar_visita_desde_formulario(self):
        """Guarda una nueva visita desde el formulario principal"""
        try:
            if not self.cliente_seleccionado:
                messagebox.showwarning("Cliente requerido", "Por favor seleccione un cliente primero.")
                return

            # Recoger datos del formulario
            folio_visita = self.entry_folio_visita.get().strip()
            folio_acta = self.entry_folio_acta.get().strip()
            fecha_inicio = self.entry_fecha_inicio.get().strip()
            fecha_termino = self.entry_fecha_termino.get().strip()
            hora_inicio = self.entry_hora_inicio.get().strip()
            hora_termino = self.entry_hora_termino.get().strip()
            supervisor = self.entry_supervisor.get().strip()

            if not folio_acta:
                messagebox.showwarning("Datos incompletos", "Por favor ingrese el folio de acta.")
                return

            # Validar que el folio acta tenga formato correcto
            if not folio_acta.startswith("AC") or len(folio_acta) != 6:
                messagebox.showwarning("Formato incorrecto", "El folio de acta debe tener formato ACXXXX (ej: AC0001).")
                return

            # Crear payload con todos los campos
            payload = {
                "folio_visita": folio_visita,
                "folio_acta": folio_acta,
                "fecha_inicio": fecha_inicio,
                "fecha_termino": fecha_termino,
                "hora_inicio": hora_inicio,
                "hora_termino": hora_termino,
                "norma": "",
                "cliente": self.cliente_seleccionado['CLIENTE'],
                "nfirma1": supervisor,  # Usamos supervisor como única firma
                "nfirma2": "",
                "estatus": "En proceso"
            }

            # Guardar visita
            self.hist_create_visita(payload)
            
            # Limpiar formulario después de guardar
            self.crear_nueva_visita()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la visita:\n{e}")

    # -----------------------------------------------------------
    # MÉTODOS PARA CARGA Y GENERACIÓN DE ARCHIVOS
    # -----------------------------------------------------------
    def cargar_excel(self):
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
        
        self.boton_cargar_excel.configure(state="disabled")
        self.boton_limpiar.configure(state="normal")
        
        self.etiqueta_estado.configure(
            text="⏳ Convirtiendo a JSON...", 
            text_color=STYLE["advertencia"]
        )
        self.check_label.configure(text="")
        self.update_idletasks()

        thread = threading.Thread(target=self.convertir_a_json, args=(file_path,))
        thread.daemon = True
        thread.start()

    def convertir_a_json(self, file_path):
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                self.mostrar_error("El archivo seleccionado no contiene datos.")
                return

            # Convertir columnas de fecha a string
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)

            # Limpiar nombres de columnas (eliminar espacios extra)
            df.columns = df.columns.str.strip()

            # Buscar y renombrar la columna de solicitud para consistencia
            col_solicitud = self._obtener_columna_solicitud(df)
            if col_solicitud and col_solicitud != 'SOLICITUD':
                df.rename(columns={col_solicitud: 'SOLICITUD'}, inplace=True)

            records = df.to_dict(orient="records")

            data_folder = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_folder, exist_ok=True)

            self.json_filename = "tabla_de_relacion.json"
            output_path = os.path.join(data_folder, self.json_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(records, f, ensure_ascii=False, indent=2)

            # EXTRAER Y GUARDAR INFORMACIÓN DE FOLIOS
            self._extraer_informacion_folios(records)

            # GUARDAR FOLIOS PARA VISITA ACTUAL CON PERSISTENCIA
            if hasattr(self, 'current_folio') and self.current_folio:
                # Crear también un backup de la tabla de relación
                backup_dir = os.path.join(data_folder, "tabla_relacion_backups")
                os.makedirs(backup_dir, exist_ok=True)
                
                backup_path = os.path.join(
                    backup_dir, 
                    f"tabla_relacion_backup_{self.current_folio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                )
                
                with open(backup_path, "w", encoding="utf-8") as backup_file:
                    json.dump(records, backup_file, ensure_ascii=False, indent=2)
                
                print(f"📁 Backup de tabla de relación creado: {backup_path}")
                
                # Guardar folios de la visita
                self.guardar_folios_visita(self.current_folio, records)

            self.after(0, self._actualizar_ui_conversion_exitosa, output_path, len(records))

        except Exception as e:
            self.after(0, self.mostrar_error, f"Error al convertir el archivo:\n{e}")
    
    def _extraer_informacion_folios(self, datos_tabla):
        """Extrae y procesa la información de folios de la tabla de relación"""
        try:
            folios_encontrados = []
            folios_numericos = []
            
            # Buscar la columna FOLIO en los datos
            for item in datos_tabla:
                if 'FOLIO' in item and item['FOLIO'] is not None:
                    folio_valor = item['FOLIO']
                    
                    # Convertir a string y limpiar
                    folio_str = str(folio_valor).strip()
                    
                    # Intentar convertir a número y formatear a 6 dígitos
                    try:
                        # Manejar casos donde folio_str puede ser decimal
                        folio_num = int(float(folio_str))
                        folios_numericos.append(folio_num)
                        folios_encontrados.append(f"{folio_num:06d}")
                    except (ValueError, TypeError):
                        # Si no se puede convertir, usar el valor original
                        folios_encontrados.append(folio_str)
            
            # Procesar la información de folios
            info_folios = {
                "total_folios": len(folios_encontrados),
                "folios_unicos": len(set(folios_encontrados)),
                "rango_folios": "",
                "lista_folios": folios_encontrados,
                "folios_formateados": folios_encontrados  # Agregar versión formateada
            }
            
            # Calcular rango si hay folios numéricos
            if folios_numericos:
                min_folio = min(folios_numericos)
                max_folio = max(folios_numericos)
                info_folios["rango_folios"] = f"{min_folio:06d} - {max_folio:06d}"
                info_folios["total_numericos"] = len(folios_numericos)
                info_folios["rango_numerico"] = f"{min_folio} - {max_folio}"
            
            # Guardar información de folios para usar después
            self.info_folios_actual = info_folios
            
            print(f"📊 Información de folios extraída:")
            print(f"   - Total folios: {info_folios['total_folios']}")
            print(f"   - Folios únicos: {info_folios['folios_unicos']}")
            print(f"   - Rango formateado: {info_folios['rango_folios']}")
            if folios_numericos:
                print(f"   - Rango numérico: {info_folios['rango_numerico']}")
            
            return info_folios
            
        except Exception as e:
            print(f"⚠️ Error extrayendo información de folios: {e}")
            return None

    def verificar_datos_folios_existentes(self):
        """Verifica y repara datos de folios existentes para asegurar consistencia"""
        try:
            print("🔍 Verificando datos de folios existentes...")
            
            if not os.path.exists(self.folios_visita_path):
                print("ℹ️ No hay carpeta de folios para verificar")
                return
            
            # Listar todos los archivos JSON de folios
            archivos_folios = [f for f in os.listdir(self.folios_visita_path) if f.endswith('.json')]
            
            archivos_reparados = 0
            for archivo in archivos_folios:
                archivo_path = os.path.join(self.folios_visita_path, archivo)
                
                try:
                    with open(archivo_path, 'r', encoding='utf-8') as f:
                        datos = json.load(f)
                    
                    datos_modificados = False
                    
                    # Verificar y reparar cada registro
                    for item in datos:
                        # Reparar formato de FOLIOS a 6 dígitos
                        if 'FOLIOS' in item:
                            folio_raw = item['FOLIOS']
                            if folio_raw:
                                try:
                                    # Intentar convertir a número y formatear
                                    folio_num = int(float(str(folio_raw)))
                                    folio_formateado = f"{folio_num:06d}"
                                    
                                    if folio_formateado != str(folio_raw):
                                        item['FOLIOS'] = folio_formateado
                                        datos_modificados = True
                                        print(f"   🔧 Reparado: {folio_raw} -> {folio_formateado}")
                                except (ValueError, TypeError):
                                    pass
                    
                    # Guardar si hubo modificaciones
                    if datos_modificados:
                        with open(archivo_path, 'w', encoding='utf-8') as f:
                            json.dump(datos, f, ensure_ascii=False, indent=2)
                        archivos_reparados += 1
                        print(f"✅ Archivo reparado: {archivo}")
                        
                except Exception as e:
                    print(f"⚠️ Error procesando archivo {archivo}: {e}")
            
            print(f"📊 Verificación completada. Archivos reparados: {archivos_reparados}/{len(archivos_folios)}")
            
        except Exception as e:
            print(f"❌ Error en verificación de datos: {e}")

    def _obtener_folios_de_tabla(self):
        """Obtiene la información de folios de la tabla de relación con formato mejorado"""
        try:
            if not hasattr(self, 'info_folios_actual') or not self.info_folios_actual:
                return "No disponible"
            
            info = self.info_folios_actual
            
            if info['rango_folios']:
                return f"Total: {info['total_folios']} | Rango: {info['rango_folios']}"
            else:
                return f"Total: {info['total_folios']} folios"
                
        except Exception as e:
            print(f"⚠️ Error obteniendo folios de tabla: {e}")
            return "Error al obtener folios"

    def _actualizar_ui_conversion_exitosa(self, output_path, num_registros):
        self.archivo_json_generado = output_path
        self.etiqueta_estado.configure(
            text=f"✅ Convertido - {num_registros} registros", 
            text_color=STYLE["exito"]
        )
        self.check_label.configure(text="✓")
        
        if self.cliente_seleccionado:
            self.boton_generar_dictamen.configure(state="normal")
        
        messagebox.showinfo(
            "Conversión exitosa",
            f"Archivo convertido correctamente.\n\n"
            f"Ubicación: {output_path}\n"
            f"Total de registros: {num_registros}"
        )

    def limpiar_archivo(self):
        self.archivo_excel_cargado = None
        self.archivo_json_generado = None
        self.json_filename = None

        self.info_archivo.configure(
            text="No se ha cargado ningún archivo",
            text_color=STYLE["texto_claro"]
        )

        self.boton_cargar_excel.configure(state="normal")
        self.boton_limpiar.configure(state="disabled")
        self.boton_generar_dictamen.configure(state="disabled")

        self.etiqueta_estado.configure(text="", text_color=STYLE["texto_claro"])
        self.check_label.configure(text="")
        self.barra_progreso.set(0)
        self.etiqueta_progreso.configure(text="")

        try:
            data_dir = os.path.join(os.path.dirname(__file__), "data")
            
            archivos_a_eliminar = [
                "base_etiquetado.json",
                "tabla_de_relacion.json"
            ]
            
            archivos_eliminados = []
            
            for archivo in archivos_a_eliminar:
                ruta_archivo = os.path.join(data_dir, archivo)
                if os.path.exists(ruta_archivo):
                    os.remove(ruta_archivo)
                    archivos_eliminados.append(archivo)
                    print(f"🗑️ {archivo} eliminado correctamente.")
            
            if archivos_eliminados:
                print(f"✅ Se eliminaron {len(archivos_eliminados)} archivos: {', '.join(archivos_eliminados)}")
            else:
                print("ℹ️ No se encontraron archivos para eliminar.")

            self.archivo_etiquetado_json = None
            self.info_etiquetado.configure(text="")
            self.info_etiquetado.pack_forget()

        except Exception as e:
            print(f"⚠️ Error al eliminar archivos: {e}")

        messagebox.showinfo("Limpieza completa", "Los datos del archivo y el etiquetado han sido limpiados.")

    def generar_dictamenes(self):
        if not self.archivo_json_generado:
            messagebox.showwarning("Sin datos", "No hay archivo JSON disponible para generar dictámenes.")
            return

        if not self.cliente_seleccionado:
            messagebox.showwarning("Cliente no seleccionado", "Por favor seleccione un cliente antes de generar los dictámenes.")
            return

        try:
            # Leer el archivo JSON y extraer los folios
            with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                datos = json.load(f)

            # Extraer folios únicos y ordenados
            folios = set()
            for item in datos:
                if 'FOLIO' in item and item['FOLIO']:
                    try:
                        folio = int(item['FOLIO'])
                        folios.add(folio)
                    except (ValueError, TypeError):
                        # Si no se puede convertir a entero, ignorar
                        pass

            # Convertir a lista ordenada
            folios_ordenados = sorted(folios)
            self.folios_utilizados_actual = folios_ordenados

            # Continuar con la generación...
            confirmacion = messagebox.askyesno(
                "Generar Dictámenes",
                f"¿Está seguro de que desea generar los dictámenes PDF?\n\n"
                f"📄 Archivo: {os.path.basename(self.archivo_json_generado)}\n"
                f"👤 Cliente: {self.cliente_seleccionado['CLIENTE']}\n"
                f"📋 RFC: {self.cliente_seleccionado.get('RFC', 'No disponible')}\n"
                f"📊 Total de folios: {len(folios_ordenados)}"
            )
            
            if not confirmacion:
                return

            self.generando_dictamenes = True
            self.boton_generar_dictamen.configure(state="disabled")
            self.barra_progreso.set(0)
            self.etiqueta_progreso.configure(
                text="⏳ Iniciando generación de dictámenes...",
                text_color=STYLE["advertencia"]
            )
            self.update_idletasks()

            thread = threading.Thread(target=self._ejecutar_generador_con_progreso)
            thread.daemon = True
            thread.start()

        except Exception as e:
            self.mostrar_error(f"No se pudo iniciar el generador:\n{e}")

    def _actualizar_ui_conversion_exitosa(self, output_path, num_registros):
        self.archivo_json_generado = output_path
        
        # Mostrar información de folios en la interfaz si está disponible
        info_folios_text = ""
        if hasattr(self, 'info_folios_actual') and self.info_folios_actual:
            info = self.info_folios_actual
            if info['rango_folios']:
                info_folios_text = f" | 📋 Folios: {info['rango_folios']}"
            else:
                info_folios_text = f" | 📋 Folios: {info['total_folios']} encontrados"
        
        self.etiqueta_estado.configure(
            text=f" {info_folios_text}", 
            text_color=STYLE["exito"]
        )
        self.check_label.configure(text="✓")
        
        if self.cliente_seleccionado:
            self.boton_generar_dictamen.configure(state="normal")

    def _ejecutar_generador_con_progreso(self):
        try:
            # VERIFICAR SI LA VENTANA SIGUE ABIERTA
            if not self.winfo_exists():
                return
                
            sys.path.append(os.path.dirname(__file__))
            from generador_dictamen import generar_dictamenes_gui
            
            def actualizar_progreso(porcentaje, mensaje):
                # VERIFICACIÓN EN CALLBACK
                if self.winfo_exists():
                    self.actualizar_progreso(porcentaje, mensaje)
            
            def finalizado(exito, mensaje, resultado):
                # VERIFICACIÓN EN CALLBACK FINAL
                if not self.winfo_exists():
                    return
                    
                if exito and resultado:
                    directorio = resultado['directorio']
                    total_gen = resultado['total_generados']
                    total_fam = resultado['total_familias']
                    
                    dictamenes_fallidos = resultado.get('dictamenes_fallidos', 0)
                    folios_fallidos = resultado.get('folios_fallidos', [])
                    folios_utilizados = resultado.get('folios_utilizados', "No disponible")
                    
                    archivos_existentes = []
                    if os.path.exists(directorio):
                        archivos_existentes = [f for f in os.listdir(directorio) if f.endswith('.pdf')]
                    
                    mensaje_final = f"✅ {mensaje}\n\n📁 Ubicación: {directorio}"
                    
                    if archivos_existentes:
                        mensaje_final += f"\n📄 Archivos creados: {len(archivos_existentes)}"
                    else:
                        mensaje_final += "\n⚠️  No se encontraron archivos PDF en la carpeta"
                    
                    if dictamenes_fallidos > 0:
                        mensaje_final += f"\n❌ Dictámenes no generados: {dictamenes_fallidos}"
                        if folios_fallidos:
                            mensaje_final += f"\n📋 Folios fallidos: {', '.join(map(str, folios_fallidos))}"
                    
                    # VERIFICAR ANTES DE MOSTRAR MESSAGEBOX
                    if self.winfo_exists():
                        self.after(0, lambda: messagebox.showinfo("Generación Completada", mensaje_final) if self.winfo_exists() else None)
                        
                        resultado['folios_utilizados_info'] = folios_utilizados
                        self.registrar_visita_automatica(resultado)
                        
                        if archivos_existentes and self.winfo_exists():
                            self.after(1000, lambda: self._abrir_carpeta(directorio) if self.winfo_exists() else None)
                    
                else:
                    if self.winfo_exists():
                        self.after(0, lambda: self.mostrar_error(mensaje) if self.winfo_exists() else None)
            
            # LLAMADA CORREGIDA - sin folios_info
            generar_dictamenes_gui(
                cliente_manual=self.cliente_seleccionado['CLIENTE'],
                rfc_manual=self.cliente_seleccionado.get('RFC', ''),
                callback_progreso=actualizar_progreso,
                callback_finalizado=finalizado
            )
            
        except Exception as e:
            error_msg = f"Error iniciando generador: {str(e)}"
            if self.winfo_exists():
                self.after(0, lambda: self.mostrar_error(error_msg) if self.winfo_exists() else None)
        finally:
            if self.winfo_exists():
                self.after(0, self._finalizar_generacion)

    def _abrir_carpeta(self, directorio):
        try:
            if os.path.exists(directorio):
                if os.name == 'nt':
                    os.startfile(directorio)
                elif os.name == 'posix':
                    os.system(f'open "{directorio}"' if sys.platform == 'darwin' else f'xdg-open "{directorio}"')
        except Exception as e:
            print(f"Error abriendo carpeta: {e}")

    def actualizar_progreso(self, porcentaje, mensaje):
        def _actualizar():
            if self.winfo_exists():  # Verificar si la ventana aún existe
                self.barra_progreso.set(porcentaje / 100.0)
                self.etiqueta_progreso.configure(text=f"⏳ {mensaje}")
                self.update_idletasks()
        
        self.after(0, _actualizar)

    def _finalizar_generacion(self):
        if self.winfo_exists():  # Verificar si la ventana aún existe
            self.generando_dictamenes = False
            self.boton_generar_dictamen.configure(state="normal")

    def mostrar_error(self, mensaje):
        if self.winfo_exists():  # Verificar si la ventana aún existe
            self.etiqueta_estado.configure(
                text="❌ Error en el proceso", 
                text_color=STYLE["peligro"]
            )
            self.check_label.configure(text="")
            messagebox.showerror("Error", mensaje)

    # -----------------------------------------------------------
    # MÉTODOS DEL HISTORIAL
    # -----------------------------------------------------------
    def _cargar_historial(self):
        """Carga los datos del historial desde el archivo JSON"""
        try:
            # Crear directorio si no existe
            os.makedirs(os.path.dirname(self.historial_path), exist_ok=True)
            
            if os.path.exists(self.historial_path):
                with open(self.historial_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Extraer solo las visitas
                    self.historial_data = data.get("visitas", [])
                    self.historial = data  # <- CARGAR EL DICCIONARIO COMPLETO
            else:
                self.historial_data = []
                self.historial = {"visitas": []}
                
            # Inicializar también historial_data_original
            self.historial_data_original = self.historial_data.copy()
            
            print(f"✅ Historial cargado: {len(self.historial_data)} registros")
                
        except Exception as e:
            print(f"❌ Error cargando historial: {e}")
            self.historial_data = []
            self.historial_data_original = []
            self.historial = {"visitas": []}

    def hist_borrar_visita(self, id_):
        """Elimina una visita y recalcula el folio actual si es necesario"""
        if not messagebox.askyesno("Confirmar borrado", "¿Eliminar este registro?"):
            return
        
        try:
            # Encontrar la visita a borrar
            visita_a_borrar = next((v for v in self.historial.get("visitas",[]) if v["_id"] == id_), None)
            if not visita_a_borrar:
                messagebox.showerror("Error", "No se encontró la visita para borrar")
                return

            folio_borrado = visita_a_borrar.get("folio_visita", "")
            
            # Borrar la visita
            self.historial["visitas"] = [v for v in self.historial.get("visitas",[]) if v["_id"] != id_]
            self._guardar_historial()
            self._poblar_historial_ui()

            # Recalcular el folio actual (buscar el siguiente disponible)
            self.cargar_ultimo_folio()
            messagebox.showinfo("Folio actualizado", f"Se borró el folio {folio_borrado}. Folio actual recalculado: {self.current_folio}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _guardar_historial(self):
        """Guarda el historial en un único archivo"""
        try:
            # ACTUALIZAR self.historial_data DESDE self.historial
            self.historial_data = self.historial.get("visitas", [])
            
            with open(self.historial_path, "w", encoding="utf-8") as f:
                json.dump(self.historial, f, ensure_ascii=False, indent=2)
                
            self.hist_info_label.configure(text=f"Guardado OK — {len(self.historial_data)} visitas")
            print(f"✅ Historial guardado: {len(self.historial_data)} registros")
            
        except Exception as e:
            print(f"❌ Error guardando historial: {e}")
            self.hist_info_label.configure(text=f"Error guardando: {e}")
    
    def hist_hacer_backup(self):
        """Crea un respaldo manual del historial"""
        try:
            if os.path.exists(self.historial_path):
                backup_dir = os.path.join(os.path.dirname(self.historial_path), "backups")
                os.makedirs(backup_dir, exist_ok=True)
                
                backup_name = f"historial_visitas_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                backup_path = os.path.join(backup_dir, backup_name)
                
                shutil.copy2(self.historial_path, backup_path)
                messagebox.showinfo("Backup", f"Backup creado:\n{backup_path}")
            else:
                messagebox.showinfo("Backup", "No existe historial para respaldar.")
        except Exception as e:
            messagebox.showerror("Backup error", str(e))

    def _limpiar_scroll_hist(self):
        for child in self.hist_scroll.winfo_children():
            child.destroy()


    # -- BOTONES DE ACCION PARA CADA VISITA -- #
    def _poblar_historial_ui(self):
        """Poblar la interfaz de historial con datos mejorados"""
        # Limpiar el scroll
        for widget in self.hist_scroll.winfo_children():
            widget.destroy()

        # Asegurarse de que los datos estén cargados
        if not hasattr(self, 'historial_data') or not self.historial_data:
            no_data_frame = ctk.CTkFrame(self.hist_scroll, fg_color=STYLE["surface"], height=40)
            no_data_frame.pack(fill="x", pady=2)
            no_data_frame.pack_propagate(False)
            
            ctk.CTkLabel(
                no_data_frame,
                text="No hay registros en el historial",
                font=FONT_SMALL,
                text_color=STYLE["texto_claro"]
            ).pack(expand=True, fill="both")
            return

        # Configuración estandarizada para todos los botones
        BUTTON_CONFIG = {
            "width": 30,
            "height": 30,
            "corner_radius": 6,
            "font": ("Segoe UI Symbol", 12),  # Reducimos tamaño de fuente
            "text_color": STYLE["surface"]
        }

        # Colores específicos para cada tipo de botón
        BUTTON_COLORS = {
            "descargar": {"fg_color": STYLE["exito"], "hover_color": "#1f8c4d"},
            "documentos": {"fg_color": STYLE["secundario"], "hover_color": "#4b4b4b"},
            "modificar": {"fg_color": STYLE["primario"], "hover_color": "#D4BF22"},
            "eliminar": {"fg_color": STYLE["advertencia"], "hover_color": "#b85a52"}
        }

        # Textos alternativos para los botones (más compactos)
        BUTTON_TEXTS = {
            "descargar": "⏬",  # Alternativa a 📥
            "documentos": "📋",  # Alternativa a 📄
            "modificar": "✎",   # Alternativa a ✏️ (lápiz más simple)
            "eliminar": "🗑"    # Alternativa a ❌
        }

        # Crear filas con mejor contraste
        for i, registro in enumerate(self.historial_data):
            # Alternar colores de fondo para mejor contraste
            if i % 2 == 0:
                row_color = STYLE["surface"]
            else:
                row_color = "#f8f9fa"

            row_frame = ctk.CTkFrame(self.hist_scroll, fg_color=row_color, height=35)
            row_frame.pack(fill="x", pady=1)
            row_frame.pack_propagate(False)

            # En la sección de datos del registro:
            hora_inicio = registro.get('hora_inicio', '')  # Ya debería estar formateada
            hora_termino = registro.get('hora_termino', '')  # Ya debería estar formateada

            # Si no están formateadas, formatearlas
            if hora_inicio and ("AM" not in hora_inicio.upper() and "PM" not in hora_inicio.upper()):
                hora_inicio = self._formatear_hora_12h(hora_inicio)
            if hora_termino and ("AM" not in hora_termino.upper() and "PM" not in hora_termino.upper()):
                hora_termino = self._formatear_hora_12h(hora_termino)
                
            datos = [
                registro.get('folio_visita', '-'),
                registro.get('folio_acta', '-'),
                registro.get('fecha_inicio', '-'),
                registro.get('fecha_termino', '-'),
                hora_inicio if hora_inicio else '-',
                hora_termino if hora_termino else '-',
                registro.get('cliente', '-'),
                registro.get('nfirma1', 'No especificado'),
                registro.get('estatus', 'Completado'),
                registro.get('folios_utilizados', '0'),
                ""  # Espacio para acciones
            ]

            # Configuración de anchos (misma que headers)
            column_widths = [90, 90, 100, 100, 90, 90, 180, 120, 90, 110, 200]

            # Crear celdas
            for j, dato in enumerate(datos):
                if j == 10:  # Columna de acciones
                    acciones_frame = ctk.CTkFrame(row_frame, fg_color="transparent", width=column_widths[j])
                    acciones_frame.pack(side="left", padx=1)
                    acciones_frame.pack_propagate(False)

                    # Contenedor para botones centrados
                    botones_container = ctk.CTkFrame(acciones_frame, fg_color="transparent")
                    botones_container.pack(expand=True, fill="both", padx=5)
                    
                    # Frame para botones en horizontal
                    botones_frame = ctk.CTkFrame(botones_container, fg_color="transparent")
                    botones_frame.pack(expand=True, fill="both")

                    # Crear una función auxiliar para crear botones consistentes
                    def crear_boton_contenedor(parent, texto, comando, tipo):
                        """Crea un botón dentro de un contenedor con tamaño fijo"""
                        # Contenedor con tamaño fijo
                        container = ctk.CTkFrame(parent, fg_color="transparent", width=32, height=32)
                        container.pack(side="left", padx=1)
                        container.pack_propagate(False)
                        
                        # Crear el botón
                        btn = ctk.CTkButton(
                            container,
                            text=texto,
                            command=comando,
                            **BUTTON_CONFIG,
                            **BUTTON_COLORS[tipo]
                        )
                        
                        # Usar grid para centrar el botón
                        btn.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9, relheight=0.9)
                        return btn

                    # Crear todos los botones usando la función auxiliar
                    crear_boton_contenedor(
                        botones_frame,
                        BUTTON_TEXTS["descargar"],
                        lambda r=registro: self.descargar_folios_visita(r),
                        "descargar"
                    )

                    crear_boton_contenedor(
                        botones_frame,
                        BUTTON_TEXTS["documentos"],
                        lambda r=registro: self.mostrar_opciones_documentos(r),
                        "documentos"
                    )

                    crear_boton_contenedor(
                        botones_frame,
                        BUTTON_TEXTS["modificar"],
                        lambda r=registro: self.hist_editar_registro(r),
                        "modificar"
                    )

                    crear_boton_contenedor(
                        botones_frame,
                        BUTTON_TEXTS["eliminar"],
                        lambda r=registro: self.hist_eliminar_registro(r),
                        "eliminar"
                    )

                else:
                    # Para datos normales
                    lbl = ctk.CTkLabel(
                        row_frame,
                        text=str(dato),
                        font=("Inter", 11),
                        text_color=STYLE["texto_oscuro"],
                        width=column_widths[j],
                        anchor="center",
                        wraplength=column_widths[j]-10
                    )
                    lbl.pack(side="left", padx=1)

        # Actualizar información del pie de página
        total_registros = len(self.historial_data) if hasattr(self, 'historial_data') else 0
        self.hist_info_label.configure(text=f"Total de registros: {total_registros} - Sistema de historial de visitas - V&C")





    def mostrar_opciones_documentos(self, registro):
        """Muestra una ventana con opciones para descargar documentos"""
        # Crear ventana modal
        modal = ctk.CTkToplevel(self)
        modal.title("Descargar Documentos")
        modal.geometry("750x400")
        modal.transient(self)
        modal.grab_set()
        
        # Centrar ventana
        modal.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - modal.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - modal.winfo_height()) // 2
        modal.geometry(f"+{x}+{y}")
        
        # Frame principal
        main_frame = ctk.CTkFrame(modal, fg_color=STYLE["surface"], corner_radius=0)
        main_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Título
        ctk.CTkLabel(
            main_frame,
            text="📄 Documentos de la Visita",
            font=("Inter", 20, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=(15, 10))
        
        # Información de la visita
        info_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        info_frame.pack(fill="x", padx=20, pady=(5, 15))
        
        ctk.CTkLabel(
            info_frame,
            text=f"Folio Visita: {registro.get('folio_visita', 'N/A')} | Cliente: {registro.get('cliente', 'N/A')}",
            font=("Inter", 13),
            text_color=STYLE["texto_oscuro"]
        ).pack()
        
        # Línea separadora
        separador = ctk.CTkFrame(main_frame, fg_color=STYLE["borde"], height=1)
        separador.pack(fill="x", padx=30, pady=(0, 20))
        
        # Frame para las opciones de documentos en horizontal
        documentos_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        documentos_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Configurar grid para 3 columnas
        documentos_frame.grid_columnconfigure(0, weight=1)
        documentos_frame.grid_columnconfigure(1, weight=1)
        documentos_frame.grid_columnconfigure(2, weight=1)
        documentos_frame.grid_rowconfigure(0, weight=1)
        
        # Función para manejar la descarga de documentos
        def descargar_documento(tipo, nombre):
            modal.destroy()
            # Por ahora solo mostramos un mensaje
            messagebox.showinfo("En desarrollo", f"Función para descargar {nombre} en desarrollo.")
        
        # Botón 1: Oficio de Comisión
        oficio_frame = ctk.CTkFrame(documentos_frame, fg_color=STYLE["surface"], 
                                    border_width=1, border_color=STYLE["borde"], 
                                    corner_radius=10)
        oficio_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")
        
        # Icono grande
        ctk.CTkLabel(
            oficio_frame,
            text="📝",
            font=("Inter", 48),
            text_color=STYLE["primario"]
        ).pack(pady=(25, 15))
        
        # Nombre del documento
        ctk.CTkLabel(
            oficio_frame,
            text="OFICIO DE COMISIÓN",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=(0, 10))
        
        # Descripción
        ctk.CTkLabel(
            oficio_frame,
            text="Documento que autoriza la comisión de inspección",
            font=("Inter", 10),
            text_color=STYLE["texto_oscuro"],
            wraplength=180,
            justify="center"
        ).pack(pady=(0, 15), padx=15)
        
        # Botón de descarga - CAMBIADO: Color secundario con texto claro
        btn_oficio = ctk.CTkButton(
            oficio_frame,
            text="Descargar",
            command=lambda: descargar_documento("oficio", "Oficio de Comisión"),
            font=("Inter", 12, "bold"),
            fg_color=STYLE["secundario"],  # Cambiado a color secundario
            hover_color="#1a1a1a",  # Hover más oscuro
            text_color=STYLE["texto_claro"],  # Cambiado a texto claro
            height=35,
            corner_radius=6
        )
        btn_oficio.pack(pady=(0, 20), padx=15, fill="x")
        
        # Botón 2: Formato de Supervisión
        formato_frame = ctk.CTkFrame(documentos_frame, fg_color=STYLE["surface"], 
                                    border_width=1, border_color=STYLE["borde"], 
                                    corner_radius=10)
        formato_frame.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")
        
        # Icono grande
        ctk.CTkLabel(
            formato_frame,
            text="📊",
            font=("Inter", 48),
            text_color=STYLE["primario"]
        ).pack(pady=(25, 15))
        
        # Nombre del documento
        ctk.CTkLabel(
            formato_frame,
            text="FORMATO DE SUPERVISIÓN",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=(0, 10))
        
        # Descripción
        ctk.CTkLabel(
            formato_frame,
            text="Formato para registrar observaciones de supervisión",
            font=("Inter", 10),
            text_color=STYLE["texto_oscuro"],
            wraplength=180,
            justify="center"
        ).pack(pady=(0, 15), padx=15)
        
        # Botón de descarga - CAMBIADO: Color secundario con texto claro
        btn_formato = ctk.CTkButton(
            formato_frame,
            text="Descargar",
            command=lambda: descargar_documento("formato", "Formato de Supervisión"),
            font=("Inter", 12, "bold"),
            fg_color=STYLE["secundario"],  # Cambiado a color secundario
            hover_color="#1a1a1a",  # Hover más oscuro
            text_color=STYLE["texto_claro"],  # Cambiado a texto claro
            height=35,
            corner_radius=6
        )
        btn_formato.pack(pady=(0, 20), padx=15, fill="x")
        
        # Botón 3: Acta de Inspección
        acta_frame = ctk.CTkFrame(documentos_frame, fg_color=STYLE["surface"], 
                                border_width=1, border_color=STYLE["borde"], 
                                corner_radius=10)
        acta_frame.grid(row=0, column=2, padx=10, pady=5, sticky="nsew")
        
        # Icono grande
        ctk.CTkLabel(
            acta_frame,
            text="📋",
            font=("Inter", 48),
            text_color=STYLE["primario"]
        ).pack(pady=(25, 15))
        
        # Nombre del documento
        ctk.CTkLabel(
            acta_frame,
            text="ACTA DE INSPECCIÓN",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=(0, 10))
        
        # Descripción
        ctk.CTkLabel(
            acta_frame,
            text="Documento oficial de la visita de inspección",
            font=("Inter", 10),
            text_color=STYLE["texto_oscuro"],
            wraplength=180,
            justify="center"
        ).pack(pady=(0, 15), padx=15)
        
        # Botón de descarga - CAMBIADO: Color secundario con texto claro
        btn_acta = ctk.CTkButton(
            acta_frame,
            text="Descargar",
            command=lambda: descargar_documento("acta", "Acta de Inspección"),
            font=("Inter", 12, "bold"),
            fg_color=STYLE["secundario"],  # Cambiado a color secundario
            hover_color="#1a1a1a",  # Hover más oscuro
            text_color=STYLE["texto_claro"],  # Cambiado a texto claro
            height=35,
            corner_radius=6
        )
        btn_acta.pack(pady=(0, 20), padx=15, fill="x")
        
        # Frame para botón cerrar
        footer_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        footer_frame.pack(fill="x", pady=(20, 0))
        
    def guardar_folios_visita(self, folio_visita, datos_tabla):
        """Guarda los folios de una visita en un archivo JSON con formato 6 dígitos"""
        try:
            if not datos_tabla:
                print(f"⚠️ No hay datos de folios para guardar en la visita {folio_visita}")
                return False
            
            # Preparar datos para el archivo JSON
            folios_data = []
            
            for item in datos_tabla:
                # Obtener y formatear el folio a 6 dígitos
                folio_raw = item.get('FOLIO', '')
                folio_formateado = ""
                
                if folio_raw is not None:
                    try:
                        # Convertir a entero y formatear a 6 dígitos
                        folio_num = int(float(str(folio_raw).strip()))
                        folio_formateado = f"{folio_num:06d}"
                    except (ValueError, TypeError):
                        # Si no se puede convertir, usar el valor original
                        folio_formateado = str(folio_raw).strip()
                
                # Obtener solicitud - buscar en varias posibles columnas
                solicitud = ""
                posibles_columnas_solicitud = ['SOLICITUD', 'SOLICITUDES', 'NO. SOLICITUD']
                for col in posibles_columnas_solicitud:
                    if col in item and item[col] is not None:
                        solicitud = str(item[col]).strip()
                        break
                
                # Extraer los campos necesarios
                folio_data = {
                    "FOLIOS": folio_formateado,
                    "MARCA": str(item.get('MARCA', '')).strip() if item.get('MARCA') else "",
                    "SOLICITUDES": solicitud,
                    "FECHA DE IMPRESION": self.entry_fecha_termino.get().strip() or datetime.now().strftime("%d/%m/%Y"),
                    "FECHA DE VERIFICACION": str(item.get('FECHA DE VERIFICACION', '')).strip() if item.get('FECHA DE VERIFICACION') else "",
                    "TIPO DE DOCUMENTO": str(item.get('TIPO DE DOCUMENTO', 'D')).strip()
                }
                
                # Agregar solo si tiene folio
                if folio_data["FOLIOS"]:
                    folios_data.append(folio_data)
            
            if not folios_data:
                print(f"⚠️ No se encontraron folios válidos para guardar en la visita {folio_visita}")
                return False
            
            # Crear archivo JSON
            archivo_folios = os.path.join(self.folios_visita_path, f"folios_{folio_visita}.json")
            
            with open(archivo_folios, 'w', encoding='utf-8') as f:
                json.dump(folios_data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ Folios guardados para visita {folio_visita}: {len(folios_data)} registros")
            return True
            
        except Exception as e:
            print(f"❌ Error guardando folios para visita {folio_visita}: {e}")
            return False

    def descargar_folios_visita(self, registro):
        """Descarga los folios de una visita en formato Excel con columnas personalizadas"""
        try:
            folio_visita = registro.get('folio_visita', '')
            if not folio_visita:
                messagebox.showwarning("Error", "No se pudo obtener el folio de la visita.")
                return
            
            # Buscar el archivo JSON de folios
            archivo_folios = os.path.join(self.folios_visita_path, f"folios_{folio_visita}.json")
            
            if not os.path.exists(archivo_folios):
                messagebox.showinfo("Sin datos", f"No se encontró archivo de folios para la visita {folio_visita}.")
                return
            
            # Cargar los datos
            with open(archivo_folios, 'r', encoding='utf-8') as f:
                folios_data = json.load(f)
            
            if not folios_data:
                messagebox.showinfo("Sin datos", f"No hay datos de folios para la visita {folio_visita}.")
                return
            
            # Crear DataFrame con el orden de columnas específico
            df = pd.DataFrame(folios_data)
            
            # Definir el orden de columnas deseado
            column_order = ["FOLIOS", "MARCA", "SOLICITUDES", "FECHA DE IMPRESION", "FECHA DE VERIFICACION", "TIPO DE DOCUMENTO"]
            
            # Reordenar columnas si existen
            existing_columns = [col for col in column_order if col in df.columns]
            df = df[existing_columns]
            
            # Preguntar donde guardar el archivo Excel
            file_path = filedialog.asksaveasfilename(
                title="Guardar archivo de folios",
                defaultextension=".xlsx",
                filetypes=[
                    ("Archivos Excel", "*.xlsx"),
                    ("Archivos Excel 97-2003", "*.xls"),
                    ("Archivos CSV", "*.csv"),
                    ("Todos los archivos", "*.*")
                ],
                initialfile=f"Folios_Visita_{folio_visita}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )
            
            if not file_path:
                return
            
            # Guardar en Excel con formato
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                # Usar ExcelWriter para aplicar formato
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Folios')
                    
                    # Obtener el libro y la hoja para aplicar formato
                    workbook = writer.book
                    worksheet = writer.sheets['Folios']
                    
                    # Ajustar ancho de columnas automáticamente
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Verificar persistencia - mantener una copia en la carpeta de respaldo
            backup_dir = os.path.join(self.folios_visita_path, "backups")
            os.makedirs(backup_dir, exist_ok=True)
            
            backup_file = os.path.join(
                backup_dir, 
                f"Folios_Visita_{folio_visita}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            # Crear copia de respaldo
            try:
                if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                    shutil.copy2(file_path, backup_file)
                    print(f"📁 Copia de respaldo creada: {backup_file}")
            except Exception as backup_error:
                print(f"⚠️ No se pudo crear copia de respaldo: {backup_error}")
            
            # Mostrar información detallada
            info_mensaje = f"""
✅ Folios descargados exitosamente:

📁 Archivo: {os.path.basename(file_path)}
📋 Total de folios: {len(folios_data)}
📅 Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
📍 Ubicación: {file_path}

📊 Columnas incluidas:
   • FOLIOS (formato 6 dígitos: 000001)
   • MARCA
   • SOLICITUDES
   • FECHA DE IMPRESION
   • FECHA DE VERIFICACION
   • TIPO DE DOCUMENTO
"""
            
            messagebox.showinfo("Descarga completada", info_mensaje)
            
            # Opcional: Abrir el archivo
            respuesta = messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo descargado?")
            if respuesta:
                self._abrir_archivo(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron descargar los folios:\n{str(e)}")

    # Agregar este método para buscar la columna correcta de solicitud
    def _obtener_columna_solicitud(self, df):
        """Busca la columna correcta que contiene las solicitudes"""
        posibles_nombres = ['SOLICITUD', 'SOLICITUDES', 'NO. SOLICITUD', 'NO SOLICITUD', 'SOLICITUD NO.', 'NÚMERO DE SOLICITUD']
        
        for nombre in posibles_nombres:
            if nombre in df.columns:
                return nombre
        
        # Si no encuentra ninguna, buscar columnas que contengan "solicitud" (case insensitive)
        for col in df.columns:
            if isinstance(col, str) and 'solicitud' in col.lower():
                return col
        
        return None

    def _abrir_archivo(self, file_path):
        """Abre un archivo en el sistema operativo correspondiente"""
        try:
            if os.path.exists(file_path):
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS o Linux
                    if sys.platform == 'darwin':  # macOS
                        subprocess.Popen(['open', file_path])
                    else:  # Linux
                        subprocess.Popen(['xdg-open', file_path])
        except Exception as e:
            print(f"Error abriendo archivo: {e}")
    
    def hist_editar_registro(self, registro):
        """Abre el formulario para editar un registro del historial"""
        self._crear_formulario_visita(registro)

    def hist_buscar_general(self, event=None):
        """Buscar en el historial por cualquier campo"""
        try:
            # Asegurarse de que los datos estén cargados
            if not hasattr(self, 'historial_data') or not self.historial_data:
                self._cargar_historial()
                
            # Guardar copia original si no existe
            if not hasattr(self, 'historial_data_original') or not self.historial_data_original:
                self.historial_data_original = self.historial_data.copy()
            
            busqueda = self.entry_buscar_general.get().lower().strip()
            
            if not busqueda:
                # Si no hay búsqueda, mostrar todos los datos
                self.historial_data = self.historial_data_original.copy()
            else:
                # Filtrar datos
                resultados = []
                for registro in self.historial_data_original:
                    # Buscar en todos los campos relevantes
                    campos_busqueda = [
                        str(registro.get('folio_visita', '')),
                        str(registro.get('folio_acta', '')),
                        str(registro.get('fecha_inicio', '')),
                        str(registro.get('fecha_termino', '')),
                        str(registro.get('cliente', '')),
                        str(registro.get('estatus', '')),
                        str(registro.get('folios_utilizados', ''))  # Cambiado de 'folios_usados' a 'folios_utilizados'
                    ]
                    
                    # Verificar si la búsqueda coincide con algún campo
                    if any(busqueda in campo.lower() for campo in campos_busqueda):
                        resultados.append(registro)
                
                self.historial_data = resultados
            
            self._poblar_historial_ui()
            
        except Exception as e:
            print(f"Error en búsqueda general: {e}")

    def hist_limpiar_busqueda(self):
        """Limpiar todas las búsquedas y mostrar todos los registros"""
        self.entry_buscar_general.delete(0, 'end')
        self.entry_buscar_folio.delete(0, 'end')
        
        # Recargar datos originales
        if hasattr(self, 'historial_data_original'):
            self.historial_data = self.historial_data_original.copy()
        else:
            self._cargar_historial()
            
        self._poblar_historial_ui()

    def hist_eliminar_registro(self, registro):
        """Eliminar un registro del historial"""
        try:
            folio = registro.get('folio_visita', '')
            confirmacion = messagebox.askyesno(
                "Confirmar eliminación", 
                f"¿Está seguro de que desea eliminar el registro del folio {folio}?"
            )
            
            if confirmacion:
                # Eliminar del historial_data
                self.historial_data = [r for r in self.historial_data if r != registro]
                self.historial_data_original = [r for r in self.historial_data_original if r != registro]
                
                # Actualizar el archivo JSON
                if os.path.exists(self.historial_path):
                    with open(self.historial_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    # Filtrar las visitas
                    data['visitas'] = [v for v in data.get('visitas', []) if v != registro]
                    
                    # Guardar los cambios
                    with open(self.historial_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                
                # Actualizar la UI
                self._poblar_historial_ui()
                messagebox.showinfo("Éxito", f"Registro del folio {folio} eliminado correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar el registro:\n{e}")

    def hist_create_visita(self, payload, es_automatica=False):
        """Crea una nueva visita en el historial"""
        try:
            # Generar ID único
            payload["_id"] = str(uuid.uuid4())
            
            # Asegurar que estatus tenga valor
            payload.setdefault("estatus", "Completada" if es_automatica else "En proceso")
            
            # Asegurar que las fechas y horas estén presentes
            payload.setdefault("fecha_inicio", "")
            payload.setdefault("fecha_termino", "")
            payload.setdefault("hora_inicio", "")
            payload.setdefault("hora_termino", "")
            
            # Agregar a la lista
            if "visitas" not in self.historial:
                self.historial["visitas"] = []
            self.historial["visitas"].append(payload)
            
            # Actualizar datos
            self.historial_data = self.historial["visitas"]
            
            self._guardar_historial()
            self._poblar_historial_ui()
            
            if not es_automatica:
                messagebox.showinfo("OK", f"Visita {payload['folio_visita']} guardada correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def hist_buscar_por_folio(self):
        """Buscar en el historial por folio de visita"""
        try:
            folio_busqueda = self.entry_buscar_folio.get().strip()
            
            if not folio_busqueda:
                # Si no hay búsqueda, mostrar todos los datos
                self.historial_data = self.historial_data_original.copy() if hasattr(self, 'historial_data_original') else self.historial_data
            else:
                # Filtrar datos por folio
                resultados = []
                for registro in (self.historial_data_original if hasattr(self, 'historial_data_original') else self.historial_data):
                    folio_actual = str(registro.get('folio_visita', ''))
                    if folio_busqueda.lower() in folio_actual.lower():
                        resultados.append(registro)
                
                self.historial_data = resultados
            
            self._poblar_historial_ui()
            
        except Exception as e:
            print(f"Error en búsqueda por folio: {e}")

    def hist_update_visita(self, id_, nuevos):
        """Actualiza una visita existente"""
        try:
            # Buscar la visita a actualizar
            for i, v in enumerate(self.historial.get("visitas", [])):
                if v["_id"] == id_:
                    # Mantener el ID y folio original
                    nuevos["_id"] = id_
                    nuevos["folio_visita"] = v.get("folio_visita", nuevos.get("folio_visita"))
                    
                    # Actualizar en el historial
                    self.historial["visitas"][i] = nuevos
                    
                    # Actualizar datos
                    self.historial_data = self.historial["visitas"]
                    
                    self._guardar_historial()
                    self._poblar_historial_ui()
                    messagebox.showinfo("OK", f"Visita {nuevos['folio_visita']} actualizada")
                    return
                    
            messagebox.showerror("Error", "No se encontró la visita para actualizar")
            
        except Exception as e:
            messagebox.showerror("Error actualizando", str(e))

    def registrar_visita_automatica(self, resultado_dictamenes):
        """Registra automáticamente una visita al generar dictámenes con información de folios"""
        try:
            if not self.cliente_seleccionado:
                return

            # Obtener datos del formulario
            folio_visita = self.entry_folio_visita.get().strip()
            folio_acta = self.entry_folio_acta.get().strip()
            fecha_inicio = self.entry_fecha_inicio.get().strip()
            fecha_termino = self.entry_fecha_termino.get().strip()
            hora_inicio = self.entry_hora_inicio.get().strip()
            hora_termino = self.entry_hora_termino.get().strip()
            supervisor = self.entry_supervisor.get().strip()

            # Convertir horas a formato consistente (24h para almacenamiento)
            def estandarizar_hora_24h(hora_str):
                """Estandariza hora a formato 24h HH:MM"""
                if not hora_str or hora_str.strip() == "":
                    return ""
                
                try:
                    hora_str = str(hora_str).strip()
                    # Reemplazar punto por dos puntos
                    hora_str = hora_str.replace(".", ":")
                    
                    if ":" in hora_str:
                        partes = hora_str.split(":")
                        hora = int(partes[0].strip())
                        minutos = partes[1].strip()[:2]
                        
                        # Asegurar rango válido
                        if hora < 0 or hora > 23:
                            hora = 0
                        
                        # Formatear a 2 dígitos
                        return f"{hora:02d}:{minutos}"
                    else:
                        return hora_str
                except:
                    return hora_str
            
            # Estandarizar horas a 24h
            hora_inicio_24h = estandarizar_hora_24h(hora_inicio)
            hora_termino_24h = estandarizar_hora_24h(hora_termino)
            
            # Formatear horas a 12h para visualización
            hora_inicio_formateada = self._formatear_hora_12h(hora_inicio_24h) if hora_inicio_24h else ""
            hora_termino_formateada = self._formatear_hora_12h(hora_termino_24h) if hora_termino_24h else ""

            # Si no hay fecha/hora de término, usar la actual
            if not fecha_termino:
                fecha_termino = datetime.now().strftime("%d/%m/%Y")
            if not hora_termino_24h:
                hora_termino_24h = datetime.now().strftime("%H:%M")
                hora_termino_formateada = self._formatear_hora_12h(hora_termino_24h)

            # OBTENER INFORMACIÓN DE FOLIOS UTILIZADOS
            folios_utilizados = self._obtener_folios_de_tabla()

            # CARGAR DATOS DE TABLA DE RELACIÓN SI EXISTEN
            datos_tabla = []
            if self.archivo_json_generado and os.path.exists(self.archivo_json_generado):
                with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                    datos_tabla = json.load(f)
                    
                # Guardar folios específicos para esta visita
                self.guardar_folios_visita(folio_visita, datos_tabla)

            # Crear payload para visita automática con información de folios
            # GUARDAR HORAS EN FORMATO 24h Y TAMBIÉN FORMATEADAS PARA VISUALIZACIÓN
            payload = {
                "folio_visita": folio_visita,
                "folio_acta": folio_acta or f"AC{self.current_folio}",
                "fecha_inicio": fecha_inicio or datetime.now().strftime("%d/%m/%Y"),
                "fecha_termino": fecha_termino,
                "hora_inicio_24h": hora_inicio_24h or datetime.now().strftime("%H:%M"),
                "hora_termino_24h": hora_termino_24h or datetime.now().strftime("%H:%M"),
                "hora_inicio": hora_inicio_formateada or self._formatear_hora_12h(datetime.now().strftime("%H:%M")),
                "hora_termino": hora_termino_formateada,
                "norma": "",
                "cliente": self.cliente_seleccionado['CLIENTE'],
                "nfirma1": supervisor or " ",  # Supervisor
                "nfirma2": "",
                "estatus": "Completada",
                "folios_utilizados": folios_utilizados,
                "total_folios": len(datos_tabla) if datos_tabla else 0
            }

            # Guardar visita automática
            self.hist_create_visita(payload, es_automatica=True)
            
            # Preparar nueva visita después de guardar
            self.crear_nueva_visita()
            
        except Exception as e:
            print(f"⚠️ Error registrando visita automática: {e}")

    def limpiar_archivo(self):
        self.archivo_excel_cargado = None
        self.archivo_json_generado = None
        self.json_filename = None
        
        # Limpiar también la información de folios
        if hasattr(self, 'info_folios_actual'):
            del self.info_folios_actual

        self.info_archivo.configure(
            text="No se ha cargado ningún archivo",
            text_color=STYLE["texto_claro"]
        )

        self.boton_cargar_excel.configure(state="normal")
        self.boton_limpiar.configure(state="disabled")
        self.boton_generar_dictamen.configure(state="disabled")

        self.etiqueta_estado.configure(text="", text_color=STYLE["texto_claro"])
        self.check_label.configure(text="")
        self.barra_progreso.set(0)
        self.etiqueta_progreso.configure(text="")

        try:
            data_dir = os.path.join(os.path.dirname(__file__), "data")
            
            # Archivos a eliminar (pero NO los de folios_visitas)
            archivos_a_eliminar = [
                "base_etiquetado.json",
                "tabla_de_relacion.json"
            ]
            
            archivos_eliminados = []
            
            for archivo in archivos_a_eliminar:
                ruta_archivo = os.path.join(data_dir, archivo)
                if os.path.exists(ruta_archivo):
                    os.remove(ruta_archivo)
                    archivos_eliminados.append(archivo)
                    print(f"🗑️ {archivo} eliminado correctamente.")
            
            if archivos_eliminados:
                print(f"✅ Se eliminaron {len(archivos_eliminados)} archivos: {', '.join(archivos_eliminados)}")
            else:
                print("ℹ️ No se encontraron archivos para eliminar.")

            self.archivo_etiquetado_json = None
            self.info_etiquetado.configure(text="")
            self.info_etiquetado.pack_forget()

        except Exception as e:
            print(f"⚠️ Error al eliminar archivos: {e}")

        messagebox.showinfo("Limpieza completa", "Los datos del archivo y el etiquetado han sido limpiados.\n\nNota: Los archivos de folios por visita se conservan en la carpeta 'folios_visitas'.")

    def _crear_formulario_visita(self, datos=None):
        """Crea un formulario modal para editar visitas con disposición organizada"""
        datos = datos or {}
        modal = ctk.CTkToplevel(self)
        modal.title("Editar Visita")
        modal.geometry("900x550")  # Aumentado altura para mejor visibilidad
        modal.transient(self)
        modal.grab_set()

        # Centrar ventana
        modal.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - modal.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - modal.winfo_height()) // 2
        modal.geometry(f"+{x}+{y}")
        
        # Frame principal
        main_frame = ctk.CTkFrame(modal, fg_color=STYLE["surface"], corner_radius=0)
        main_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Título
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.pack(fill="x", padx=25, pady=(20, 10))
        
        ctk.CTkLabel(
            title_frame,
            text="✏️ Editar Visita",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w")
        
        # Línea separadora
        separador = ctk.CTkFrame(main_frame, fg_color=STYLE["borde"], height=1)
        separador.pack(fill="x", padx=25, pady=(0, 10))
        
        # Frame para contenido principal con scroll
        content_scroll = ctk.CTkScrollableFrame(
            main_frame, 
            fg_color="transparent",
            scrollbar_button_color=STYLE["primario"],
            scrollbar_button_hover_color=STYLE["primario"],
            height=350
        )
        content_scroll.pack(fill="both", expand=True, padx=25, pady=(5, 10))
        
        # Frame para contenido en grid (3 columnas para mejor organización)
        content_frame = ctk.CTkFrame(content_scroll, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # Configurar 3 columnas
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_columnconfigure(2, weight=1)
        
        entries = {}
        
        # Definir campos organizados por columnas
        campos_por_columna = [
            [  # Columna 1: Información básica
                ("folio_visita", "Folio Visita"),
                ("folio_acta", "Folio Acta"),
                ("norma", "Norma"),
            ],
            [  # Columna 2: Fechas y horas
                ("fecha_inicio", "Fecha Inicio"),
                ("fecha_termino", "Fecha Termino"),
                ("hora_inicio", "Hora Inicio"),
                ("hora_termino", "Hora Termino"),
            ],
            [  # Columna 3: Cliente y estatus
                ("cliente", "Cliente"),
                ("estatus", "Estatus"),
                ("folios_utilizados", "Folios Utilizados"),
                ("nfirma1", "Nombre Supervisor"),
            ]
        ]
        
        # Crear campos para cada columna
        for col_idx, campos in enumerate(campos_por_columna):
            col_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            col_frame.grid(row=0, column=col_idx, padx=10, pady=0, sticky="nsew")
            
            for i, (key, label) in enumerate(campos):
                field_frame = ctk.CTkFrame(col_frame, fg_color="transparent")
                field_frame.pack(fill="x", pady=(0, 12))
                
                ctk.CTkLabel(
                    field_frame, 
                    text=label, 
                    anchor="w", 
                    font=FONT_SMALL,
                    text_color=STYLE["texto_oscuro"]
                ).pack(anchor="w", pady=(0, 5))
                
                if key == "cliente":
                    # Obtener lista de clientes
                    clientes_lista = ["Seleccione un cliente..."]
                    if hasattr(self, 'clientes_data') and self.clientes_data:
                        clientes_lista.extend([cliente['CLIENTE'] for cliente in self.clientes_data])
                    
                    # Crear combobox para clientes
                    ent = ctk.CTkComboBox(
                        field_frame,
                        values=clientes_lista,
                        font=FONT_SMALL,
                        dropdown_font=FONT_SMALL,
                        state="readonly",
                        height=35,
                        corner_radius=8,
                        width=250
                    )
                    ent.pack(fill="x")
                    
                    # Establecer cliente si existe en datos
                    if datos and "cliente" in datos:
                        cliente_actual = datos.get("cliente", "")
                        if cliente_actual in clientes_lista:
                            ent.set(cliente_actual)
                        else:
                            ent.set("Seleccione un cliente...")
                    else:
                        ent.set("Seleccione un cliente...")
                        
                elif key == "estatus":
                    # Combobox para estatus
                    ent = ctk.CTkComboBox(
                        field_frame,
                        values=["En proceso", "Completada", "Cancelada", "Pendiente"],
                        font=FONT_SMALL,
                        dropdown_font=FONT_SMALL,
                        state="readonly",
                        height=35,
                        corner_radius=8,
                        width=250
                    )
                    ent.pack(fill="x")
                    
                    if datos and "estatus" in datos:
                        ent.set(datos.get("estatus", "En proceso"))
                    else:
                        ent.set("En proceso")
                        
                else:
                    # Campo de texto normal
                    ent = ctk.CTkEntry(
                        field_frame, 
                        height=35,
                        corner_radius=8,
                        font=FONT_SMALL,
                        placeholder_text=f"Ingrese {label.lower()}" if key not in ["hora_inicio", "hora_termino"] else "HH:MM"
                    )
                    ent.pack(fill="x")
                    
                    # Insertar datos si existen
                    if datos and key in datos:
                        ent.insert(0, str(datos.get(key, "")))
                
                entries[key] = ent
        
        # Frame para botones
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 20), padx=25)
        
        def _guardar():
            # Recoger datos de todos los campos
            payload = {}
            for key, entry in entries.items():
                if key in ["cliente", "estatus"]:
                    # Para combobox, obtener el valor seleccionado
                    value = entry.get()
                    if key == "cliente" and value == "Seleccione un cliente...":
                        value = ""
                else:
                    value = entry.get().strip()
                payload[key] = value
            
            # Validaciones
            if not payload.get("cliente"):
                messagebox.showwarning("Validación", "Por favor seleccione un cliente")
                return
            
            if not payload.get("estatus"):
                payload["estatus"] = "En proceso"
            
            # Actualizar la visita
            self.hist_update_visita(datos["_id"], payload)
            modal.destroy()
        
        # Botones mejorados
        ctk.CTkButton(
            btn_frame, 
            text="Cancelar", 
            command=modal.destroy,
            font=("Inter", 13),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["texto_claro"],
            height=38,
            width=130,
            corner_radius=8
        ).pack(side="right", padx=(8, 0))
        
        ctk.CTkButton(
            btn_frame, 
            text="Guardar Cambios", 
            command=_guardar,
            font=("Inter", 13, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=38,
            width=150,
            corner_radius=8
        ).pack(side="right")
        
        # Agregar un pequeño espaciador para empujar botones a la derecha
        ctk.CTkLabel(btn_frame, text="", fg_color="transparent").pack(side="left", expand=True)

    # -----------------------------------------------------------
    # NUEVOS MÉTODOS PARA DIAGNÓSTICO Y LIMPIEZA
    # -----------------------------------------------------------
    def verificar_integridad_datos(self):
        """Verifica la integridad de los datos cargados"""
        try:
            if not self.archivo_json_generado:
                messagebox.showwarning("Sin datos", "No hay archivo cargado para verificar")
                return
            
            with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            
            # Contar valores únicos en LISTA para determinar familias/dictámenes
            listas_unicas = set()
            for item in datos:
                if 'LISTA' in item and item['LISTA'] is not None:
                    # Convertir a string y eliminar espacios para consistencia
                    lista_valor = str(item['LISTA']).strip()
                    if lista_valor:  # Solo agregar si no está vacío
                        listas_unicas.add(lista_valor)
            
            # Si no hay campo LISTA, contar cada registro como un dictamen
            if not listas_unicas:
                total_dictamenes = len(datos)
            else:
                total_dictamenes = len(listas_unicas)
            
            # Verificar duplicados por ID único
            ids_vistos = set()
            duplicados = 0
            for item in datos:
                item_id = str(item.get('ID', '') or str(item.get('FOLIO', '')) or str(item))
                if item_id in ids_vistos:
                    duplicados += 1
                ids_vistos.add(item_id)
            
            # Mostrar reporte
            reporte = f"""
    📊 REPORTE DE INTEGRIDAD DE DATOS

    📁 Total de registros: {len(datos)}
    📋 Dictámenes que generara el sistema: {total_dictamenes}
            """
            
            messagebox.showinfo("Reporte de Integridad", reporte)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo verificar integridad:\n{e}")

# ================== EJECUCIÓN ================== #
if __name__ == "__main__":
    app = SistemaDictamenesVC()
    app.mainloop()