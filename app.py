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
    "texto_claro": "#4b4b4b",
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
        self.geometry("1200x600")
        self.minsize(1200, 600)
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

        # ===== HEADER =====
        self.crear_header()

        # ===== CONTENEDOR PRINCIPAL CON PESTAÑAS =====
        self.crear_contenido_principal()

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

    def crear_header(self):
        header = ctk.CTkFrame(self, fg_color=STYLE["fondo"], corner_radius=0, height=60)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(expand=True, fill="both", padx=25, pady=15)

        ctk.CTkLabel(
            header_content,
            text="Generador de Dictámenes",
            font=FONT_TITLE,
            text_color=STYLE["secundario"]
        ).pack(anchor="center", expand=True, fill="both", pady=(0, 5))

    def crear_contenido_principal(self):
        """Contenido principal con pestañas"""
        main_container = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
        main_container.pack(fill="both", expand=True, padx=20, pady=15)

        # Tabview con pestañas
        self.tabs = ctk.CTkTabview(main_container, width=860, height=480)
        self.tabs.pack(fill="both", expand=True)

        # Pestaña principal
        self.tabs.add("Principal")
        self._construir_tab_principal(self.tabs.tab("Principal"))

        # Pestaña historial
        self.tabs.add("Historial")
        self._construir_tab_historial(self.tabs.tab("Historial"))

    def _construir_tab_principal(self, parent):
        """Construye la interfaz principal con dos tarjetas de tamaño fijo"""
        # ===== CONTENEDOR PRINCIPAL CON 2 COLUMNAS =====
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True)

        # Configurar grid para 2 columnas con tamaño fijo
        main_frame.grid_columnconfigure(0, weight=1, minsize=430)
        main_frame.grid_columnconfigure(1, weight=1, minsize=430)
        main_frame.grid_rowconfigure(0, weight=1)

        # ===== TARJETA INFORMACIÓN DE VISITA (IZQUIERDA) - TAMAÑO FIJO =====
        card_visita = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12, height=450)
        card_visita.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="nsew")
        card_visita.grid_propagate(False)  # Tamaño fijo

        ctk.CTkLabel(
            card_visita,
            text="📋 Información de Visita",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 15))

        visita_frame = ctk.CTkFrame(card_visita, fg_color="transparent")
        visita_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor scrollable para el formulario
        scroll_form = ctk.CTkScrollableFrame(visita_frame, height=350, fg_color="transparent")
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

        # ===== TARJETA GENERACIÓN (DERECHA) - TAMAÑO FIJO =====
        card_generacion = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12, height=450)
        card_generacion.grid(row=0, column=1, padx=(10, 0), pady=0, sticky="nsew")
        card_generacion.grid_propagate(False)  # Tamaño fijo

        ctk.CTkLabel(
            card_generacion,
            text="🚀 Generación de Dictámenes",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 15))

        generacion_frame = ctk.CTkFrame(card_generacion, fg_color="transparent")
        generacion_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor scrollable para la generación
        scroll_generacion = ctk.CTkScrollableFrame(generacion_frame, height=350, fg_color="transparent")
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

        self.info_etiquetado = ctk.CTkLabel(
            botones_carga_frame,
            text="",
            font=FONT_SMALL,
            text_color=STYLE["texto_claro"],
            wraplength=350
        )

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

        # Botón para verificar integridad de datos
        botones_extra_frame = ctk.CTkFrame(generar_section, fg_color="transparent")
        botones_extra_frame.pack(fill="x", pady=(5, 0))
        
        ctk.CTkButton(
            botones_extra_frame,
            text="🔍 Verificar Datos",
            command=self.verificar_integridad_datos,
            font=("Inter", 11),
            fg_color=STYLE["advertencia"],
            hover_color="#b85a52",
            text_color=STYLE["surface"],
            height=30,
            width=100,
            corner_radius=6
        ).pack(side="left", padx=(0, 8))
        
        ctk.CTkButton(
            botones_extra_frame,
            text="🧹 Limpiar Cache",
            command=self.limpiar_cache_completo,
            font=("Inter", 11),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=30,
            width=100,
            corner_radius=6
        ).pack(side="left")

    def _construir_tab_historial(self, parent):
        """Construye la pestaña de historial con columnas mejor organizadas"""
        cont = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=8)
        cont.pack(fill="both", expand=True, padx=12, pady=12)

        # ===== BARRA SUPERIOR CON BUSCADOR =====
        barra_superior_historial = ctk.CTkFrame(cont, fg_color="transparent", height=50)
        barra_superior_historial.pack(fill="x", pady=(0, 10))
        barra_superior_historial.pack_propagate(False)

        # Buscador
        buscador_frame = ctk.CTkFrame(barra_superior_historial, fg_color="transparent")
        buscador_frame.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            buscador_frame,
            text="Buscar folio visita:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(0, 8))

        self.entry_buscar_folio = ctk.CTkEntry(
            buscador_frame,
            placeholder_text="Folio visita...",
            width=180,
            height=35,
            corner_radius=8
        )
        self.entry_buscar_folio.pack(side="left", padx=(0, 8))

        btn_buscar = ctk.CTkButton(
            buscador_frame,
            text="Buscar",
            command=self.hist_buscar_por_folio,
            font=("Inter", 12, "bold"),
            fg_color=STYLE["secundario"],
            hover_color="#1a1a1a",
            text_color=STYLE["surface"],
            height=35,
            width=80,
            corner_radius=8
        )
        btn_buscar.pack(side="left")

        # ===== TABLA MEJORADA =====
        tabla_container = ctk.CTkFrame(cont, fg_color=STYLE["fondo"], corner_radius=8)
        tabla_container.pack(fill="both", expand=True, pady=(0, 10))

        # Encabezados de la tabla
        header_frame = ctk.CTkFrame(tabla_container, fg_color=STYLE["secundario"], height=40)
        header_frame.pack(fill="x", padx=0, pady=(0, 2))
        header_frame.pack_propagate(False)

        # Configuración de anchos fijos para cada columna
        column_widths = [90, 90, 100, 100, 80, 80, 200, 100, 120, 100]  # Suma: 1060

        headers = [
            "Folio Visita",
            "Folio Acta", 
            "Fecha Inicio",
            "Fecha Termino",
            "Hora Inicio",
            "Hora Termino",
            "Cliente",
            "Estatus",
            "Folios Usados",
            "Acciones"
        ]

        # Crear headers
        for i, header_text in enumerate(headers):
            lbl = ctk.CTkLabel(
                header_frame, 
                text=header_text, 
                font=FONT_SMALL,
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
            scrollbar_button_hover_color=STYLE["secundario"]
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

            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)

            records = df.to_dict(orient="records")

            data_folder = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_folder, exist_ok=True)

            self.json_filename = "tabla_de_relacion.json"
            output_path = os.path.join(data_folder, self.json_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(records, f, ensure_ascii=False, indent=2)

            self.after(0, self._actualizar_ui_conversion_exitosa, output_path, len(records))

        except Exception as e:
            self.after(0, self.mostrar_error, f"Error al convertir el archivo:\n{e}")

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
            confirmacion = messagebox.askyesno(
                "Generar Dictámenes",
                f"¿Está seguro de que desea generar los dictámenes PDF?\n\n"
                f"📄 Archivo: {os.path.basename(self.archivo_json_generado)}\n"
                f"👤 Cliente: {self.cliente_seleccionado['CLIENTE']}\n"
                f"📋 RFC: {self.cliente_seleccionado.get('RFC', 'No disponible')}"
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
        try:
            os.makedirs(os.path.join(os.path.dirname(__file__), "data"), exist_ok=True)
            if not os.path.exists(self.historial_path):
                base = {
                    "last_folio_visita": 0,
                    "visitas": []
                }
                with open(self.historial_path, "w", encoding="utf-8") as f:
                    json.dump(base, f, ensure_ascii=False, indent=2)
                self.historial = base
            else:
                with open(self.historial_path, "r", encoding="utf-8") as f:
                    self.historial = json.load(f)
            self.historial.setdefault("last_folio_visita", 0)
            self.historial.setdefault("visitas", [])
        except Exception as e:
            print(f"❌ Error cargando historial: {e}")
            self.historial = {"last_folio_visita": 0, "visitas": []}

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
            with open(self.historial_path, "w", encoding="utf-8") as f:
                json.dump(self.historial, f, ensure_ascii=False, indent=2)
            self.hist_info_label.configure(text=f"Guardado OK — {len(self.historial.get('visitas',[]))} visitas")
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

    def _poblar_historial_ui(self, filtro_text=None):
        """Puebla la lista del historial con columnas perfectamente alineadas"""
        self._limpiar_scroll_hist()
        visitas = self.historial.get("visitas", [])
        
        if filtro_text:
            filtro_text = str(filtro_text).lower()
            visitas = [v for v in visitas if filtro_text in str(v.get("folio_visita","")).lower() 
                      or filtro_text in str(v.get("cliente","")).lower()]

        if not visitas:
            # Mostrar mensaje cuando no hay registros
            empty_frame = ctk.CTkFrame(self.hist_scroll, fg_color=STYLE["surface"], height=60)
            empty_frame.pack(fill="x", padx=8, pady=20)
            
            ctk.CTkLabel(
                empty_frame,
                text="📭 No hay registros de visitas",
                font=FONT_SMALL,
                text_color=STYLE["texto_claro"]
            ).pack(expand=True, pady=20)
            return

        # Configuración de anchos fijos (debe coincidir con los headers)
        column_widths = [90, 90, 100, 100, 80, 80, 200, 100, 120, 100]

        for v in visitas:
            row = ctk.CTkFrame(self.hist_scroll, fg_color=STYLE["surface"], height=45)
            row.pack(fill="x", padx=2, pady=2)
            
            # Folio Visita
            lbl_folio = ctk.CTkLabel(
                row, 
                text=str(v.get("folio_visita", "")), 
                font=FONT_SMALL,
                width=column_widths[0],
                anchor="center"
            )
            lbl_folio.pack(side="left", padx=1)
            
            # Folio Acta
            lbl_acta = ctk.CTkLabel(
                row, 
                text=str(v.get("folio_acta", "")), 
                font=FONT_SMALL,
                width=column_widths[1],
                anchor="center"
            )
            lbl_acta.pack(side="left", padx=1)
            
            # Fecha Inicio
            lbl_fecha_inicio = ctk.CTkLabel(
                row, 
                text=str(v.get("fecha_inicio", "")), 
                font=FONT_SMALL,
                width=column_widths[2],
                anchor="center"
            )
            lbl_fecha_inicio.pack(side="left", padx=1)
            
            # Fecha Termino
            lbl_fecha_termino = ctk.CTkLabel(
                row, 
                text=str(v.get("fecha_termino", "")), 
                font=FONT_SMALL,
                width=column_widths[3],
                anchor="center"
            )
            lbl_fecha_termino.pack(side="left", padx=1)
            
            # Hora Inicio
            lbl_hora_inicio = ctk.CTkLabel(
                row, 
                text=str(v.get("hora_inicio", "")), 
                font=FONT_SMALL,
                width=column_widths[4],
                anchor="center"
            )
            lbl_hora_inicio.pack(side="left", padx=1)
            
            # Hora Termino
            lbl_hora_termino = ctk.CTkLabel(
                row, 
                text=str(v.get("hora_termino", "")), 
                font=FONT_SMALL,
                width=column_widths[5],
                anchor="center"
            )
            lbl_hora_termino.pack(side="left", padx=1)
            
            # Cliente (con texto truncado si es muy largo)
            cliente_text = str(v.get("cliente", ""))
            if len(cliente_text) > 25:
                cliente_text = cliente_text[:22] + "..."
            lbl_cliente = ctk.CTkLabel(
                row, 
                text=cliente_text, 
                font=FONT_SMALL,
                width=column_widths[6],
                anchor="center"
            )
            lbl_cliente.pack(side="left", padx=1)
            
            # Estatus con color según el estado
            estatus = str(v.get("estatus", ""))
            estatus_color = STYLE["exito"] if estatus == "Completada" else STYLE["advertencia"]
            lbl_status = ctk.CTkLabel(
                row, 
                text=estatus, 
                font=FONT_SMALL,
                text_color=estatus_color,
                width=column_widths[7],
                anchor="center"
            )
            lbl_status.pack(side="left", padx=1)

            # Folios Utilizados
            folios_utilizados = str(v.get("folios_utilizados", "N/A"))
            if len(folios_utilizados) > 15:
                folios_utilizados = folios_utilizados[:12] + "..."
            lbl_folios = ctk.CTkLabel(
                row, 
                text=folios_utilizados, 
                font=FONT_SMALL,
                width=column_widths[8],
                anchor="center",
                text_color=STYLE["texto_claro"]
            )
            lbl_folios.pack(side="left", padx=1)

            # Acciones
            actions_frame = ctk.CTkFrame(row, fg_color="transparent", width=column_widths[9])
            actions_frame.pack(side="left", padx=1)
            actions_frame.pack_propagate(False)
            
            btn_edit = ctk.CTkButton(
                actions_frame, 
                text="✏️", 
                width=32, 
                height=28,
                font=("Arial", 12),
                fg_color=STYLE["primario"],
                hover_color="#D4BF22",
                text_color=STYLE["secundario"],
                command=lambda _id=v["_id"]: self.hist_editar_visita(_id)
            )
            btn_edit.pack(side="left", padx=1)
            
            btn_del = ctk.CTkButton(
                actions_frame, 
                text="🗑️", 
                width=32, 
                height=28,
                font=("Arial", 12),
                fg_color=STYLE["peligro"], 
                hover_color="#c13a2d", 
                command=lambda _id=v["_id"]: self.hist_borrar_visita(_id)
            )
            btn_del.pack(side="left", padx=1)

        # Actualizar información del pie de página
        self.hist_info_label.configure(
            text=f"Total de visitas: {len(visitas)} | Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )

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
            self.historial.setdefault("visitas", []).append(payload)
            self._guardar_historial()
            self._poblar_historial_ui()
            
            if not es_automatica:
                messagebox.showinfo("OK", f"Visita {payload['folio_visita']} guardada correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def hist_buscar_por_folio(self):
        clave = self.entry_buscar_folio.get().strip()
        if not clave:
            self._poblar_historial_ui()
            return
        self._poblar_historial_ui(filtro_text=clave)

    def hist_editar_visita(self, id_):
        visita = next((v for v in self.historial.get("visitas",[]) if v["_id"]==id_), None)
        if not visita:
            messagebox.showerror("No encontrado", "Registro no encontrado")
            return
        self._crear_formulario_visita(visita)

    def hist_update_visita(self, id_, nuevos):
        try:
            for i, v in enumerate(self.historial.get("visitas",[])):
                if v["_id"] == id_:
                    nuevos["_id"] = id_
                    nuevos["folio_visita"] = v.get("folio_visita", nuevos.get("folio_visita"))
                    self.historial["visitas"][i] = nuevos
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

            # Si no hay fecha/hora de término, usar la actual
            if not fecha_termino:
                fecha_termino = datetime.now().strftime("%d/%m/%Y")
            if not hora_termino:
                hora_termino = datetime.now().strftime("%H:%M")

            # OBTENER INFORMACIÓN DE FOLIOS UTILIZADOS
            folios_utilizados = resultado_dictamenes.get('folios_utilizados_info', 'No disponible')
            
            # Si no está disponible en el resultado, obtener de la tabla
            if folios_utilizados == 'No disponible':
                folios_utilizados = self._obtener_folios_de_tabla()

            # Crear payload para visita automática con información de folios
            payload = {
                "folio_visita": folio_visita,
                "folio_acta": folio_acta or f"AC{self.current_folio}",
                "fecha_inicio": fecha_inicio or datetime.now().strftime("%d/%m/%Y"),
                "fecha_termino": fecha_termino,
                "hora_inicio": hora_inicio or datetime.now().strftime("%H:%M"),
                "hora_termino": hora_termino,
                "norma": "VARIAS",
                "cliente": self.cliente_seleccionado['CLIENTE'],
                "nfirma1": supervisor or "Sistema Automático",
                "nfirma2": "",
                "estatus": "Completada",
                "folios_utilizados": folios_utilizados  # NUEVO CAMPO CON INFORMACIÓN DE FOLIOS
            }

            # Guardar visita automática
            self.hist_create_visita(payload, es_automatica=True)
            
            # Preparar nueva visita después de guardar
            self.crear_nueva_visita()
            
        except Exception as e:
            print(f"⚠️ Error registrando visita automática: {e}")

    def _obtener_folios_de_tabla(self):
        """Obtiene la información de folios de la tabla de relación"""
        try:
            if not self.archivo_json_generado:
                return "No disponible"
            
            with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                datos_tabla = json.load(f)
            
            # Buscar la columna FOLIOS en los datos
            folios_encontrados = []
            for item in datos_tabla:
                if 'FOLIOS' in item and item['FOLIOS'] and str(item['FOLIOS']).strip():
                    folios_encontrados.append(str(item['FOLIOS']).strip())
            
            if folios_encontrados:
                return ", ".join(folios_encontrados)
            else:
                return "No se encontraron folios en la tabla"
                
        except Exception as e:
            print(f"⚠️ Error leyendo folios de tabla: {e}")
            return f"Error: {str(e)}"

    def _crear_formulario_visita(self, datos=None):
        """Crea un formulario modal para editar visitas con todos los campos"""
        datos = datos or {}
        modal = ctk.CTkToplevel(self)
        modal.title("Editar Visita")
        modal.geometry("500x700")  # Un poco más alto para incluir el nuevo campo
        modal.transient(self)
        modal.grab_set()

        form = ctk.CTkFrame(modal, fg_color=STYLE["surface"])
        form.pack(fill="both", expand=True, padx=12, pady=12)

        ctk.CTkLabel(
            form,
            text="Editar visita",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 15))

        campos = [
            ("folio_visita", "Folio Visita"),
            ("folio_acta", "Folio Acta"),
            ("fecha_inicio", "Fecha Inicio (dd/mm/yyyy)"),
            ("fecha_termino", "Fecha Termino (dd/mm/yyyy)"),
            ("hora_inicio", "Hora Inicio (HH:MM)"),
            ("hora_termino", "Hora Termino (HH:MM)"),
            ("norma", "Norma"),
            ("cliente", "Cliente"),
            ("nfirma1", "Nombre Supervisor"),
            ("nfirma2", "Nombre Inspector"),
            ("estatus", "Estatus"),
            ("folios_utilizados", "Folios Utilizados")  # NUEVO CAMPO
        ]
        entries = {}
        
        scroll_frame = ctk.CTkScrollableFrame(form, height=550)  # Más alto para el nuevo campo
        scroll_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        r = 0
        for key, label in campos:
            ctk.CTkLabel(scroll_frame, text=label, anchor="w", font=FONT_SMALL).grid(row=r, column=0, sticky="w", padx=8, pady=6)
            ent = ctk.CTkEntry(scroll_frame, width=300, height=35)
            ent.grid(row=r, column=1, padx=8, pady=6, sticky="w")
            entries[key] = ent
            r += 1

        if datos:
            for k in entries:
                entries[k].insert(0, str(datos.get(k,"")))

        btn_frame = ctk.CTkFrame(form, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(12,6))
        
        def _guardar():
            payload = {k: entries[k].get().strip() for k in entries}
            if not payload.get("cliente"):
                messagebox.showwarning("Validación", "Cliente requerido")
                return

            self.hist_update_visita(datos["_id"], payload)
            modal.destroy()
            
        ctk.CTkButton(
            btn_frame, 
            text="Guardar", 
            command=_guardar, 
            fg_color=STYLE["exito"],
            height=35,
            width=100
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btn_frame, 
            text="Cancelar", 
            command=modal.destroy,
            height=35,
            width=100
        ).pack(side="left", padx=8)

    # -----------------------------------------------------------
    # NUEVOS MÉTODOS PARA DIAGNÓSTICO Y LIMPIEZA
    # -----------------------------------------------------------
    
    def limpiar_cache_completo(self):
        """Limpia completamente el cache y archivos temporales"""
        try:
            data_dir = os.path.join(os.path.dirname(__file__), "data")
            temp_files = [
                "tabla_de_relacion.json",
                "base_etiquetado.json",
                "__pycache__",
                "*.pyc",
                "*.tmp"
            ]
            
            print("🧹 Iniciando limpieza completa de cache...")
            
            # Eliminar archivos específicos
            for pattern in temp_files:
                if pattern.endswith('.json'):
                    file_path = os.path.join(data_dir, pattern)
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        print(f"🗑️ Eliminado: {file_path}")
                
                elif pattern == "__pycache__":
                    cache_dir = os.path.join(os.path.dirname(__file__), "__pycache__")
                    if os.path.exists(cache_dir):
                        shutil.rmtree(cache_dir)
                        print(f"🗑️ Eliminado directorio: {cache_dir}")
                    
                    # También limpiar cache en subdirectorios
                    for root, dirs, files in os.walk(os.path.dirname(__file__)):
                        if "__pycache__" in dirs:
                            cache_path = os.path.join(root, "__pycache__")
                            shutil.rmtree(cache_path)
                            print(f"🗑️ Eliminado: {cache_path}")
                
                elif pattern in ["*.pyc", "*.tmp"]:
                    for root, dirs, files in os.walk(os.path.dirname(__file__)):
                        for file in files:
                            if file.endswith('.pyc') or file.endswith('.tmp'):
                                file_path = os.path.join(root, file)
                                try:
                                    os.remove(file_path)
                                    print(f"🗑️ Eliminado: {file_path}")
                                except:
                                    pass
            
            # Limpiar variables internas
            self.archivo_excel_cargado = None
            self.archivo_json_generado = None
            self.json_filename = None
            self.archivo_etiquetado_json = None
            
            print("✅ Limpieza de cache completada")
            messagebox.showinfo("Limpieza Completa", "Se ha limpiado completamente el cache del sistema.")
            
        except Exception as e:
            print(f"⚠️ Error durante limpieza: {e}")
            messagebox.showerror("Error", f"No se pudo completar la limpieza:\n{e}")

    def verificar_integridad_datos(self):
        """Verifica la integridad de los datos cargados"""
        try:
            if not self.archivo_json_generado:
                messagebox.showwarning("Sin datos", "No hay archivo cargado para verificar")
                return
            
            with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            
            # Análisis completo
            info = {
                "total_registros": len(datos),
                "columnas": list(datos[0].keys()) if datos else [],
                "familias_unicas": set(),
                "duplicados": 0
            }
            
            # Buscar columna de familia
            columna_familia = None
            for col in ['CODIGO FAMILIA', 'FAMILIA', 'CODIGO_FAMILIA', 'CODIGO']:
                if datos and col in datos[0]:
                    columna_familia = col
                    break
            
            if columna_familia:
                familias = set()
                for item in datos:
                    if columna_familia in item and item[columna_familia]:
                        familias.add(str(item[columna_familia]).strip())
                info['familias_unicas'] = familias
            
            # Verificar duplicados por ID único
            ids_vistos = set()
            for item in datos:
                item_id = str(item.get('ID', '') or str(item.get('FOLIO', '')) or str(item))
                if item_id in ids_vistos:
                    info['duplicados'] += 1
                ids_vistos.add(item_id)
            
            # Mostrar reporte
            reporte = f"""
    📊 REPORTE DE INTEGRIDAD DE DATOS

    📁 Total de registros: {info['total_registros']}
    📋 Columnas encontradas: {', '.join(info['columnas'])}
    👨‍👩‍👧‍👦 Familias únicas: {len(info['familias_unicas'])}
    🔄 Registros duplicados: {info['duplicados']}

    🔍 Columna de familia: {columna_familia or 'NO ENCONTRADA'}
            """
            
            if info['familias_unicas']:
                reporte += f"\n📝 Primeras 5 familias: {sorted(list(info['familias_unicas']))[:5]}"
            
            messagebox.showinfo("Reporte de Integridad", reporte)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo verificar integridad:\n{e}")

# ================== EJECUCIÓN ================== #
if __name__ == "__main__":
    app = SistemaDictamenesVC()
    app.mainloop()