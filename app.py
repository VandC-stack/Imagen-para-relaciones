# -- SISTEMA V&C - GENERADOR DE DICTÁMENES -- #
import os, sys, uuid, shutil
import json
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter as tk
import threading
import subprocess
import importlib
import importlib.util
from datetime import datetime
import unicodedata
import time
import platform
from datetime import datetime

# ---------- ESTILO VISUAL V&C ---------- #
STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008D53",
    "advertencia": "#ff1500",
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
    # --- PAGINACIÓN HISTORIAL ---
    HISTORIAL_PAGINA_ACTUAL = 1
    HISTORIAL_REGS_POR_PAGINA = 1000

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
        self.domicilio_seleccionado = None
        self.archivo_etiquetado_json = None

        # Variables para nueva visita
        self.current_folio = "000001"

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
        # Cargar reservas persistentes
        self.pending_folios = []
        try:
            self._load_pending_folios()
        except Exception:
            self.pending_folios = []
        # Directorio donde están los generadores/documentos (ReportLab, tablas, etc.)
        self.documentos_dir = os.path.join(os.path.dirname(__file__), "Documentos Inspeccion")

        # ===== NUEVA ESTRUCTURA DE NAVEGACIÓN =====
        self.crear_navegacion()
        self.crear_area_contenido()

        # ===== FOOTER =====
        self.crear_footer()

        # Cargar configuración de exportación Excel (persistente)
        self._cargar_config_exportacion()

        # Cargar clientes al iniciar
        self.cargar_clientes_desde_json()
        self.cargar_ultimo_folio()
        try:
            self._generar_datos_exportable()
        except Exception:
            pass

    # ----------------- Overlay de Acciones (botones interactivos) -----------------
    def _create_actions_overlay(self, parent, actions_col=None):
        """Crea un frame flotante con botones que se posiciona sobre la columna 'Acciones'."""
        try:
            overlay = tk.Frame(parent, bg=STYLE.get('fondo', '#fff'))
            overlay.place_forget()
            self._actions_overlay = overlay

            # Botones: Folios, Archivos, Editar, Borrar
            try:
                self._btn_folios = ctk.CTkButton(overlay, text="Folios", width=80, height=26, corner_radius=6, command=lambda: self._overlay_action('folios'))
                self._btn_archivos = ctk.CTkButton(overlay, text="Archivos", width=80, height=26, corner_radius=6, command=lambda: self._overlay_action('archivos'))
                self._btn_editar = ctk.CTkButton(overlay, text="Editar", width=60, height=26, corner_radius=6, command=lambda: self._overlay_action('editar'))
                self._btn_borrar = ctk.CTkButton(overlay, text="Borrar", width=60, height=26, corner_radius=6, fg_color=STYLE['peligro'], command=lambda: self._overlay_action('borrar'))
            except Exception:
                self._btn_folios = tk.Button(overlay, text="Folios", width=8, command=lambda: self._overlay_action('folios'))
                self._btn_archivos = tk.Button(overlay, text="Archivos", width=8, command=lambda: self._overlay_action('archivos'))
                self._btn_editar = tk.Button(overlay, text="Editar", width=6, command=lambda: self._overlay_action('editar'))
                self._btn_borrar = tk.Button(overlay, text="Borrar", width=6, command=lambda: self._overlay_action('borrar'))

            # Empacar botones para ocupar el ancho del overlay equitativamente
            for w in (self._btn_folios, self._btn_archivos, self._btn_editar, self._btn_borrar):
                try:
                    w.pack(side='left', fill='both', expand=True, padx=(2, 2), pady=2)
                except Exception:
                    w.pack(side='left', padx=(2, 2), pady=2)

            # Bindings para mostrar/ocultar y reposicionar el overlay
            self.hist_tree.bind('<Motion>', self._on_tree_motion)
            self.hist_tree.bind('<Leave>', lambda e: self._hide_actions_overlay())
            self.hist_tree.bind('<Button-1>', self._on_tree_click)
            self.hist_tree.bind('<MouseWheel>', lambda e: self._hide_actions_overlay())
        except Exception:
            pass

    def _on_tree_motion(self, event):
        try:
            iid = self.hist_tree.identify_row(event.y)
            if not iid:
                self._hide_actions_overlay()
                return
            bbox = self.hist_tree.bbox(iid, column=self.hist_tree['columns'][-1])
            if not bbox:
                self._hide_actions_overlay()
                return
            x, y, w, h = bbox
            tree_x = self.hist_tree.winfo_x()
            tree_y = self.hist_tree.winfo_y()
            abs_x = tree_x + x
            abs_y = tree_y + y
            try:
                self._actions_overlay.place(x=abs_x, y=abs_y, width=w, height=h)
                self._actions_overlay.lift()
                self._overlay_iid = iid
            except Exception:
                pass
        except Exception:
            pass

    def _hide_actions_overlay(self):
        try:
            if hasattr(self, '_actions_overlay'):
                self._actions_overlay.place_forget()
                self._overlay_iid = None
        except Exception:
            pass

    def _on_tree_click(self, event):
        try:
            col = self.hist_tree.identify_column(event.x)
            iid = self.hist_tree.identify_row(event.y)
            if not iid:
                self._hide_actions_overlay()
                return
            last_col = f"#{len(self.hist_tree['columns'])}"
            if col == last_col:
                bbox = self.hist_tree.bbox(iid, column=self.hist_tree['columns'][-1])
                if bbox:
                    x, y, w, h = bbox
                    tree_x = self.hist_tree.winfo_x()
                    tree_y = self.hist_tree.winfo_y()
                    abs_x = tree_x + x
                    abs_y = tree_y + y
                    self._actions_overlay.place(x=abs_x, y=abs_y, width=w, height=h)
                    self._actions_overlay.lift()
                    self._overlay_iid = iid
            else:
                self._hide_actions_overlay()
        except Exception:
            pass

    def _overlay_action(self, action):
        iid = getattr(self, '_overlay_iid', None)
        if not iid:
            return
        reg = self._hist_map.get(iid)
        if not reg:
            return
        try:
            if action == 'folios':
                self.descargar_folios_visita(reg)
            elif action == 'archivos':
                self.mostrar_opciones_documentos(reg)
            elif action == 'editar':
                self.hist_editar_registro(reg)
            elif action == 'borrar':
                self.hist_eliminar_registro(reg)
        except Exception:
            pass

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

        # Botón Reportes
        self.btn_reportes = ctk.CTkButton(
            botones_frame,
            text="📑 Reportes",
            command=self.mostrar_reportes,
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
        self.btn_reportes.pack(side="left", padx=(0, 10))
        
        # Espacio flexible
        ctk.CTkLabel(botones_frame, text="", fg_color="transparent").pack(side="left", expand=True)
        
        # Información del sistema
        self.lbl_info_sistema = ctk.CTkLabel(
            botones_frame,
            text="Sistema de Dictámenes - V&C",
            font=("Inter", 12),
            text_color=STYLE["texto_claro"]
        )
        # Botón Backup en la barra de navegación (no mostrar por defecto)
        try:
            self.btn_backup = ctk.CTkButton(
                botones_frame,
                text="💾 Backup",
                command=self.hist_hacer_backup,
                height=34, width=110, corner_radius=8,
                fg_color=STYLE["primario"], text_color=STYLE["secundario"], hover_color="#D4BF22"
            )
        except Exception:
            self.btn_backup = None

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

        # Frame para reportes
        self.frame_reportes = ctk.CTkFrame(self.contenido_frame, fg_color="transparent")
        
        # Construir el contenido de cada sección
        self._construir_tab_principal(self.frame_principal)
        self._construir_tab_historial(self.frame_historial)
        self._construir_tab_reportes(self.frame_reportes)
        
        # Mostrar la sección principal por defecto
        self.mostrar_principal()

    def mostrar_principal(self):
        """Muestra la sección principal y oculta las demás"""
        # Ocultar todos los frames primero
        self.frame_principal.pack_forget()
        self.frame_historial.pack_forget()
        # Asegurarse de ocultar la pestaña Reportes también
        try:
            self.frame_reportes.pack_forget()
        except Exception:
            pass
        
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
        try:
            self.btn_reportes.configure(fg_color=STYLE["surface"], text_color=STYLE["secundario"], border_color=STYLE["secundario"])
        except Exception:
            pass
        # Ocultar backup nav cuando no estemos en Historial
        try:
            if getattr(self, 'btn_backup', None):
                self.btn_backup.pack_forget()
        except Exception:
            pass

    def mostrar_historial(self):
            """Muestra la sección de historial y oculta las demás"""
            # Ocultar todos los frames primero
            self.frame_principal.pack_forget()
            self.frame_historial.pack_forget()
            # Asegurarse de ocultar la pestaña Reportes también
            try:
                self.frame_reportes.pack_forget()
            except Exception:
                pass
            
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
            try:
                self.btn_reportes.configure(fg_color=STYLE["surface"], text_color=STYLE["secundario"], border_color=STYLE["secundario"])
            except Exception:
                pass
            
            # Verificar y reparar datos existentes al mostrar historial
            self.verificar_datos_folios_existentes()
            
            # Refrescar el historial si es necesario
            self._cargar_historial()
            self._poblar_historial_ui()

            # Mostrar backup nav cuando estemos en Historial
            try:
                if getattr(self, 'btn_backup', None):
                    self.btn_backup.pack(side="right", padx=(0, 10))
            except Exception:
                pass

    def mostrar_reportes(self):
        """Muestra la sección de reportes y oculta las demás"""
        # Ocultar todos los frames primero
        self.frame_principal.pack_forget()
        self.frame_historial.pack_forget()
        self.frame_reportes.pack(fill="both", expand=True)
        # Asegurarse de ocultar backup nav en la pestaña Reportes
        try:
            if getattr(self, 'btn_backup', None):
                self.btn_backup.pack_forget()
        except Exception:
            pass

        # Actualizar estado de los botones
        try:
            self.btn_principal.configure(fg_color=STYLE["surface"], text_color=STYLE["secundario"], border_color=STYLE["secundario"])
            self.btn_historial.configure(fg_color=STYLE["surface"], text_color=STYLE["secundario"], border_color=STYLE["secundario"])
            self.btn_reportes.configure(fg_color=STYLE["primario"], text_color=STYLE["secundario"], border_color=STYLE["primario"])
        except Exception:
            pass

    def _construir_tab_principal(self, parent):
        """Construye la interfaz principal con dos tarjetas en proporción 30%/70%"""
        # ===== CONTENEDOR PRINCIPAL CON 2 COLUMNAS =====
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True)

        # Configurar grid para 2 columnas con proporción 30%/70%
        main_frame.grid_columnconfigure(0, weight=3)  # 30%
        main_frame.grid_columnconfigure(1, weight=7)  # 70%
        # Mantener ambas tarjetas (izquierda/derecha) del mismo tamaño incluso
        # cuando se muestran/ocultan widgets según el tipo de documento.
        main_frame.grid_rowconfigure(0, weight=1, minsize=480)

        # ===== TARJETA INFORMACIÓN DE VISITA (IZQUIERDA) - 30% =====
        card_visita = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12)
        card_visita.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="nsew")
        try:
            card_visita.grid_propagate(False)
        except Exception:
            pass

        ctk.CTkLabel(
            card_visita,
            text="📋 Información de Visita",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=20, pady=(20, 15))

        visita_frame = ctk.CTkFrame(card_visita, fg_color="transparent")
        visita_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor para el formulario con scrollbar
        scroll_form = ctk.CTkScrollableFrame(
            visita_frame,
            fg_color="transparent",
            scrollbar_button_color="#ecd925",
            scrollbar_button_hover_color="#ecd925"
        )
        scroll_form.pack(fill="both", expand=True)

        # === Tipo de Documento (Dictamen, Negación de dictamen, Constancia, Negación de Constancia) ===
        tipo_doc_frame = ctk.CTkFrame(scroll_form, fg_color="transparent")
        tipo_doc_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            tipo_doc_frame,
            text="Tipo de documento:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", pady=(0, 5))

        self.combo_tipo_documento = ctk.CTkComboBox(
            tipo_doc_frame,
            values=["Dictamen", "Negación de Dictamen", "Constancia", "Negación de Constancia"],
            font=FONT_SMALL,
            dropdown_font=FONT_SMALL,
            state="readonly",
            command=self.actualizar_tipo_documento,
            height=35,
            corner_radius=8
        )
        self.combo_tipo_documento.pack(fill="x", pady=(0, 5))
        self.combo_tipo_documento.set("Dictamen")

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
            placeholder_text="CP0001",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_folio_visita.pack(fill="x", pady=(0, 5))
        folio_con_prefijo = f"CP{self.current_folio}"
        self.entry_folio_visita.insert(0, folio_con_prefijo)
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
        self.entry_hora_inicio.configure(state="readonly")

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
            placeholder_text="18:00",
            font=FONT_SMALL,
            height=35,
            corner_radius=8
        )
        self.entry_hora_termino.pack(fill="x", pady=(0, 5))
        self.entry_hora_termino.insert(0, "18:00")
        self.entry_hora_termino.configure(state="readonly")

        
        # Supervisor field removed from UI: supervisor is derived from the loaded tabla de relación

        # ===== TARJETA GENERADOR (DERECHA) - 70% =====
        card_generacion = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=12)
        card_generacion.grid(row=0, column=1, padx=(10, 0), pady=0, sticky="nsew")
        try:
            card_generacion.grid_propagate(False)
        except Exception:
            pass

        self.generacion_title = ctk.CTkLabel(
            card_generacion,
            text="🚀 Generador de Dictámenes",
            font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        )
        self.generacion_title.pack(anchor="w", padx=20, pady=(20, 15))

        generacion_frame = ctk.CTkFrame(card_generacion, fg_color="transparent")
        generacion_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Contenedor principal de generador con scrollbar
        scroll_generacion = ctk.CTkScrollableFrame(
            generacion_frame,
            fg_color="transparent",
            scrollbar_button_color="#ecd925",
            scrollbar_button_hover_color="#ecd925"
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
        # Encabezado para selector de domicilio
        ctk.CTkLabel(
            cliente_controls_frame,
            text="Domicilio:",
            font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(8,4))

        # --- DOMICILIOS DEL CLIENTE (se rellena al seleccionar cliente) ---
        self.combo_domicilios = ctk.CTkComboBox(
            cliente_controls_frame,
            values=["Seleccione un domicilio..."],
            font=FONT_SMALL,
            dropdown_font=FONT_SMALL,
            state="disabled",
            height=35,
            corner_radius=8,
            command=self._seleccionar_domicilio
        )
        # lo colocamos a la derecha del combo de cliente pero no expandimos
        self.combo_domicilios.pack(side="left", padx=(8, 0))

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

        # --- FOLIOS RESERVADOS (Debajo del selector de cliente) ---
        cliente_folios_frame = ctk.CTkFrame(cliente_section, fg_color="transparent")
        cliente_folios_frame.pack(fill="x", pady=(8, 10))

        self.lbl_folios_pendientes = ctk.CTkLabel(
            cliente_folios_frame,
            text="Folios reservados (seleccione para usar / desmarcar):",
            font=("Inter", 10),
            text_color=STYLE["texto_oscuro"]
        )
        self.lbl_folios_pendientes.pack(side="left", padx=(0,8))

        self.combo_folios_pendientes = ctk.CTkComboBox(
            cliente_folios_frame,
            values=[],
            font=FONT_SMALL,
            dropdown_font=FONT_SMALL,
            state="normal",
            height=30,
            corner_radius=8,
            command=self._seleccionar_folio_pendiente
        )
        self.combo_folios_pendientes.pack(side="left", fill="x", expand=True, padx=(0, 8))

        # Botones compactos: desmarcar y eliminar
        self.btn_desmarcar_folio = ctk.CTkButton(
            cliente_folios_frame,
            text="Desmarcar",
            width=32,
            height=30,
            corner_radius=6,
            fg_color=STYLE["secundario"],
            text_color=STYLE["surface"],
            command=self._desmarcar_folio_seleccionado
        )
        self.btn_desmarcar_folio.pack(side="left", padx=(0, 6))

        self.btn_eliminar_folio_pendiente = ctk.CTkButton(
            cliente_folios_frame,
            text="Eliminar Folio Seleccionado",
            width=32,
            height=30,
            corner_radius=6,
            fg_color=STYLE["peligro"],
            text_color=STYLE["surface"],
            command=self._eliminar_folio_pendiente
        )
        self.btn_eliminar_folio_pendiente.pack(side="left")

        # Botón para configurar carpetas de evidencias (abre modal para elegir grupo y carpeta)
        # Tres botones para modos de pegado (no empacados hasta selección de cliente)
        self.boton_pegado_simple = ctk.CTkButton(
            cliente_section,
            text="🖼️ Pegado Simple",
            command=self.handle_pegado_simple,
            font=("Inter", 11),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=32,
            width=150,
            corner_radius=8
        )

        self.boton_pegado_carpetas = ctk.CTkButton(
            cliente_section,
            text="📁 Pegado Carpetas",
            command=self.handle_pegado_carpetas,
            font=("Inter", 11),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=32,
            width=150,
            corner_radius=8
        )

        self.boton_pegado_indice = ctk.CTkButton(
            cliente_section,
            text="📑 Pegado Índice",
            command=self.handle_pegado_indice,
            font=("Inter", 11),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=32,
            width=150,
            corner_radius=8
        )
        # NOTA: no empacamos el botón aquí para que permanezca oculto hasta
        # que el usuario seleccione un cliente que requiera configuración
        # (se mostrará desde `actualizar_cliente_seleccionado`).

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

        # Dropdown de folios pendientes reubicado debajo del selector de cliente
        # (se crea más abajo en la sección de cliente para mejor flujo UX)

        # Botón de etiquetado DECATHLON (inicialmente oculto)
        self.boton_subir_etiquetado = ctk.CTkButton(
            botones_fila1,
            text="📦 Subir Base de Etiquetado",
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
            text="🧾 Generar Documentos PDF",
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

        # Label para avisar si hay folio pendiente para el tipo seleccionado
        self.info_folio_pendiente = ctk.CTkLabel(
            generar_section,
            text="",
            font=("Inter", 11),
            text_color=STYLE["advertencia"]
        )
        self.info_folio_pendiente.pack(pady=(0, 6))

        # Botón para guardar un folio incompleto / reservado en el historial
        self.boton_guardar_folio = ctk.CTkButton(
            generar_section,
            text="Guardar folio (reservar)",
            command=self.guardar_folio_historial,
            font=("Inter", 12, "bold"),
            fg_color=STYLE["primario"],
            hover_color="#D4BF22",
            text_color=STYLE["secundario"],
            height=34,
            corner_radius=8,
            state="disabled"
        )
        self.boton_guardar_folio.pack(pady=(0, 6))

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
        # Aplicar estado inicial de UI según tipo de documento seleccionado
        try:
            self.actualizar_tipo_documento()
        except Exception:
            pass

    def _construir_tab_historial(self, parent):
        cont = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=8)
        cont.pack(fill="both", expand=True, padx=10, pady=10)

        # ===========================================================
        # BARRA SUPERIOR EN UNA SOLA LÍNEA (COMO EN LA IMAGEN)
        # ===========================================================
        barra_superior = ctk.CTkFrame(cont, fg_color="transparent", height=50)
        barra_superior.pack(fill="x", pady=(0, 10))
        barra_superior.pack_propagate(False)


        # --- FOLIO Y BÚSQUEDA EN MISMA LÍNEA ---
        linea_busqueda = ctk.CTkFrame(barra_superior, fg_color="transparent")
        linea_busqueda.pack(fill="x", pady=5)

        # Folio (izquierda)
        ctk.CTkLabel(
            linea_busqueda, text="Folio visita:", 
            font=("Inter", 11), text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(0, 8))

        self.entry_buscar_folio = ctk.CTkEntry(
            linea_busqueda, width=100, height=25,
            corner_radius=6, placeholder_text="CP0001"
        )
        self.entry_buscar_folio.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            linea_busqueda, text="Buscar",
            command=self.hist_buscar_por_folio,
            width=40, height=25, corner_radius=6,
            fg_color=STYLE["secundario"], text_color=STYLE["surface"]
        ).pack(side="left", padx=(0, 8))

        # Botón Limpiar búsqueda
        ctk.CTkButton(
            linea_busqueda, text="Limpiar",
            command=self.hist_limpiar_busqueda,
            width=60, height=25, corner_radius=6,
            fg_color=STYLE["secundario"], text_color=STYLE["surface"]
        ).pack(side="left", padx=(0, 8))

        # (Se eliminó el botón global 'Borrar' aquí; el borrado ahora está disponible por fila)

       

        # Búsqueda general (derecha)
        ctk.CTkLabel(
            linea_busqueda, text="Búsqueda general:",
            font=("Inter", 11), text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(30, 8))

        self.entry_buscar_general = ctk.CTkEntry(
            linea_busqueda, width=250, height=25,
            corner_radius=6, placeholder_text="Cliente, folio, fecha, supervisor..."
        )
        self.entry_buscar_general.pack(side="left", padx=(0, 8))
        self.entry_buscar_general.bind("<KeyRelease>", self.hist_buscar_general)

        ctk.CTkButton(
            linea_busqueda, text="X",
            command=self.hist_limpiar_busqueda,
            width=40, height=25, corner_radius=6,
            fg_color=STYLE["advertencia"], text_color=STYLE["surface"]
        ).pack(side="left")

        # Espaciador para empujar todo a la izquierda (opcional)
        # ctk.CTkFrame(linea_busqueda, fg_color="transparent").pack(side="left", expand=True)

        # ===========================================================
        # TABLA CON ENCABEZADOS CORREGIDOS (como en la imagen)
        # ===========================================================
        tabla_container = ctk.CTkFrame(cont, fg_color="transparent", corner_radius=8)
        tabla_container.pack(fill="both", expand=True)

        # Encabezados: usamos los headings del Treeview directamente

        # ANCHOS MEJORADOS Y ENCABEZADOS COMO EN LA IMAGEN
        column_widths = [
                40,    # Folio (más pequeño)
                40,    # Acta (más pequeño)
                40,    # Inicio (más pequeño)
                40,    # Término (más pequeño)
                40,    # Hora Ini
                40,    # Hora Fin
                150,   # Cliente (más ancho)
                80,    # Supervisor (ligeramente más pequeño)
                90,   # Tipo de documento
                50,    # Estatus
                60,    # Folios (más compacto)
                400    # Acciones (ajustado: más ancho para mostrar todas las acciones)
            ]

        # Encabezados exactamente como en la imagen
        headers = [
            "Folio", "Acta", "Inicio", "Término", 
            "Hora Inicio", "Hora Fin", "Cliente", 
            "Supervisor", "Tipo de documento", "Estatus", "Folios", "Acciones"
        ]

        # Reemplazar cabecera y scroll por un Treeview virtualizado (más eficiente)
        # Configurar estilo del Treeview para que combine con tema
        style = ttk.Style()
        try:
            # Usar tema 'clam' para permitir colorear encabezados en la mayoría de plataformas
            try:
                style.theme_use('clam')
            except Exception:
                pass

            style.configure("mystyle.Treeview", font=("Inter", 10), rowheight=28,
                            background=STYLE["surface"], fieldbackground=STYLE["surface"], foreground=STYLE["texto_oscuro"]) 
            style.configure("mystyle.Treeview.Heading", font=("Inter", 10, "bold"), background=STYLE["secundario"], foreground=STYLE["surface"], relief='flat')
            # Ajustes del mapa para que el heading mantenga color al interactuar
            style.map('mystyle.Treeview.Heading', background=[('active', STYLE['secundario'])], foreground=[('active', STYLE['surface'])])
        except Exception:
            pass

        # Contenedor para el Treeview
        tree_container = ctk.CTkFrame(tabla_container, fg_color=STYLE["fondo"])
        tree_container.pack(fill="both", expand=True)

        cols = [f"c{i}" for i in range(len(column_widths))]
        self.hist_tree = ttk.Treeview(tree_container, columns=cols, show='headings', style='mystyle.Treeview')
        # Configurar encabezados y anchos (permitir estirar columnas excepto la de Acciones)
        last_idx = len(headers) - 1
        for i, h in enumerate(headers):
            self.hist_tree.heading(cols[i], text=h)
            try:
                stretch = False if i == last_idx else True
                anchor = 'w' if i == last_idx else 'center'
                self.hist_tree.column(cols[i], width=column_widths[i], anchor=anchor, stretch=stretch)
            except Exception:
                self.hist_tree.column(cols[i], width=100, anchor='center')

        # Asegurar que la columna 'Acciones' sea siempre visible y no se reduzca
        try:
            # Aumentar el ancho por defecto y el ancho mínimo de la columna 'Acciones'
            # para que siempre muestre las cuatro opciones sin recortarse.
            self.hist_tree.column(cols[-1], width=360, minwidth=300, stretch=False, anchor='w')
        except Exception:
            try:
                # Fallback seguro
                self.hist_tree.column(cols[-1], width=300, anchor='w')
            except Exception:
                pass

        # Scrollbars: vertical y horizontal — usar grid para posicionar correctamente
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.hist_tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.hist_tree.xview)
        self.hist_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout con grid: tree en (0,0), vsb en (0,1), hsb en (1,0) colspan 2
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        self.hist_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew', columnspan=2)

        # Crear overlay de acciones (botones) y enlazarlo al Treeview
        try:
            self._create_actions_overlay(tree_container, cols[-1])
        except Exception:
            pass

        # Map para acceder al registro original por iid
        self._hist_map = {}

        # Menú contextual para acciones por fila
        self.hist_context_menu = tk.Menu(self, tearoff=0)
        self.hist_context_menu.add_command(label="Folios", command=lambda: self._hist_menu_action('folios'))
        self.hist_context_menu.add_command(label="Archivos", command=lambda: self._hist_menu_action('archivos'))
        self.hist_context_menu.add_command(label="Editar", command=lambda: self._hist_menu_action('editar'))
        self.hist_context_menu.add_command(label="Borrar", command=lambda: self._hist_menu_action('borrar'))

        # Bind derecho y doble-click
        self.hist_tree.bind("<Button-3>", self._hist_show_context_menu)
        self.hist_tree.bind("<Double-1>", self._hist_on_double_click)
        # Click izquierdo en columna de acciones abrirá el menú contextual
        self.hist_tree.bind("<Button-1>", self._hist_on_left_click)

        # ===========================================================
        # PIE DE PÁGINA (COMO EN LA IMAGEN) - Layout mejorado
        # ===========================================================
        footer = ctk.CTkFrame(cont, fg_color="transparent", height=60)
        footer.pack(fill="x", pady=(10, 0))
        footer.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(expand=True, fill="both", padx=12, pady=10)

        # --- Estructura de paginación: columnas izquierda/central/derecha ---
        # Creamos subframes con ancho fijo a izquierda/derecha para asegurar
        # que los botones queden pegados a los bordes, y el centro expande.
        pag_left = ctk.CTkFrame(footer_content, fg_color="transparent")
        pag_center = ctk.CTkFrame(footer_content, fg_color="transparent")
        pag_right = ctk.CTkFrame(footer_content, fg_color="transparent")

        # Fijar anchos laterales para que actúen como 'zonas' pegadas a bordes
        pag_left.configure(width=130)
        pag_right.configure(width=130)
        pag_left.pack_propagate(False)
        pag_right.pack_propagate(False)

        pag_left.pack(side='left', fill='y')
        pag_center.pack(side='left', expand=True, fill='both')
        pag_right.pack(side='right', fill='y')

        # Botón Anterior pegado al borde izquierdo dentro del subframe izquierdo
        self.btn_hist_prev = ctk.CTkButton(
            pag_left, text="⏪ Anterior",
            command=self.hist_pagina_anterior,
            height=28, width=100, corner_radius=6,
            fg_color=STYLE["secundario"], text_color=STYLE["surface"],
            hover_color="#1a1a1a"
        )
        self.btn_hist_prev.pack(side='left', anchor='w', padx=(6,0))

        # Contador centrado en el área central: asegurar centrado absoluto
        # Quitamos widgets laterales que desalineen la etiqueta y la centramos
        self.hist_pagina_label = ctk.CTkLabel(
            pag_center, text="Página 1",
            font=("Inter", 10), text_color=STYLE["texto_oscuro"]
        )
        # pack con expand=True y sin side para centrar horizontalmente
        self.hist_pagina_label.pack(expand=True)

        # Botón Siguiente pegado al borde derecho dentro del subframe derecho
        self.btn_hist_next = ctk.CTkButton(
            pag_right, text="Siguiente ⏩",
            command=self.hist_pagina_siguiente,
            height=28, width=100, corner_radius=6,
            fg_color=STYLE["secundario"], text_color=STYLE["surface"],
            hover_color="#1a1a1a"
        )
        self.btn_hist_next.pack(side='right', anchor='e', padx=(0,6))

        # Note: los botones de EMA/Anual y Backup se muestran en la pestaña "Reportes".

        # Cargar data
        self._cargar_historial()
        self._poblar_historial_ui()
        # Asegurar que el dropdown de folios pendientes se rellene al iniciar
        try:
            if hasattr(self, '_refresh_pending_folios_dropdown'):
                self._refresh_pending_folios_dropdown()
        except Exception:
            pass

    def _construir_tab_reportes(self, parent):
        """Construye la pestaña 'Reportes' con botones EMA y Anual y Backup en la esquina superior derecha."""
        cont = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=8)
        cont.pack(fill="both", expand=True, padx=10, pady=10)

        # Barra superior: título y backup a la derecha
        barra = ctk.CTkFrame(cont, fg_color="transparent", height=50)
        barra.pack(fill="x", pady=(0, 10))
        barra.pack_propagate(False)

        ctk.CTkLabel(barra, text="📑 Reportes", font=FONT_SUBTITLE, text_color=STYLE["texto_oscuro"]).pack(side="left", padx=12)

        # Nota: el botón de Backup se gestiona desde la barra de navegación
        # (evitar duplicarlo aquí para que solo exista una instancia).

        # Contenido central con dos botones grandes: Anual y EMA
        contenido = ctk.CTkFrame(cont, fg_color="transparent")
        contenido.pack(fill="both", expand=True, pady=(10,0))

        btn_frame = ctk.CTkFrame(contenido, fg_color="transparent")
        btn_frame.pack(expand=True)

        ctk.CTkButton(
            btn_frame, text="📈 Generar Anual",
            command=self.descargar_excel_anual,
            height=60, width=220, corner_radius=10,
            fg_color=("#1976D2", "#0D47A1"), text_color=STYLE["secundario"], font=("Inter", 14, "bold")
        ).pack(side="left", padx=20, pady=40)

        ctk.CTkButton(
            btn_frame, text="📊 Generar EMA",
            command=self.descargar_excel_ema,
            height=60, width=220, corner_radius=10,
            fg_color=("#2E7D32", "#1B5E20"), text_color=STYLE["secundario"], font=("Inter", 14, "bold")
        ).pack(side="left", padx=20, pady=40)

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
        """Carga `data/Clientes.json` y rellena el combo de clientes.

        Esta implementación es tolerante a variaciones en la clave del nombre
        (`CLIENTE` o `RAZÓN SOCIAL `) y no modifica el archivo en disco. Los
        datos leídos se guardan en `self.clientes_data` (lista original de dicts)
        y el combobox `self.combo_cliente` se rellena con los nombres detectados.
        """
        posibles_rutas = [
            os.path.join(os.path.dirname(__file__), 'data', 'Clientes.json'),
            os.path.join(os.path.dirname(__file__), 'Clientes.json'),
            'data/Clientes.json',
            'Clientes.json',
            '../data/Clientes.json'
        ]

        archivo_encontrado = None
        for ruta in posibles_rutas:
            try:
                if os.path.exists(ruta):
                    archivo_encontrado = ruta
                    break
            except Exception:
                continue

        if not archivo_encontrado:
            # No hay archivo; dejar combo con valor por defecto
            try:
                self.combo_cliente.configure(values=['Seleccione un cliente...'])
                self.combo_cliente.set('Seleccione un cliente...')
            except Exception:
                pass
            self.clientes_data = []
            return

        try:
            with open(archivo_encontrado, 'r', encoding='utf-8') as f:
                datos = json.load(f)
        except Exception:
            datos = []

        # Guardar lista original en memoria
        self.clientes_data = datos if isinstance(datos, list) else []

        # Construir lista de nombres para mostrar en el combobox
        nombres = []
        for cliente in self.clientes_data:
            # Priorizar claves: 'CLIENTE' > 'RAZÓN SOCIAL ' > 'RAZON SOCIAL' > RFC/CONTRATO
            nombre = None
            if isinstance(cliente, dict):
                nombre = cliente.get('CLIENTE') or cliente.get('RAZÓN SOCIAL ') or cliente.get('RAZON SOCIAL') or cliente.get('RAZON_SOCIAL')
                if not nombre:
                    # Fallbacks
                    nombre = cliente.get('RFC') or cliente.get('NÚMERO_DE_CONTRATO') or cliente.get('NOMBRE')
            if nombre and isinstance(nombre, str) and nombre.strip() != '':
                nombres.append(nombre.strip())

        # Remover duplicados manteniendo orden
        seen = set()
        nombres_unicos = []
        for n in nombres:
            if n not in seen:
                seen.add(n)
                nombres_unicos.append(n)

        # Preparar valores para el combo
        valores = ['Seleccione un cliente...'] + sorted(nombres_unicos, key=lambda s: s.lower())

        try:
            self.combo_cliente.configure(values=valores)
            self.combo_cliente.set('Seleccione un cliente...')
        except Exception:
            pass

    def safe_forget(self, widget):
        """Evita errores al ocultar widgets ya olvidados."""
        try:
            if widget and widget.winfo_ismapped():
                widget.pack_forget()
        except Exception:
            pass

    def safe_pack(self, widget, **kwargs):
        """Evita errores de Tkinter al volver a empacar widgets."""
        try:
            if widget and not widget.winfo_ismapped():
                widget.pack(**kwargs)
        except Exception:
            pass

    def actualizar_cliente_seleccionado(self, cliente_nombre):

        # Reset si selecciona opción vacía
        if cliente_nombre == "Seleccione un cliente...":
            self.cliente_seleccionado = None
            self.info_cliente.configure(
                text="No se ha seleccionado ningún cliente",
                text_color=STYLE["texto_claro"]
            )
            self.boton_limpiar_cliente.configure(state="disabled")
            self.safe_forget(self.boton_subir_etiquetado)
            self.safe_forget(self.info_etiquetado)
            # Ocultar botones de pegado cuando no hay cliente seleccionado
            try:
                self.safe_forget(self.boton_pegado_simple)
                self.safe_forget(self.boton_pegado_carpetas)
                self.safe_forget(self.boton_pegado_indice)
            except Exception:
                pass
            return

        # ─────────────────────────────────────────────
        #  1) CLIENTES QUE SOLO PEGAN EVIDENCIA
        # ─────────────────────────────────────────────
        CLIENTES_EVIDENCIA = {
            # Pegado índice
            "BASECO SAPI DE CV",
            "BLUE STRIPES SA DE CV",
            "GRUPO GUESS S DE RL DE CV",
            "EAST COAST MODA SA DE CV",
            "I NOSTRI FRATELLI S DE RL DE CV",
            "LEDERY MEXICO SA DE CV",
            "MODA RAPSODIA SA DE CV",
            "MULTIBRAND OUTLET STORES SAPI DE CV",
            "RED STRIPES SA DE CV",

            # Pegado simple
            "ROBERT BOSCH S DE RL DE CV",

            # Pegado en múltiples carpetas
            "UNILEVER MANUFACTURERA S DE RL DE CV",
            "UNILEVER DE MÉXICO S DE RL DE CV",
        }

        # ─────────────────────────────────────────────
        # 2) CLIENTES QUE PEGAN ETIQUETAS
        # ─────────────────────────────────────────────
        CLIENTES_ETIQUETA = {
            "ARTICULOS DEPORTIVOS DECATHLON SA DE CV",
            "FERRAGAMO MEXICO S DE RL DE CV",
            "ULTA BEAUTY SAPI DE CV",  # Regla especial
        }

        # Buscar cliente en la lista; aceptar varias claves de nombre
        encontrado = None
        for cliente in self.clientes_data:
            if not isinstance(cliente, dict):
                continue
            nombre_cliente = cliente.get('CLIENTE') or cliente.get('RAZÓN SOCIAL ') or cliente.get('RAZON SOCIAL') or cliente.get('RAZON_SOCIAL') or cliente.get('NOMBRE')
            if nombre_cliente and isinstance(nombre_cliente, str) and nombre_cliente.strip() == cliente_nombre:
                encontrado = cliente
                break

        if encontrado is None:
            # No se encontró por claves comunes; intentar por RFC o contrato si el nombre coincide
            for cliente in self.clientes_data:
                if not isinstance(cliente, dict):
                    continue
                # construir un nombre de fallback mostrado en el combo (tal como se poblaron los valores)
                fallback = cliente.get('CLIENTE') or cliente.get('RAZÓN SOCIAL ') or cliente.get('RAZON SOCIAL') or cliente.get('RAZON_SOCIAL') or cliente.get('RFC') or cliente.get('NÚMERO_DE_CONTRATO')
                if fallback and isinstance(fallback, str) and fallback.strip() == cliente_nombre:
                    encontrado = cliente
                    break

        if not encontrado:
            # No encontrado; mostrar mensaje y salir
            try:
                self.info_cliente.configure(text="Cliente no encontrado", text_color=STYLE["advertencia"])
            except Exception:
                pass
            return

        cliente = encontrado
        self.cliente_seleccionado = cliente
        rfc = cliente.get("RFC", "No disponible")

        display_name = cliente.get('CLIENTE') or cliente.get('RAZÓN SOCIAL ') or cliente.get('RAZON SOCIAL') or cliente.get('RAZON_SOCIAL') or cliente.get('RFC') or cliente.get('NÚMERO_DE_CONTRATO')

        self.info_cliente.configure(
            text=f"✅ {display_name}\n📋 RFC: {rfc}",
            text_color=STYLE["exito"]
        )
        self.boton_limpiar_cliente.configure(state="normal")

        # Rellenar lista de domicilios para este cliente (si existen)
        domicilios = []
        try:
            direcciones = cliente.get('DIRECCIONES')
            if isinstance(direcciones, list) and direcciones:
                for d in direcciones:
                    if not isinstance(d, dict):
                        continue
                    parts = []
                    for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO', 'CALLE_Y_NRO', 'NUMERO'):
                        v = d.get(k) or d.get(k.upper()) if isinstance(d, dict) else None
                        if v:
                            parts.append(str(v))
                    for k in ('COLONIA O POBLACION', 'COLONIA'):
                        v = d.get(k)
                        if v:
                            parts.append(str(v))
                    for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                        v = d.get(k)
                        if v:
                            parts.append(str(v))
                    if d.get('CIUDAD O ESTADO'):
                        parts.append(str(d.get('CIUDAD O ESTADO')))
                    if d.get('CP'):
                        parts.append(str(d.get('CP')))
                    addr = ", ".join(parts).strip()
                    if addr:
                        domicilios.append(addr)

            # si no hay lista de direcciones, intentar con campos a nivel superior
            if not domicilios:
                parts = []
                for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO'):
                    v = cliente.get(k) or cliente.get(k.upper())
                    if v:
                        parts.append(str(v))
                for k in ('COLONIA O POBLACION', 'COLONIA'):
                    v = cliente.get(k)
                    if v:
                        parts.append(str(v))
                for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                    v = cliente.get(k)
                    if v:
                        parts.append(str(v))
                if cliente.get('CIUDAD O ESTADO'):
                    parts.append(str(cliente.get('CIUDAD O ESTADO')))
                if cliente.get('CP') is not None:
                    parts.append(str(cliente.get('CP')))
                addr = ", ".join(parts).strip()
                if addr:
                    domicilios.append(addr)
        except Exception:
            domicilios = []

        if not domicilios:
            domicilios = ["Domicilio no disponible"]

        # Configurar combo de domicilios
        try:
            vals = ['Seleccione un domicilio...'] + domicilios
            self.combo_domicilios.configure(values=vals, state='readonly')
            self.combo_domicilios.set('Seleccione un domicilio...')
            # almacenar lista para referencia y raw dicts alineados
            self._domicilios_list = domicilios
            # construir _domicilios_raw: si DIRECCIONES existían usamos dicts, else build one
            raw = []
            try:
                direcciones = cliente.get('DIRECCIONES')
                if isinstance(direcciones, list) and direcciones:
                    for d in direcciones:
                        if isinstance(d, dict):
                            raw.append(d)
                else:
                    # fallback: construir dict a partir de campos de cliente
                    d = {
                        'CALLE Y NO': cliente.get('CALLE Y NO') or cliente.get('CALLE') or cliente.get('CALLE_Y_NO') or '',
                        'COLONIA O POBLACION': cliente.get('COLONIA O POBLACION') or cliente.get('COLONIA') or '',
                        'MUNICIPIO O ALCADIA': cliente.get('MUNICIPIO O ALCADIA') or cliente.get('MUNICIPIO') or '',
                        'CIUDAD O ESTADO': cliente.get('CIUDAD O ESTADO') or cliente.get('CIUDAD') or '',
                        'CP': cliente.get('CP')
                    }
                    raw.append(d)
            except Exception:
                raw = []

            # Ensure lengths match: if not, pad with minimal dicts
            if len(raw) != len(self._domicilios_list):
                # try to align by creating dicts from the display strings
                aligned = []
                for s in self._domicilios_list:
                    aligned.append({'_display': s})
                raw = aligned

            self._domicilios_raw = raw
            self.domicilio_seleccionado = None
            # limpiar campos individuales
            self.direccion_seleccionada = None
            self.colonia_seleccionada = None
            self.municipio_seleccionado = None
            self.ciudad_seleccionada = None
            self.cp_seleccionado = None
        except Exception:
            pass

        # ─────────────────────────────────────────────
        # CASO 1: SOLO EVIDENCIA
        # ─────────────────────────────────────────────
        if cliente_nombre in CLIENTES_EVIDENCIA:
            self.tipo_operacion = "EVIDENCIA"
            self.safe_forget(self.boton_subir_etiquetado)
            self.safe_forget(self.info_etiquetado)
            # Mostrar botones de pegado (no persisten rutas)
            try:
                self.safe_pack(self.boton_pegado_simple, side="left", padx=(0, 8))
                self.safe_pack(self.boton_pegado_carpetas, side="left", padx=(0, 8))
                self.safe_pack(self.boton_pegado_indice, side="left", padx=(0, 8))
            except Exception:
                pass

        # ─────────────────────────────────────────────
        # CASO 2: PEGADO DE ETIQUETAS
        # ─────────────────────────────────────────────
        if cliente_nombre in CLIENTES_ETIQUETA:
            self.tipo_operacion = "ETIQUETA"

            # ULTA BEAUTY — flujo mixto dependiendo de la NOM
            if cliente_nombre == "ULTA BEAUTY SAPI DE CV":
                self.tipo_operacion = "ULTA"

            # Mostrar botón de carga de etiquetado
            self.safe_pack(self.boton_subir_etiquetado, side="left", padx=(0, 8))

            if self.archivo_etiquetado_json:
                self.safe_pack(self.info_etiquetado, anchor="w", fill="x", pady=(5, 0))

            # Asegurarse de ocultar los botones de evidencias en flujos de etiquetas
            try:
                self.safe_forget(self.boton_pegado_simple)
                self.safe_forget(self.boton_pegado_carpetas)
                self.safe_forget(self.boton_pegado_indice)
            except Exception:
                pass

        # ─────────────────────────────────────────────
        # SI YA SE CARGÓ EL JSON → habilitar dictamen
        # ─────────────────────────────────────────────
        # Mostrar/ocultar el botón de configuración según el cliente seleccionado
        try:
            if cliente_nombre in CLIENTES_EVIDENCIA:
                self.safe_pack(self.boton_pegado_simple, anchor="w", pady=(8, 0))
                self.safe_pack(self.boton_pegado_carpetas, anchor="w", pady=(8, 0))
                self.safe_pack(self.boton_pegado_indice, anchor="w", pady=(8, 0))
            else:
                self.safe_forget(self.boton_pegado_simple)
                self.safe_forget(self.boton_pegado_carpetas)
                self.safe_forget(self.boton_pegado_indice)
        except Exception:
            pass

        if self.archivo_json_generado:
            self.boton_generar_dictamen.configure(state="normal")

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
        try:
            self.boton_pegado_simple.pack_forget()
            self.boton_pegado_carpetas.pack_forget()
            self.boton_pegado_indice.pack_forget()
        except Exception:
            pass
        self.info_etiquetado.pack_forget()
        try:
            self.combo_domicilios.configure(values=["Seleccione un domicilio..."], state='disabled')
            self.combo_domicilios.set('Seleccione un domicilio...')
            self.domicilio_seleccionado = None
        except Exception:
            pass

    def _seleccionar_domicilio(self, domicilio_text):
        """Handler para seleccionar domicilio del cliente."""
        try:
            if domicilio_text == 'Seleccione un domicilio...' or not domicilio_text:
                self.domicilio_seleccionado = None
                # reset component fields
                self.direccion_seleccionada = None
                self.colonia_seleccionada = None
                self.municipio_seleccionado = None
                self.ciudad_seleccionada = None
                self.cp_seleccionado = None
            else:
                # almacenar texto seleccionado y mapear a raw dict si existe
                self.domicilio_seleccionado = domicilio_text
                try:
                    idx = self._domicilios_list.index(domicilio_text)
                except Exception:
                    idx = None
                raw = None
                try:
                    if idx is not None and hasattr(self, '_domicilios_raw') and idx < len(self._domicilios_raw):
                        raw = self._domicilios_raw[idx]
                except Exception:
                    raw = None

                if raw and isinstance(raw, dict):
                    # prefer explicit keys
                    self.direccion_seleccionada = raw.get('CALLE Y NO') or raw.get('CALLE') or raw.get('calle_numero') or raw.get('CALLE_Y_NO') or raw.get('_display')
                    self.colonia_seleccionada = raw.get('COLONIA O POBLACION') or raw.get('COLONIA') or raw.get('colonia')
                    self.municipio_seleccionado = raw.get('MUNICIPIO O ALCADIA') or raw.get('MUNICIPIO') or raw.get('municipio')
                    self.ciudad_seleccionada = raw.get('CIUDAD O ESTADO') or raw.get('CIUDAD') or raw.get('ciudad_estado')
                    self.cp_seleccionado = raw.get('CP')
                else:
                    # fallback: store full text in direccion_seleccionada
                    self.direccion_seleccionada = domicilio_text
                    self.colonia_seleccionada = None
                    self.municipio_seleccionado = None
                    self.ciudad_seleccionada = None
                    self.cp_seleccionado = None

            # Actualizar la vista de info_cliente para mostrar domicilio elegido
            try:
                if self.cliente_seleccionado:
                    display_name = self.cliente_seleccionado.get('CLIENTE') or self.cliente_seleccionado.get('RAZÓN SOCIAL ') or self.cliente_seleccionado.get('RAZON SOCIAL') or self.cliente_seleccionado.get('RFC') or ''
                    rfc = self.cliente_seleccionado.get('RFC', 'No disponible')
                    if self.domicilio_seleccionado:
                        self.info_cliente.configure(text=f"✅ {display_name}\n📋 RFC: {rfc}\n🏠 {self.direccion_seleccionada}", text_color=STYLE['exito'])
                    else:
                        self.info_cliente.configure(text=f"✅ {display_name}\n📋 RFC: {rfc}", text_color=STYLE['exito'])
            except Exception:
                pass
        except Exception:
            pass

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
                        folio_raw = visita.get("folio_visita", "")
                        # Extraer solo los dígitos del folio (soporta prefijo CP)
                        folio_digits = ''.join([c for c in str(folio_raw) if c.isdigit()])
                        if folio_digits:
                            try:
                                folios_existentes.add(int(folio_digits))
                            except Exception:
                                pass
                    
                    # Encontrar el primer folio disponible
                    folio_disponible = 1
                    while folio_disponible in folios_existentes:
                        folio_disponible += 1
                    
                    self.current_folio = f"{folio_disponible:06d}"
                else:
                    self.current_folio = "000001"

                # Actualizar el campo en la interfaz con prefijo CP (si existen widgets)
                try:
                    if hasattr(self, 'entry_folio_visita') and hasattr(self, 'entry_folio_acta'):
                        self.entry_folio_visita.configure(state="normal")
                        self.entry_folio_visita.delete(0, "end")
                        folio_con_prefijo = f"CP{self.current_folio}"
                        self.entry_folio_visita.insert(0, folio_con_prefijo)
                        self.entry_folio_visita.configure(state="readonly")

                        # Actualizar también el folio del acta
                        self.entry_folio_acta.configure(state="normal")
                        self.entry_folio_acta.delete(0, "end")
                        self.entry_folio_acta.insert(0, f"AC{self.current_folio}")
                        self.entry_folio_acta.configure(state="readonly")
                except Exception:
                    pass
                    
        except Exception as e:
            print(f"❌ Error cargando último folio: {e}")
            self.current_folio = "000001"

    def crear_nueva_visita(self):
        """Prepara el formulario para una nueva visita"""
        try:
            # Obtener el siguiente folio disponible
            self.cargar_ultimo_folio()
            
            # Actualizar campos con prefijo CP
            self.entry_folio_visita.configure(state="normal")
            self.entry_folio_visita.delete(0, "end")
            folio_con_prefijo = f"CP{self.current_folio}"
            self.entry_folio_visita.insert(0, folio_con_prefijo)
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
            # Supervisor input removed from UI; nothing to clear here
            
            # No forzamos el tipo de documento: respetar la selección actual
            try:
                seleccionado = self.combo_tipo_documento.get().strip() if hasattr(self, 'combo_tipo_documento') else None
            except Exception:
                seleccionado = None

            # Si existe un folio pendiente para este tipo, ofrecer reutilizarlo
            try:
                if seleccionado:
                    pendientes = [r for r in getattr(self, 'historial_data', []) if (r.get('tipo_documento') or '').strip() == seleccionado and (r.get('estatus','').lower() == 'pendiente')]
                    if pendientes:
                        primero = pendientes[0]
                        usar = messagebox.askyesno("Folio pendiente encontrado", f"Se encontró un folio pendiente: {primero.get('folio_visita','-')} / {primero.get('folio_acta','-')}\n¿Desea usarlo para esta visita?")
                        if usar:
                            # Cargar folios pendientes en el formulario
                            try:
                                self.entry_folio_visita.configure(state='normal')
                                self.entry_folio_visita.delete(0, 'end')
                                self.entry_folio_visita.insert(0, primero.get('folio_visita',''))
                                self.entry_folio_visita.configure(state='readonly')

                                self.entry_folio_acta.configure(state='normal')
                                self.entry_folio_acta.delete(0, 'end')
                                self.entry_folio_acta.insert(0, primero.get('folio_acta',''))
                                self.entry_folio_acta.configure(state='readonly')
                            except Exception:
                                pass

                            # Marcar el registro como en proceso para no sugerirlo de nuevo
                            try:
                                rid = primero.get('_id') or primero.get('id')
                                if rid:
                                    self.hist_update_visita(rid, {'estatus': 'En proceso'})
                            except Exception:
                                pass
                            return

            except Exception:
                pass

            messagebox.showinfo("Nueva Visita", "Formulario listo para nueva visita")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear nueva visita:\n{e}")

    def guardar_visita_desde_formulario(self):
        """Guarda una nueva visita desde el formulario principal"""
        try:
            if not self.cliente_seleccionado:
                messagebox.showwarning("Cliente requerido", "Por favor seleccione un cliente primero.")
                return

            if not getattr(self, 'domicilio_seleccionado', None):
                messagebox.showwarning("Domicilio requerido", "Por favor seleccione un domicilio para el cliente antes de guardar la visita.")
                return

            # Recoger datos del formulario
            folio_visita = self.entry_folio_visita.get().strip()
            folio_acta = self.entry_folio_acta.get().strip()
            fecha_inicio = self.entry_fecha_inicio.get().strip()
            fecha_termino = self.entry_fecha_termino.get().strip()
            hora_inicio = self.entry_hora_inicio.get().strip()
            hora_termino = self.entry_hora_termino.get().strip()
            # Leer supervisor de forma segura (puede no existir en algunos flujos)
            safe_supervisor_widget = getattr(self, 'entry_supervisor', None)
            try:
                supervisor = safe_supervisor_widget.get().strip() if safe_supervisor_widget and safe_supervisor_widget.winfo_exists() else ""
            except Exception:
                supervisor = ""

            # Leer tipo de documento (conservar la selección tal cual: título / mayúsculas según opciones)
            tipo_documento = (self.combo_tipo_documento.get().strip()
                               if hasattr(self, 'combo_tipo_documento') else "Dictamen")

            # Permitir guardar aunque no haya folio_acta si hay tipo_documento
            if not folio_acta:
                if tipo_documento:
                    # Guardar registro incompleto solo con tipo de documento
                    payload = {
                        "folio_visita": folio_visita,
                        "folio_acta": folio_acta,
                        "fecha_inicio": fecha_inicio,
                        "fecha_termino": fecha_termino,
                        "hora_inicio": hora_inicio,
                        "hora_termino": hora_termino,
                        "norma": "",
                        "cliente": self.cliente_seleccionado['CLIENTE'],
                        "nfirma1": supervisor,
                        "nfirma2": "",
                        "estatus": "En proceso",
                            "tipo_documento": tipo_documento,
                        "folios_utilizados": f"{folio_visita} - {folio_visita}"  # Guardar el folio como rango único
                    }
                    self.hist_create_visita(payload)
                    self.crear_nueva_visita()
                    messagebox.showinfo("Registro guardado", "El folio se guardó como registro incompleto. Podrá completarlo más adelante.")
                    return
                else:
                    messagebox.showwarning("Datos incompletos", "Por favor ingrese el folio de acta o seleccione un tipo de documento.")
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
                "estatus": "En proceso",
                "tipo_documento": tipo_documento
            }

            # Añadir datos de dirección seleccionada si existen
            try:
                payload['direccion'] = getattr(self, 'direccion_seleccionada', '') or getattr(self, 'domicilio_seleccionado', '')
                # también guardar alias `calle_numero` para compatibilidad con generadores
                payload['calle_numero'] = payload.get('direccion') or getattr(self, 'direccion_seleccionada', '')
                payload['colonia'] = getattr(self, 'colonia_seleccionada', '')
                payload['municipio'] = getattr(self, 'municipio_seleccionado', '')
                payload['ciudad_estado'] = getattr(self, 'ciudad_seleccionada', '')
                payload['cp'] = getattr(self, 'cp_seleccionado', '')
            except Exception:
                pass

            # Guardar visita
            self.hist_create_visita(payload)
            # Limpiar formulario después de guardar
            self.crear_nueva_visita()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la visita:\n{e}")
            return

        # Si algo falla al guardar, ya se manejó arriba

    def cargar_excel(self):
        """Carga un archivo Excel y actualiza el UI con el nombre del archivo cargado."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            return

        self.archivo_excel_cargado = file_path
        nombre_archivo = os.path.basename(file_path)
        
        try:
            self.info_archivo.configure(
                text=f"📄 {nombre_archivo}",
                text_color=STYLE["exito"]
            )
        except Exception:
            pass
        try:
            self.boton_cargar_excel.configure(state="disabled")
            self.boton_limpiar.configure(state="normal")
        except Exception:
            pass
        
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
                # Regenerar cache exportable para Excel (persistente)
                try:
                    self._generar_datos_exportable()
                except Exception:
                    pass
                self.guardar_folios_visita(self.current_folio, records)

                # Nota: la generación del Acta NO se realiza automáticamente aquí.
                # El acta se generará únicamente cuando el usuario pulse "Descargar" desde
                # el menú de Archivos de la visita correspondiente. Esto evita confusión
                # y generación de archivos no solicitados.

            self.after(0, self._actualizar_ui_conversion_exitosa, output_path, len(records))

        except Exception as e:
            self.after(0, self.mostrar_error, f"Error al convertir el archivo:\n{e}")
    
    def _extraer_informacion_folios(self, datos_tabla):
        """Extrae y procesa la información de folios de la tabla de relación"""
        try:
            # Verificar si hay datos en la tabla
            if not datos_tabla:
                return {
                    "hay_folios": False,
                    "total_folios": 0,
                    "total_folios_numericos": 0,
                    "mensaje": "No hay datos en la tabla"
                }
            
            folios_encontrados = []
            folios_numericos = []
            hay_folios_asignados = False
            
            # Buscar la columna FOLIO en los datos
            for item in datos_tabla:
                if 'FOLIO' in item:
                    folio_valor = item['FOLIO']
                    
                    # Verificar si el folio tiene un valor asignado (no NaN, None o vacío)
                    if (folio_valor is not None and 
                        str(folio_valor).strip() != "" and 
                        str(folio_valor).lower() != "nan" and
                        str(folio_valor).lower() != "none"):
                        
                        hay_folios_asignados = True
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
                "hay_folios": hay_folios_asignados,
                "total_folios": len(folios_encontrados),
                "total_registros": len(datos_tabla),
                "total_folios_numericos": len(folios_numericos),
                "rango_folios": "",
                "lista_folios": folios_encontrados,
                "folios_formateados": folios_encontrados,
                "mensaje": ""
            }
            
            # Calcular rango si hay folios numéricos
            if folios_numericos:
                min_folio = min(folios_numericos)
                max_folio = max(folios_numericos)
                info_folios["rango_folios"] = f"{min_folio:06d} - {max_folio:06d}"
                info_folios["rango_numerico"] = f"{min_folio} - {max_folio}"
                
                # Determinar mensaje
                if len(folios_numericos) == 1:
                    info_folios["mensaje"] = f"Folio: {min_folio:06d}"
                else:
                    es_consecutivo = all(
                        folios_numericos[i] + 1 == folios_numericos[i + 1] 
                        for i in range(len(folios_numericos) - 1)
                    )
                    if es_consecutivo:
                        info_folios["mensaje"] = f"Total: {len(folios_numericos)} | Rango: {min_folio:06d} - {max_folio:06d}"
                    else:
                        info_folios["mensaje"] = f"Total: {len(folios_numericos)} | Folios asignados"
            elif hay_folios_asignados:
                # Si hay folios pero no son numéricos
                info_folios["mensaje"] = f"Total: {len(folios_encontrados)} | Folios no numéricos"
            else:
                # Si no hay folios asignados
                info_folios["mensaje"] = f"Total: {len(datos_tabla)} | Sin folios asignados"
            
            # Guardar información de folios para usar después
            self.info_folios_actual = info_folios
            
            print(f"📊 Información de folios extraída:")
            print(f"   - ¿Hay folios asignados?: {'Sí' if hay_folios_asignados else 'No'}")
            print(f"   - Total registros: {info_folios['total_registros']}")
            print(f"   - Folios asignados: {info_folios['total_folios']}")
            print(f"   - Folios numéricos: {info_folios['total_folios_numericos']}")
            print(f"   - Mensaje: {info_folios['mensaje']}")
            if folios_numericos and len(folios_numericos) > 1:
                print(f"   - Rango: {info_folios['rango_folios']}")
            
            return info_folios
            
        except Exception as e:
            print(f"⚠️ Error extrayendo información de folios: {e}")
            return {
                "hay_folios": False,
                "total_folios": 0,
                "total_folios_numericos": 0,
                "mensaje": f"Error: {str(e)}"
            }

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

        if not getattr(self, 'domicilio_seleccionado', None):
            messagebox.showwarning("Domicilio no seleccionado", "Por favor seleccione un domicilio para el cliente antes de generar los dictámenes.")
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
            # Si hay un folio reservado seleccionado, advertir al usuario
            if getattr(self, 'usando_folio_reservado', False) and getattr(self, 'selected_pending_id', None):
                sel_msg = f"Hay un folio reservado seleccionado (ID: {self.selected_pending_id}). Si confirma, se usará ese folio para la visita.\n\n"
                sel_msg += "¿Desea continuar y usar el folio reservado?"
                if not messagebox.askyesno("Folio reservado seleccionado", sel_msg):
                    return

            # Antes de confirmar, validar que los inspectores asignados estén acreditados
            try:
                # Determinar visit actual en el historial por folio
                visit_folio_key = f"CP{self.current_folio}" if hasattr(self, 'current_folio') and self.current_folio else None
                visit_actual = None
                try:
                    visitas_hist = self.historial.get('visitas', []) if isinstance(self.historial, dict) else []
                    for v in (visitas_hist or []):
                        try:
                            if visit_folio_key and (v.get('folio_visita') or v.get('folio') ) == visit_folio_key:
                                visit_actual = v
                                break
                        except Exception:
                            continue
                except Exception:
                    visit_actual = None

                # Extraer normas desde los datos (buscar claves comunes)
                normas_en_datos = set()
                for item in (datos or []):
                    try:
                        # buscar claves que contengan 'norm' o exactamente 'nom'
                        for k, val in (item or {}).items():
                            if not k or val is None:
                                continue
                            kn = str(k).lower()
                            if 'norm' in kn or kn == 'nom':
                                # soportar listas o strings
                                if isinstance(val, (list, tuple)):
                                    for s in val:
                                        try:
                                            normas_en_datos.add(str(s).strip())
                                        except Exception:
                                            continue
                                else:
                                    for s in str(val).split(','):
                                        s2 = s.strip()
                                        if s2:
                                            normas_en_datos.add(s2)
                    except Exception:
                        continue

                problemas = []
                sugerencias = {}
                if visit_actual and normas_en_datos:
                    # cargar Firmas.json para mapa nombre->normas
                    try:
                        firmas_path = os.path.join(os.path.dirname(__file__), 'data', 'Firmas.json')
                        with open(firmas_path, 'r', encoding='utf-8') as ff:
                            firmas_data = json.load(ff)
                    except Exception:
                        firmas_data = []

                    firma_map = {}
                    for f in (firmas_data or []):
                        try:
                            name = f.get('NOMBRE DE INSPECTOR') or f.get('NOMBRE') or ''
                            normas_ac = f.get('Normas acreditadas') or f.get('Normas') or []
                            firma_map[name] = [str(n).strip() for n in (normas_ac or [])]
                        except Exception:
                            continue

                    # inspectores asignados en la visita
                    assigned_raw = (visit_actual.get('supervisores_tabla') or visit_actual.get('nfirma1') or '')
                    assigned = [s.strip() for s in str(assigned_raw).split(',') if s.strip()]

                    for norma in normas_en_datos:
                        # verificar si al menos un assigned tiene la norma
                        ok = False
                        missing_inspectores = []
                        for a in assigned:
                            try:
                                acc = firma_map.get(a, [])
                                if any(str(n).strip() == str(norma).strip() for n in (acc or [])):
                                    ok = True
                                    break
                                else:
                                    missing_inspectores.append(a)
                            except Exception:
                                missing_inspectores.append(a)
                        if not ok:
                            problemas.append((norma, missing_inspectores))
                            # sugerir inspectores que sí tienen la norma
                            sugeridas = []
                            for name, acc in firma_map.items():
                                try:
                                    if any(str(n).strip() == str(norma).strip() for n in (acc or [])):
                                        sugeridas.append(name)
                                except Exception:
                                    continue
                            sugerencias[norma] = sugeridas

                if problemas:
                    # Construir mensaje resumido
                    msg_lines = ["Se detectaron posibles conflictos de acreditación para las siguientes normas:"]
                    for norma, miss in problemas:
                        msg_lines.append(f"- {norma}: inspectores asignados no acreditados -> {', '.join(miss) or 'N/A'}")
                        sug = sugerencias.get(norma) or []
                        if sug:
                            msg_lines.append(f"  Sugeridos: {', '.join(sug[:5])}{'...' if len(sug)>5 else ''}")
                        else:
                            msg_lines.append(f"  Sugeridos: (ninguno encontrado)")

                    msg_lines.append("")
                    # Preparar lista única de sugerencias
                    sugeridos_unicos = []
                    try:
                        for norma, lst in sugerencias.items():
                            for s in (lst or []):
                                if s and s not in sugeridos_unicos:
                                    sugeridos_unicos.append(s)
                    except Exception:
                        sugeridos_unicos = []

                    # Si no hay sugerencias, mantener comportamiento anterior
                    if not sugeridos_unicos:
                        msg_lines.append("¿Desea abrir el editor de la visita para corregir los inspectores antes de generar?")
                        if messagebox.askyesno("Inspectores no acreditados", "\n".join(msg_lines)):
                            try:
                                if visit_actual:
                                    self.hist_editar_registro(visit_actual)
                                    return
                            except Exception:
                                pass
                    else:
                        # Construir modal para seleccionar inspectores sugeridos
                        dlg = tk.Toplevel(self)
                        dlg.title("Seleccionar inspectores sugeridos")
                        dlg.geometry("600x420")
                        dlg.transient(self)
                        dlg.grab_set()

                        tk.Label(dlg, text="Se detectaron conflictos de acreditación:", font=(None, 10, 'bold')).pack(anchor='w', padx=12, pady=(8,4))
                        text = tk.Text(dlg, height=6, wrap='word')
                        text.insert('1.0', "\n".join(msg_lines))
                        text.configure(state='disabled')
                        text.pack(fill='both', padx=12, pady=(0,8), expand=False)

                        tk.Label(dlg, text="Seleccione uno o más inspectores sugeridos para aplicar a la visita:", anchor='w').pack(anchor='w', padx=12)
                        frame_checks = tk.Frame(dlg)
                        frame_checks.pack(fill='both', padx=12, pady=(6,6), expand=True)

                        check_vars = []
                        for name in sugeridos_unicos:
                            var = tk.BooleanVar(value=False)
                            cb = tk.Checkbutton(frame_checks, text=name, variable=var, anchor='w')
                            cb.pack(anchor='w')
                            check_vars.append((name, var))

                        # Radio: modo aplicar
                        modo_var = tk.StringVar(value='append')
                        modo_frame = tk.Frame(dlg)
                        modo_frame.pack(fill='x', padx=12, pady=(6,4))
                        tk.Label(modo_frame, text='Modo:').pack(side='left')
                        tk.Radiobutton(modo_frame, text='Añadir a inspectores asignados', variable=modo_var, value='append').pack(side='left', padx=6)
                        tk.Radiobutton(modo_frame, text='Reemplazar inspectores asignados', variable=modo_var, value='replace').pack(side='left', padx=6)

                        btn_frame = tk.Frame(dlg)
                        btn_frame.pack(fill='x', padx=12, pady=10)

                        result = {'action': None, 'selected': []}

                        def _apply_and_continue():
                            sel = [n for n, v in check_vars if v.get()]
                            result['action'] = 'apply'
                            result['selected'] = sel
                            dlg.destroy()

                        def _skip_and_continue():
                            result['action'] = 'skip'
                            dlg.destroy()

                        def _cancel():
                            result['action'] = 'cancel'
                            dlg.destroy()

                        tk.Button(btn_frame, text='Aplicar y continuar', command=_apply_and_continue).pack(side='left')
                        tk.Button(btn_frame, text='Omitir y continuar', command=_skip_and_continue).pack(side='left', padx=8)
                        tk.Button(btn_frame, text='Cancelar', command=_cancel).pack(side='right')

                        # Esperar cierre modal
                        self.wait_window(dlg)

                        if result.get('action') == 'cancel':
                            return

                        if result.get('action') == 'apply' and result.get('selected'):
                            chosen = result.get('selected')
                            # Calcular nuevos supervisores según el modo
                            try:
                                existing_assigned = [s.strip() for s in str(assigned_raw).split(',') if s.strip()]
                            except Exception:
                                existing_assigned = []
                            if modo_var.get() == 'replace':
                                new_list = chosen
                            else:
                                new_list = existing_assigned.copy()
                                for c in chosen:
                                    if c and c not in new_list:
                                        new_list.append(c)

                            joined_new = ', '.join(new_list)
                            # Actualizar visita en el historial (si existe)
                            try:
                                if visit_actual:
                                    id_for_update = visit_actual.get('_id') or visit_actual.get('id') or visit_actual.get('folio_visita') or visit_actual.get('folio_acta')
                                    nuevos = {'supervisores_tabla': joined_new, 'nfirma1': joined_new}
                                    try:
                                        self.hist_update_visita(id_for_update, nuevos)
                                    except Exception:
                                        # fallback: modificar en memoria y guardar
                                        visit_actual['supervisores_tabla'] = joined_new
                                        visit_actual['nfirma1'] = joined_new
                                        try:
                                            self._guardar_historial()
                                        except Exception:
                                            pass
                            except Exception:
                                pass
                            # continuar a confirmación
                        # si skip o no se eligieron, continuar normalmente

            except Exception:
                # en caso de errores en la validación, continuar con confirmación normal
                pass

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
                    
                    
                    if dictamenes_fallidos > 0:
                        mensaje_final += f"\n❌ Dictámenes no generados: {dictamenes_fallidos}"
                        if folios_fallidos:
                            mensaje_final += f"\n📋 Folios fallidos: {', '.join(map(str, folios_fallidos))}"
                    
                    # VERIFICAR ANTES DE MOSTRAR MESSAGEBOX
                    if self.winfo_exists():
                        self.after(0, lambda: messagebox.showinfo("Generación Completada", mensaje_final) if self.winfo_exists() else None)
                        
                        resultado['folios_utilizados_info'] = folios_utilizados
                        self.registrar_visita_automatica(resultado)

                        # Si se generó usando un folio reservado seleccionado, marcarlo como completado
                        try:
                            sel_id = getattr(self, 'selected_pending_id', None)
                            if sel_id:
                                try:
                                    self.hist_update_visita(sel_id, {'estatus': 'Completada'})
                                except Exception:
                                    # fallback: buscar y modificar manualmente
                                    for v in self.historial.get('visitas', []):
                                        if v.get('_id') == sel_id or v.get('id') == sel_id:
                                            v['estatus'] = 'Completada'
                                    try:
                                        self._guardar_historial()
                                    except Exception:
                                        pass

                                # Eliminar de archivo de reservas
                                try:
                                    pf = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
                                    if os.path.exists(pf):
                                        with open(pf, 'r', encoding='utf-8') as f:
                                            arr = json.load(f) or []
                                        # Eliminar por _id / id si coincide con sel_id
                                        try:
                                            arr = [p for p in arr if ((p.get('_id') or p.get('id')) != sel_id)]
                                        except Exception:
                                            arr = [p for p in arr if p.get('folio_visita') != (getattr(self, 'entry_folio_visita', None).get() if hasattr(self, 'entry_folio_visita') else None)]
                                        with open(pf, 'w', encoding='utf-8') as f:
                                            json.dump(arr, f, ensure_ascii=False, indent=2)
                                        self.pending_folios = arr
                                except Exception:
                                    pass
                                # limpiar selección
                                try:
                                    self.selected_pending_id = None
                                    self.usando_folio_reservado = False
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        
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
                # Asegurar porcentaje entre 0 y 100
                try:
                    pct = float(porcentaje)
                except Exception:
                    pct = 0.0
                pct = max(0.0, min(100.0, pct))
                self.barra_progreso.set(pct / 100.0)
                # Mostrar porcentaje y mensaje breve
                if mensaje:
                    self.etiqueta_progreso.configure(text=f"{int(pct)}% - {mensaje}")
                else:
                    self.etiqueta_progreso.configure(text=f"{int(pct)}%")
                self.update_idletasks()
        
        self.after(0, _actualizar)

    def actualizar_tipo_documento(self, valor=None):
        """Actualiza la UI del panel Generador según el tipo de documento seleccionado."""
        try:
            # Obtener selección (si valor pasado por callback, usarlo)
            seleccionado = valor if valor else (self.combo_tipo_documento.get() if hasattr(self, 'combo_tipo_documento') else 'Dictamen')
            seleccionado = seleccionado.strip()

            # Actualizar título del panel de generación
            title_map = {
                'Dictamen': 'Generar Dictámenes',
                'Negación de Dictamen': 'Generar Negación de Dictamen',
                'Constancia': 'Generar Constancias',
                'Negación de Constancia': 'Generar Negación de Constancia'
            }
            nuevo_titulo = title_map.get(seleccionado, 'Generador')
            if hasattr(self, 'generacion_title'):
                self.generacion_title.configure(text=f"🚀 {nuevo_titulo}")

            # Cambiar texto del botón principal de generación
            if hasattr(self, 'boton_generar_dictamen'):
                self.boton_generar_dictamen.configure(text=f"{nuevo_titulo}")

            # Mostrar/ocultar controles de carga según requerimiento (siempre mostrar cliente + subir/limpiar/verificar)
            # Habilitar el botón de guardar folio para permitir reservar folio incompleto
            if hasattr(self, 'boton_guardar_folio'):
                self.boton_guardar_folio.configure(state='normal')

            # Buscar en historial si hay folio pendiente para este tipo
            pendiente_msg = ''
            try:
                pendientes = [r for r in getattr(self, 'historial_data', []) if (r.get('tipo_documento') or '').strip() == seleccionado and r.get('estatus','').lower() in ('en proceso','pendiente')]
                if pendientes:
                    first = pendientes[0]
                    pendiente_msg = f"Folio pendiente: {first.get('folio_visita','-')} / {first.get('folio_acta','-')} (puede usarlo al iniciar una {seleccionado})"
            except Exception:
                pendiente_msg = ''

            if hasattr(self, 'info_folio_pendiente'):
                self.info_folio_pendiente.configure(text=pendiente_msg)

            # Ajustes visuales adicionales: si no hay cliente seleccionado, deshabilitar generación
            if not getattr(self, 'cliente_seleccionado', None):
                if hasattr(self, 'boton_generar_dictamen'):
                    self.boton_generar_dictamen.configure(state='disabled')
            else:
                if hasattr(self, 'boton_generar_dictamen'):
                    self.boton_generar_dictamen.configure(state='normal')

        except Exception as e:
            # No bloquear la aplicación por errores en esta actualización
            print(f"Error actualizando tipo de documento UI: {e}")

        # Refrescar lista de folios pendientes en el combobox (si existe)
        try:
            if hasattr(self, '_refresh_pending_folios_dropdown'):
                self._refresh_pending_folios_dropdown()
            elif hasattr(self, 'combo_folios_pendientes'):
                self._refresh_pending_folios_dropdown()
        except Exception:
            pass

    def guardar_folio_historial(self):
        """Guarda el folio actual en el historial como registro incompleto para retomarlo después."""
        try:
            tipo_documento = (self.combo_tipo_documento.get().strip() if hasattr(self, 'combo_tipo_documento') else 'Dictamen')
            folio_visita = (self.entry_folio_visita.get().strip() if hasattr(self, 'entry_folio_visita') else '')
            folio_acta = (self.entry_folio_acta.get().strip() if hasattr(self, 'entry_folio_acta') else '')

            if not folio_visita:
                messagebox.showwarning("Folio requerido", "No hay folio de visita disponible para guardar.")
                return

            # Requerir cliente y domicilio para guardar/registrar visita
            if not getattr(self, 'cliente_seleccionado', None):
                messagebox.showwarning("Cliente requerido", "Por favor seleccione un cliente antes de guardar la visita.")
                return
            if not getattr(self, 'domicilio_seleccionado', None):
                messagebox.showwarning("Domicilio requerido", "Por favor seleccione un domicilio para el cliente antes de guardar la visita.")
                return

            # Garantizar que si no hay fecha/hora/cliente en el formulario, se registre la marca temporal y el cliente actual
            fecha_inicio_val = (self.entry_fecha_inicio.get().strip() if hasattr(self, 'entry_fecha_inicio') else '')
            hora_inicio_val = (self.entry_hora_inicio.get().strip() if hasattr(self, 'entry_hora_inicio') else '')
            if not fecha_inicio_val:
                fecha_inicio_val = datetime.now().strftime("%d/%m/%Y")
            if not hora_inicio_val:
                hora_inicio_val = datetime.now().strftime("%H:%M")

            cliente_val = ""
            try:
                # soportar dict con clave 'CLIENTE' o 'cliente'
                if getattr(self, 'cliente_seleccionado', None):
                    cliente_val = self.cliente_seleccionado.get('CLIENTE') or self.cliente_seleccionado.get('cliente') or str(self.cliente_seleccionado)
            except Exception:
                cliente_val = ""
            # Determinar los folios reales usados para generación (tomados de la tabla cargada)
            folios_utilizados_val = ""
            try:
                info = getattr(self, 'info_folios_actual', None)
                if info:
                    if info.get('rango_folios'):
                        folios_utilizados_val = info.get('rango_folios')
                    elif info.get('lista_folios'):
                        # unir una lista corta para mostrar
                        lf = info.get('lista_folios')
                        if isinstance(lf, (list, tuple)) and lf:
                            folios_utilizados_val = ','.join(lf[:20])
                        else:
                            folios_utilizados_val = str(lf)
            except Exception:
                folios_utilizados_val = ""

            # Normalizar tipo de documento a valores esperados
            def _normalizar_td(raw):
                if not raw:
                    return 'Dictamen'
                s = str(raw).strip()
                low = s.lower()
                if 'dictamen' in low and ('neg' in low or 'negación' in low or 'negacion' in low):
                    return 'Negación de Dictamen'
                if 'dictamen' in low:
                    return 'Dictamen'
                if 'constancia' in low and ('neg' in low or 'negación' in low or 'negacion' in low):
                    return 'Negación de Constancia'
                if 'constancia' in low:
                    return 'Constancia'
                return s

            tipo_documento_norm = _normalizar_td(tipo_documento)

            payload = {
                "folio_visita": folio_visita,
                "folio_acta": folio_acta,
                "fecha_inicio": fecha_inicio_val,
                "fecha_termino": self.entry_fecha_termino.get().strip() if hasattr(self, 'entry_fecha_termino') and self.entry_fecha_termino.get().strip() else "",
                "hora_inicio": hora_inicio_val,
                "hora_termino": self.entry_hora_termino.get().strip() if hasattr(self, 'entry_hora_termino') and self.entry_hora_termino.get().strip() else "",
                "norma": "",
                "cliente": cliente_val,
                "nfirma1": "",
                "nfirma2": "",
                "estatus": "Pendiente",
                "tipo_documento": tipo_documento_norm,
                "folios_utilizados": folios_utilizados_val,
                # Dirección seleccionada (si existe)
                "direccion": getattr(self, 'direccion_seleccionada', '') or getattr(self, 'domicilio_seleccionado', ''),
                "colonia": getattr(self, 'colonia_seleccionada', ''),
                "municipio": getattr(self, 'municipio_seleccionado', ''),
                "ciudad_estado": getattr(self, 'ciudad_seleccionada', ''),
                "cp": getattr(self, 'cp_seleccionado', '')
            }

            # DEBUG: imprimir payload que vamos a guardar
            try:
                print(f"[DEBUG] guardar_folio_historial payload: {json.dumps(payload, ensure_ascii=False)}")
            except Exception:
                print(f"[DEBUG] guardar_folio_historial payload: {payload}")

            # Guardar usando la función existente
            self.hist_create_visita(payload)
            # Persistir también en archivo de reservas (pending_folios.json)
            try:
                pf_path = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
                arr = []
                if os.path.exists(pf_path):
                    try:
                        with open(pf_path, 'r', encoding='utf-8') as f:
                            arr = json.load(f) or []
                    except Exception:
                        arr = []
                # evitar duplicados por folio_visita
                if not any(p.get('folio_visita') == payload.get('folio_visita') for p in arr):
                    arr.append(payload)
                    with open(pf_path, 'w', encoding='utf-8') as f:
                        json.dump(arr, f, ensure_ascii=False, indent=2)
                    self.pending_folios = arr
            except Exception as e:
                print(f"[WARN] No se pudo persistir reserva en pending_folios.json: {e}")
            # DEBUG: leer historial inmediatamente y confirmar último registro
            try:
                self._cargar_historial()
                print(f"[DEBUG] after guardar_folio_historial -> total historial: {len(self.historial_data)}")
                if self.historial_data:
                    print(f"[DEBUG] ultimo registro: {self.historial_data[-1]}")
            except Exception:
                pass
            messagebox.showinfo("Folio guardado", f"El folio {folio_visita} ha sido guardado como {tipo_documento} pendiente.")
            # Preparar siguiente folio
            try:
                self.crear_nueva_visita()
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el folio: {e}")
        finally:
            # Refrescar lista de folios pendientes
            try:
                # Programar el refresco para que ocurra después de que `hist_create_visita`
                # haya tenido oportunidad de aplicar la visita en el hilo principal.
                try:
                    self.after(250, self._refresh_pending_folios_dropdown)
                except Exception:
                    # Fallback síncrono si after no está disponible
                    self._refresh_pending_folios_dropdown()
            except Exception:
                pass

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
        """Carga los datos del historial desde el archivo JSON con validación"""
        try:
            # Crear directorio si no existe
            os.makedirs(os.path.dirname(self.historial_path), exist_ok=True)
            
            if os.path.exists(self.historial_path):
                with open(self.historial_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Extraer solo las visitas
                    self.historial_data = data.get("visitas", [])
                    self.historial = data  # CARGAR EL DICCIONARIO COMPLETO
                    
                    # Validar que los datos sean consistentes
                    if not isinstance(self.historial_data, list):
                        self.historial_data = []
                    
                    # Log de carga exitosa
                    print(f"✅ Historial cargado: {len(self.historial_data)} registros desde {self.historial_path}")
            else:
                self.historial_data = []
                self.historial = {"visitas": []}
                print(f"📝 Archivo de historial no existe, se creará uno nuevo")
                
            # Inicializar también historial_data_original
            self.historial_data_original = self.historial_data.copy()
            # Normalizar campo 'tipo_documento' en los registros existentes
            try:
                cambios = False
                for reg in self.historial_data:
                    if not isinstance(reg, dict):
                        continue
                    td = reg.get('tipo_documento')
                    # Normalizar a las formas amigables: Dictamen, Negación de Dictamen, Constancia, Negación de Constancia
                    def _normalizar_td(raw):
                        if not raw:
                            return 'Dictamen'
                        s = str(raw).strip().lower()
                        if 'dictamen' in s and ('neg' in s or 'negación' in s or 'negacion' in s):
                            return 'Negación de Dictamen'
                        if 'dictamen' in s:
                            return 'Dictamen'
                        if 'constancia' in s and ('neg' in s or 'negación' in s or 'negacion' in s):
                            return 'Negación de Constancia'
                        if 'constancia' in s:
                            return 'Constancia'
                        # Default: capitalize first letter
                        return str(raw).strip()

                    nuevo_td = _normalizar_td(td)
                    if nuevo_td != td:
                        reg['tipo_documento'] = nuevo_td
                        cambios = True
                if cambios:
                    # Actualizar estructura y persistir cambios
                    self.historial['visitas'] = self.historial_data
                    self._guardar_historial()
            except Exception:
                pass
                
        except json.JSONDecodeError as e:
            print(f"❌ Error decodificando JSON: {e}")
            # Intentar recuperar del backup
            backup_path = self.historial_path + ".backup"
            if os.path.exists(backup_path):
                try:
                    print(f"🔄 Recuperando desde backup...")
                    with open(backup_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        self.historial_data = data.get("visitas", [])
                        self.historial = data
                except Exception:
                    self.historial_data = []
                    self.historial = {"visitas": []}
            else:
                self.historial_data = []
                self.historial = {"visitas": []}
        except Exception as e:
            print(f"❌ Error cargando historial: {e}")
            self.historial_data = []
            self.historial_data_original = []
            self.historial = {"visitas": []}

    def _sincronizar_historial(self):
        """Sincroniza los datos en memoria con el archivo JSON para asegurar persistencia"""
        try:
            # Actualizar self.historial con los datos actuales de historial_data
            self.historial["visitas"] = self.historial_data
            
            # Guardar el archivo
            self._guardar_historial()
            
            # Actualizar original
            self.historial_data_original = self.historial_data.copy()
            # Regenerar cache exportable para Excel (persistente)
            try:
                self._generar_datos_exportable()
            except Exception:
                pass
            
            return True
        except Exception as e:
            print(f"❌ Error sincronizando historial: {e}")
            return False

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
        """Guarda el historial en un único archivo con validación de persistencia"""
        try:
            # ACTUALIZAR self.historial_data DESDE self.historial
            self.historial_data = self.historial.get("visitas", [])
            self.historial_data_original = self.historial_data.copy()
            
            # Determinar ruta de guardado (soporte para .exe congelado y rutas no escribibles)
            target_path = self.historial_path
            try:
                base_dir = os.path.dirname(self.historial_path)
                os.makedirs(base_dir, exist_ok=True)
            except Exception:
                base_dir = None

            try:
                # Si estamos en un ejecutable congelado (PyInstaller) o el directorio no es escribible,
                # redirigir a APPDATA\GeneradorDictamenes para persistencia.
                if getattr(sys, 'frozen', False) or (base_dir and not os.access(base_dir, os.W_OK)):
                    alt_base = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GeneradorDictamenes')
                    os.makedirs(alt_base, exist_ok=True)
                    target_path = os.path.join(alt_base, os.path.basename(self.historial_path))
                    # actualizar historial_path para futuras operaciones
                    self.historial_path = target_path
            except Exception:
                target_path = self.historial_path

            # Guardar con respaldo (backup)
            backup_path = target_path + ".backup"
            if os.path.exists(target_path):
                try:
                    shutil.copy2(target_path, backup_path)
                except Exception:
                    pass

            # Escribir archivo principal
            with open(target_path, "w", encoding="utf-8") as f:
                json.dump(self.historial, f, ensure_ascii=False, indent=2)

            # Verificar que se escribió correctamente
            if os.path.exists(target_path):
                with open(target_path, 'r', encoding='utf-8') as f:
                    verificacion = json.load(f)
                    if verificacion.get('visitas'):
                        lbl = getattr(self, 'hist_info_label', None)
                        if lbl and hasattr(lbl, 'winfo_exists') and lbl.winfo_exists():
                            try:
                                lbl.configure(text=f"✅ Guardado — {len(self.historial_data)} registros")
                            except Exception:
                                pass
            else:
                print("⚠️ Error: No se pudo verificar el archivo guardado")
            print(f"✅ Historial guardado: {len(self.historial_data)} registros (ruta: {target_path})")
            
        except Exception as e:
            print(f"❌ Error guardando historial: {e}")
            lbl = getattr(self, 'hist_info_label', None)
            if lbl and hasattr(lbl, 'configure'):
                try:
                    lbl.configure(text=f"Error guardando: {e}")
                except Exception:
                    pass
    
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
        # Si el Treeview está presente, vaciarlo (operación rápida)
        try:
            if hasattr(self, 'hist_tree') and self.hist_tree is not None:
                try:
                    self.hist_tree.delete(*self.hist_tree.get_children())
                except Exception:
                    pass
                return
        except Exception:
            pass

        # Fallback: recrear el scrollable frame para versiones antiguas
        old = getattr(self, 'hist_scroll', None)
        parent = old.master if old is not None else None
        if parent is None:
            return

        new_scroll = ctk.CTkScrollableFrame(
            parent,
            fg_color=STYLE["fondo"],
            scrollbar_button_color=STYLE["primario"],
            scrollbar_button_hover_color=STYLE["primario"]
        )
        new_scroll.pack(fill="both", expand=True, padx=0, pady=0)
        self.hist_scroll = new_scroll

        if old is not None and isinstance(old, ctk.CTkScrollableFrame):
            try:
                self.after(250, old.destroy)
            except Exception:
                try:
                    old.destroy()
                except Exception:
                    pass

    # -- BOTONES DE ACCION PARA CADA VISITA -- #
    def _poblar_historial_ui(self):
        """Poblar historial usando Treeview virtualizado (más eficiente)."""
        # Cargar datos solo si no existen o si se solicita recarga (permite que búsquedas filtradas persistan)
        if (not hasattr(self, 'historial_data') or not self.historial_data) or getattr(self, '_force_reload_hist', False):
            self._cargar_historial()
            self._force_reload_hist = False
        regs = getattr(self, 'historial_data', []) or []

        total_registros = len(regs)
        regs_pagina = self.HISTORIAL_REGS_POR_PAGINA
        pagina_actual = getattr(self, 'HISTORIAL_PAGINA_ACTUAL', 1)
        total_paginas = max(1, (total_registros + regs_pagina - 1) // regs_pagina)
        inicio = (pagina_actual - 1) * regs_pagina
        fin = min(inicio + regs_pagina, total_registros)

        # actualizar controles de paginación si existen
        try:
            if hasattr(self, 'hist_pagina_label'):
                self.hist_pagina_label.configure(text=f"Página {pagina_actual} de {total_paginas}")
            if hasattr(self, 'btn_hist_prev'):
                self.btn_hist_prev.configure(state="normal" if pagina_actual > 1 else "disabled")
            if hasattr(self, 'btn_hist_next'):
                self.btn_hist_next.configure(state="normal" if pagina_actual < total_paginas else "disabled")
        except Exception:
            pass

        # Vaciar Treeview actual
        try:
            self.hist_tree.delete(*self.hist_tree.get_children())
        except Exception:
            pass

        # Insertar registros de la página actual
        for idx in range(inicio, fin):
            registro = regs[idx]
            hora_inicio = self._formatear_hora_12h(registro.get('hora_inicio', ''))
            hora_termino = self._formatear_hora_12h(registro.get('hora_termino', ''))

            folios_str = registro.get('folios_utilizados', '') or ''
            if not folios_str or folios_str in ('0', '-'):
                folios_display = ''
            else:
                folios_display = self._formatear_folios_rango(folios_str)

            cliente_short = self._acortar_texto(registro.get('cliente', '-'), 20)
            nfirma1_short = self._acortar_texto(registro.get('nfirma1', 'N/A'), 12)

            datos = [
                registro.get('folio_visita', '-') or '-',
                registro.get('folio_acta', '-') or '-',
                registro.get('fecha_inicio', '-') or '-',
                registro.get('fecha_termino', '-') or '-',
                hora_inicio or '-',
                hora_termino or '-',
                cliente_short,
                nfirma1_short,
                registro.get('tipo_documento', '-') or '-',
                registro.get('estatus', 'Completado') or 'Completado',
                folios_display,
                "📁 Folios  •  📎 Archivos  •  ✏️ Editar  •  🗑️ Borrar"
            ]

            # Insertar en tree
            iid = f"h_{idx}"
            try:
                self.hist_tree.insert('', 'end', iid=iid, values=datos)
                self._hist_map[iid] = registro
            except Exception:
                pass

        # actualizar info pie
        try:
            if hasattr(self, 'hist_info_label'):
                self.hist_info_label.configure(
                    text=f"Registros: {total_registros} | Sistema V&C - Generador de Dictámenes de Comprimiento"
                )
        except Exception:
            pass

    def _hist_show_context_menu(self, event):
        try:
            iid = self.hist_tree.identify_row(event.y)
            if not iid:
                return
            self.hist_tree.selection_set(iid)
            self._hist_context_selected = iid
            self.hist_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.hist_context_menu.grab_release()
            except Exception:
                pass

    def _hist_on_double_click(self, event):
        iid = self.hist_tree.identify_row(event.y)
        if not iid:
            return
        reg = self._hist_map.get(iid)
        if reg:
            try:
                self.mostrar_opciones_documentos(reg)
            except Exception:
                pass

    def _hist_on_left_click(self, event):
        # Si el usuario hizo click en la columna de Acciones, abrir menú contextual
        try:
            col = self.hist_tree.identify_column(event.x)
            # columnas vienen como '#1', '#2', ...
            cols = list(self.hist_tree['columns'])
            if not cols:
                return
            last_index = len(cols)
            if col == f"#{last_index}":
                iid = self.hist_tree.identify_row(event.y)
                if not iid:
                    return
                self.hist_tree.selection_set(iid)
                self._hist_context_selected = iid
                try:
                    self.hist_context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    try:
                        self.hist_context_menu.grab_release()
                    except Exception:
                        pass
        except Exception:
            pass

    def _hist_menu_action(self, action):
        iid = getattr(self, '_hist_context_selected', None)
        if not iid:
            sel = self.hist_tree.selection()
            iid = sel[0] if sel else None
        if not iid:
            return
        reg = self._hist_map.get(iid)
        if not reg:
            return
        try:
            if action == 'folios':
                self.descargar_folios_visita(reg)
            elif action == 'archivos':
                self.mostrar_opciones_documentos(reg)
            elif action == 'editar':
                self.hist_editar_registro(reg)
            elif action == 'borrar':
                self.hist_eliminar_registro(reg)
        except Exception:
            pass

    def hist_pagina_anterior(self):
        if self.HISTORIAL_PAGINA_ACTUAL > 1:
            self.HISTORIAL_PAGINA_ACTUAL -= 1
            self._poblar_historial_ui()

    def hist_pagina_siguiente(self):
        total_registros = len(self.historial_data)
        total_paginas = max(1, (total_registros + self.HISTORIAL_REGS_POR_PAGINA - 1) // self.HISTORIAL_REGS_POR_PAGINA)
        if self.HISTORIAL_PAGINA_ACTUAL < total_paginas:
            self.HISTORIAL_PAGINA_ACTUAL += 1
            self._poblar_historial_ui()

    def _formatear_folios_rango(self, folios_str):
        """Formatea los folios para mostrar solo el rango (inicio-fin)"""
        if not folios_str or folios_str == '0' or folios_str == '-':
            return '-'
        
        # Si ya es un rango simple
        if ' - ' in folios_str and not folios_str.startswith('Total:'):
            # Extraer solo el rango (puede venir como "000001 - 000010")
            return folios_str
        
        # Si es un solo folio
        if folios_str.startswith('Folio: '):
            return folios_str.replace('Folio: ', '')
        
        # Si tiene formato de total con lista
        if folios_str.startswith('Total: '):
            # Intentar extraer folios de la lista
            if '| Folios:' in folios_str:
                try:
                    partes = folios_str.split('| Folios:')
                    if len(partes) > 1:
                        folios_lista = partes[1].strip().split(', ')
                        if folios_lista:
                            # Obtener primer y último folio
                            primer = folios_lista[0].strip()
                            ultimo = folios_lista[-1].strip().replace('...', '').strip()
                            if primer and ultimo and primer != ultimo:
                                return f"{primer} - {ultimo}"
                            elif primer:
                                return primer
                except:
                    pass
            
            # Si no se pudo extraer, mostrar solo el total
            try:
                total_part = folios_str.split('|')[0].strip()
                return total_part.replace('Total: ', '') + ' folios'
            except:
                return folios_str
        
        # Si tiene muchos folios separados por comas
        if ',' in folios_str:
            folios_list = [f.strip() for f in folios_str.split(',') if f.strip()]
            if len(folios_list) > 1:
                return f"{folios_list[0]} - {folios_list[-1]}"
            elif folios_list:
                return folios_list[0]
        
        return folios_str[:20] + ('...' if len(folios_str) > 20 else '')

    def _acortar_texto(self, texto, max_caracteres=20):
        """Acorta el texto si es muy largo, agregando '...' al final"""
        if not texto:
            return ""
        
        texto_str = str(texto)
        if len(texto_str) <= max_caracteres:
            return texto_str
        
        return texto_str[:max_caracteres-3] + "..."

    def _formatear_hora_12h(self, hora_str):
        """Formatea una hora a formato 12h (AM/PM)"""
        try:
            if ":" in hora_str:
                partes = hora_str.split(":")
                if len(partes) >= 2:
                    horas = int(partes[0])
                    minutos = int(partes[1])
                    
                    periodo = "AM" if horas < 12 else "PM"
                    horas_12 = horas if horas <= 12 else horas - 12
                    if horas_12 == 0:
                        horas_12 = 12
                    
                    return f"{horas_12}:{minutos:02d} {periodo}"
        except:
            pass
        
        return hora_str

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
            try:
                if tipo == "acta":
                    folio = registro.get('folio_visita', '')
                    if not folio:
                        messagebox.showwarning("Error", "No se encontró el folio de la visita para generar el acta.")
                        return

                    # Preferir el backup más reciente en data/tabla_relacion_backups si existe
                    data_dir_local = os.path.join(os.path.dirname(__file__), 'data')
                    backups_dir = os.path.join(data_dir_local, 'tabla_relacion_backups')
                    tabla_dest = os.path.join(data_dir_local, 'tabla_de_relacion.json')

                    if os.path.exists(backups_dir):
                        archivos = [os.path.join(backups_dir, f) for f in os.listdir(backups_dir) if f.lower().endswith('.json')]
                        if archivos:
                            latest = max(archivos, key=os.path.getmtime)
                            try:
                                shutil.copy2(latest, tabla_dest)
                                print(f"📁 Usando backup de tabla de relación: {latest}")
                            except Exception as e:
                                print(f"⚠️ No se pudo copiar backup de tabla de relación: {e}")

                    # Pedir al usuario dónde guardar el PDF (Explorador de archivos)
                    default_name = f"Acta_{folio}.pdf"
                    save_path = filedialog.asksaveasfilename(
                        title="Guardar Acta de Inspección",
                        defaultextension=".pdf",
                        filetypes=[("PDF", "*.pdf")],
                        initialfile=default_name
                    )

                    if not save_path:
                        return

                    # Importar dinámicamente el generador de actas y generar
                    try:
                        import importlib.util
                        acta_file = os.path.join(os.path.dirname(__file__), 'Documentos Inspeccion', 'Acta_inspeccion.py')
                        if not os.path.exists(acta_file):
                            messagebox.showerror("Error", f"No se encontró el generador de actas: {acta_file}")
                            return

                        spec = importlib.util.spec_from_file_location("Acta_inspeccion", acta_file)
                        if spec is None or getattr(spec, 'loader', None) is None:
                            messagebox.showerror("Error", f"No se pudo cargar el módulo Acta_inspeccion (spec inválido): {acta_file}")
                            return
                        acta_mod = importlib.util.module_from_spec(spec)
                        import sys
                        # Registrar el módulo temporalmente para permitir reload() desde el código
                        sys.modules["Acta_inspeccion"] = acta_mod
                        spec.loader.exec_module(acta_mod)

                        # Generar acta para el folio y guardarla en la ruta indicada
                        ruta_generada = acta_mod.generar_acta_desde_visita(folio_visita=folio, ruta_salida=save_path)

                        # Persistir la ruta del acta en el historial (si corresponde)
                        try:
                            for v in self.historial.get('visitas', []):
                                if v.get('folio_visita') == folio:
                                    v['ruta_acta'] = save_path
                                    break
                            # Guardar historial
                            self._guardar_historial()
                        except Exception as e:
                            print(f"⚠️ Error guardando ruta de acta en historial: {e}")

                        messagebox.showinfo("Acta generada", f"Acta guardada en:\n{ruta_generada}")
                        return
                    except Exception as e:
                        messagebox.showerror("Error", f"Error generando acta:\n{e}")
                        return

                # Otros documentos: Oficio y Formato -> generar desde módulos correspondientes
                folio = registro.get('folio_visita', '')
                if not folio:
                    messagebox.showwarning("Error", "No se encontró el folio de la visita para generar el documento.")
                    return

                # Asegurar que usamos el backup más reciente para tabla_de_relacion
                data_dir_local = os.path.join(os.path.dirname(__file__), 'data')
                backups_dir = os.path.join(data_dir_local, 'tabla_relacion_backups')
                tabla_dest = os.path.join(data_dir_local, 'tabla_de_relacion.json')

                if os.path.exists(backups_dir):
                    archivos = [os.path.join(backups_dir, f) for f in os.listdir(backups_dir) if f.lower().endswith('.json')]
                    if archivos:
                        latest = max(archivos, key=os.path.getmtime)
                        try:
                            shutil.copy2(latest, tabla_dest)
                            print(f"📁 Usando backup de tabla de relación: {latest}")
                        except Exception as e:
                            print(f"⚠️ No se pudo copiar backup de tabla de relación: {e}")

                # Calcular lista de folios asociados a la visita (int)
                folios_list = []
                try:
                    fol_num = ''.join([c for c in folio if c.isdigit()])
                    folios_file = os.path.join(data_dir_local, 'folios_visitas', f'folios_{fol_num}.json')
                    if os.path.exists(folios_file):
                        with open(folios_file, 'r', encoding='utf-8') as ff:
                            data_f = json.load(ff)
                            if isinstance(data_f, list):
                                folios_list = [int(x) for x in data_f if str(x).isdigit()]
                            elif isinstance(data_f, dict) and 'folios' in data_f:
                                folios_list = [int(x) for x in data_f.get('folios', []) if str(x).isdigit()]
                except Exception:
                    folios_list = []

                # fallback: parsear rango en registro
                if not folios_list:
                    fu = registro.get('folios_utilizados') or ''
                    if fu and isinstance(fu, str):
                        if '-' in fu:
                            parts = [p.strip() for p in fu.split('-')]
                            try:
                                start = int(parts[0])
                                end = int(parts[1]) if len(parts) > 1 else start
                                folios_list = list(range(start, end+1))
                            except Exception:
                                folios_list = []
                        elif ',' in fu:
                            vals = [p.strip() for p in fu.split(',')]
                            for v in vals:
                                if v.isdigit():
                                    folios_list.append(int(v))

                # Cargar tabla_de_relacion y extraer solicitudes únicas y FECHA DE VERIFICACION
                solicitudes = set()
                fecha_verificacion = None
                try:
                    if os.path.exists(tabla_dest):
                        with open(tabla_dest, 'r', encoding='utf-8') as tf:
                            tabla = json.load(tf)
                            for rec in tabla:
                                try:
                                    fol = rec.get('FOLIO')
                                    fol_int = int(fol) if fol is not None and str(fol).isdigit() else None
                                except Exception:
                                    fol_int = None
                                if folios_list and fol_int in folios_list:
                                    sol = rec.get('SOLICITUD') or rec.get('SOLICITUDES')
                                    if sol:
                                        solicitudes.add(str(sol).strip())
                                    if not fecha_verificacion and rec.get('FECHA DE VERIFICACION'):
                                        fecha_verificacion = rec.get('FECHA DE VERIFICACION')
                            # si no filtró por folios, intentar colectar solicitudes globales
                            if not solicitudes and isinstance(tabla, list):
                                for rec in tabla:
                                    sol = rec.get('SOLICITUD') or rec.get('SOLICITUDES')
                                    if sol:
                                        solicitudes.add(str(sol).strip())
                except Exception as e:
                    print(f"⚠️ Error leyendo tabla de relación para generar documento: {e}")

                # Formatear fecha_verificacion si está en ISO
                fecha_formateada = None
                if fecha_verificacion:
                    try:
                        if '-' in fecha_verificacion:
                            dt = datetime.strptime(fecha_verificacion[:10], '%Y-%m-%d')
                        else:
                            dt = datetime.strptime(fecha_verificacion[:10], '%d/%m/%Y')
                        fecha_formateada = dt.strftime('%d/%m/%Y')
                    except Exception:
                        fecha_formateada = fecha_verificacion

                # Preparar nombre por defecto y pedir ruta
                default_name = f"{nombre.replace(' ', '_')}_{folio}.pdf"
                save_path = filedialog.asksaveasfilename(
                    title=f"Guardar {nombre}",
                    defaultextension=".pdf",
                    filetypes=[("PDF", "*.pdf")],
                    initialfile=default_name
                )
                if not save_path:
                    return

                # Importar dinámicamente y generar según tipo
                try:
                    import importlib.util
                    if tipo == 'formato':
                        formato_file = os.path.join(os.path.dirname(__file__), 'Documentos Inspeccion', 'Formato_supervision.py')
                        if not os.path.exists(formato_file):
                            messagebox.showerror('Error', f'No se encontró el módulo: {formato_file}')
                            return
                        spec = importlib.util.spec_from_file_location('Formato_supervision', formato_file)
                        if spec is None or getattr(spec, 'loader', None) is None:
                            messagebox.showerror('Error', f'No se pudo cargar el módulo Formato_supervision (spec inválido): {formato_file}')
                            return
                        mod = importlib.util.module_from_spec(spec)
                        import sys
                        sys.modules['Formato_supervision'] = mod
                        spec.loader.exec_module(mod)

                        datos = {
                            'solicitud': ', '.join(sorted(list(solicitudes))) if solicitudes else registro.get('folio_visita',''),
                            'servicio': None,
                            'fecha': registro.get('fecha_inicio') or registro.get('fecha_termino') or datetime.now().strftime('%d/%m/%Y'),
                            'cliente': registro.get('cliente',''),
                            'supervisor': 'Mario Terrez Gonzalez'
                        }
                        # Determinar servicio a partir de tabla (si hay alguna entrada con TIP0...)
                        try:
                            # buscar primer registro en tabla_dest que corresponda a folios_list
                            if os.path.exists(tabla_dest):
                                with open(tabla_dest, 'r', encoding='utf-8') as tf:
                                    tabla = json.load(tf)
                                    for rec in tabla:
                                        try:
                                            fol = rec.get('FOLIO')
                                            fol_int = int(fol) if fol is not None and str(fol).isdigit() else None
                                        except Exception:
                                            fol_int = None
                                        if folios_list and fol_int in folios_list:
                                            tipo_doc = (rec.get('TIPO DE DOCUMENTO') or rec.get('TIPO_DE_DOCUMENTO') or 'D')
                                            datos['servicio'] = 'DICTAMEN' if str(tipo_doc).strip().upper() == 'D' else str(tipo_doc)
                                            break
                        except Exception:
                            datos['servicio'] = datos.get('servicio') or 'DICTAMEN'

                        # Llamar al generador
                        try:
                            mod.generar_supervision(datos, save_path)
                        except Exception as e:
                            messagebox.showerror('Error', f'Error generando Formato de Supervisión:\n{e}')
                            return

                    elif tipo == 'oficio':
                        # Prefer a fixed fallback module if present to avoid importing a corrupted original
                        oficio_file = os.path.join(os.path.dirname(__file__), 'Documentos Inspeccion', 'Oficio_comision.py')
                        oficio_fixed = os.path.join(os.path.dirname(__file__), 'Documentos Inspeccion', 'Oficio_comision_fixed.py')
                        if os.path.exists(oficio_fixed):
                            oficio_file = oficio_fixed
                        if not os.path.exists(oficio_file):
                            messagebox.showerror('Error', f'No se encontró el módulo: {oficio_file}')
                            return
                        spec = importlib.util.spec_from_file_location('Oficio_comision', oficio_file)
                        if spec is None or getattr(spec, 'loader', None) is None:
                            messagebox.showerror('Error', f'No se pudo cargar el módulo Oficio_comision (spec inválido): {oficio_file}')
                            return
                        mod = importlib.util.module_from_spec(spec)
                        import sys
                        sys.modules['Oficio_comision'] = mod
                        spec.loader.exec_module(mod)

                        # Preferir usar la función de preparación del propio módulo si existe
                        datos_oficio = None
                        try:
                            if hasattr(mod, 'preparar_datos_desde_visita'):
                                datos_oficio = mod.preparar_datos_desde_visita(registro)
                            else:
                                # Heurística local: priorizar 'calle_numero' y anexar CP a colonia
                                calle = registro.get('calle_numero') or registro.get('direccion','') or ''
                                colonia = registro.get('colonia','') or ''
                                cp = registro.get('cp') or registro.get('CP') or ''
                                if cp and colonia:
                                    colonia = f"{colonia}, {cp}"
                                datos_oficio = {
                                    'no_oficio': registro.get('folio_visita',''),
                                    'fecha_inspeccion': fecha_formateada or registro.get('fecha_termino') or datetime.now().strftime('%d/%m/%Y'),
                                    'normas': registro.get('norma','').split(', ') if registro.get('norma') else [],
                                    'empresa_visitada': registro.get('cliente',''),
                                    'calle_numero': calle,
                                    'colonia': colonia,
                                    'municipio': registro.get('municipio',''),
                                    'ciudad_estado': registro.get('ciudad_estado',''),
                                    'fecha_confirmacion': registro.get('fecha_inicio') or datetime.now().strftime('%d/%m/%Y'),
                                    'medio_confirmacion': 'correo electrónico',
                                    'inspectores': [s.strip() for s in (registro.get('supervisores_tabla') or registro.get('nfirma1') or '').split(',') if s.strip()],
                                    'observaciones': registro.get('observaciones',''),
                                    'num_solicitudes': ', '.join(sorted(list(solicitudes))) if solicitudes else ''
                                }
                        except Exception:
                            datos_oficio = {
                                'no_oficio': registro.get('folio_visita',''),
                                'fecha_inspeccion': fecha_formateada or registro.get('fecha_termino') or datetime.now().strftime('%d/%m/%Y'),
                                'normas': registro.get('norma','').split(', ') if registro.get('norma') else [],
                                'empresa_visitada': registro.get('cliente',''),
                                'calle_numero': registro.get('calle_numero') or registro.get('direccion',''),
                                'colonia': registro.get('colonia',''),
                                'municipio': registro.get('municipio',''),
                                'ciudad_estado': registro.get('ciudad_estado',''),
                                'fecha_confirmacion': registro.get('fecha_inicio') or datetime.now().strftime('%d/%m/%Y'),
                                'medio_confirmacion': 'correo electrónico',
                                'inspectores': [s.strip() for s in (registro.get('supervisores_tabla') or registro.get('nfirma1') or '').split(',') if s.strip()],
                                'observaciones': registro.get('observaciones',''),
                                'num_solicitudes': ', '.join(sorted(list(solicitudes))) if solicitudes else ''
                            }

                        try:
                            mod.generar_oficio_pdf(datos_oficio, save_path)
                        except TypeError as e:
                            # Intentar recargar el módulo y reintentar: puede ocurrir si el archivo fue editado
                            try:
                                import importlib
                                importlib.reload(mod)
                                mod.generar_oficio_pdf(datos_oficio, save_path)
                            except Exception as e2:
                                messagebox.showerror('Error', f'Error generando Oficio de Comisión:\n{e2}')
                                return
                        except Exception as e:
                            messagebox.showerror('Error', f'Error generando Oficio de Comisión:\n{e}')
                            return

                    # Persistir la ruta en historial
                    try:
                        for v in self.historial.get('visitas', []):
                            if v.get('folio_visita') == folio:
                                key = 'ruta_' + nombre.replace(' ', '_').lower()
                                v[key] = save_path
                                break
                        self._guardar_historial()
                    except Exception as e:
                        print(f"⚠️ Error guardando ruta en historial: {e}")

                    messagebox.showinfo(f"{nombre} generado", f"{nombre} guardado en:\n{save_path}")
                    return
                except Exception as e:
                    messagebox.showerror("Error", f"Error generando documento {nombre}:\n{e}")
                    return
            except Exception as e:
                messagebox.showerror("Error", f"Error al procesar descarga:\n{e}")
        
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
    # ----------------- FOLIOS PENDIENTES (UI helpers) -----------------
    def _get_folios_pendientes(self):
        """Retorna lista de registros pendientes.

        Los pendientes se cargan preferentemente desde `data/pending_folios.json` si existe
        (persistencia explícita de reservas). Si no existe, se hace fallback a `historial_data`.
        """
        try:
            # Si cargamos reservas persistentes, trabajar con esa lista
            if hasattr(self, 'pending_folios') and isinstance(self.pending_folios, list):
                return list(self.pending_folios)

            pendientes = []
            fuente = getattr(self, 'historial_data', []) or []
            for r in fuente:
                est = (r.get('estatus','') or '').strip().lower()
                if est != 'pendiente':
                    continue
                pendientes.append(r)
            return pendientes
        except Exception:
            return []

    def _refresh_pending_folios_dropdown(self):
        """Actualiza los valores del combobox de folios pendientes."""
        try:
            # Intentar filtrar por el tipo de documento actualmente seleccionado
            tipo_sel = None
            try:
                tipo_sel = self.combo_tipo_documento.get().strip() if hasattr(self, 'combo_tipo_documento') else None
            except Exception:
                tipo_sel = None

            # Cargar pendientes persistentes si existen
            pendientes_source = []
            try:
                # cargar desde memoria si ya leido
                if hasattr(self, 'pending_folios') and isinstance(self.pending_folios, list):
                    pendientes_source = list(self.pending_folios)
                else:
                    # intentar leer archivo
                    pf = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
                    if os.path.exists(pf):
                        with open(pf, 'r', encoding='utf-8') as f:
                            pendientes_source = json.load(f) or []
                    else:
                        pendientes_source = list(getattr(self, 'historial_data', []) or [])
            except Exception:
                pendientes_source = list(getattr(self, 'historial_data', []) or [])

            pendientes = []
            for r in pendientes_source:
                try:
                    est = (r.get('estatus','') or '').strip().lower()
                    td = (r.get('tipo_documento','') or '').strip()
                    # Mostrar todas las reservas pendientes independientemente del tipo
                    if est != 'pendiente':
                        continue
                    pendientes.append(r)
                except Exception:
                    continue
            vals = []
            pendientes_map = {}
            # Construir valores legibles e índices alternativos por folio/acta
            for i, r in enumerate(pendientes):
                fol = r.get('folio_visita','')
                act = r.get('folio_acta','')
                cliente = r.get('cliente','')
                fecha = r.get('fecha_inicio','')
                tipo = r.get('tipo_documento','') or ''
                display = f"{fol} — {tipo} — {cliente} ({fecha})"
                vals.append(display)
                # mapear la etiqueta completa
                pendientes_map[display] = r
                # mapear por folio exacto (con y sin prefijo)
                try:
                    if fol:
                        pendientes_map[fol] = r
                        # sin prefijo CP/AC
                        no_pref = fol
                        if fol.upper().startswith('CP') or fol.upper().startswith('AC'):
                            no_pref = fol[2:]
                        pendientes_map[no_pref] = r
                except Exception:
                    pass
                # mapear por folio de acta
                try:
                    if act:
                        pendientes_map[act] = r
                except Exception:
                    pass

            self._pendientes_map = pendientes_map

            if hasattr(self, 'combo_folios_pendientes'):
                try:
                    self.combo_folios_pendientes.configure(values=vals)
                    # Mostrar placeholder claro cuando no hay valores
                    if vals:
                        # Dejar sin seleccionar para evitar consumir automáticamente
                        try:
                            # Si hay una selección pendiente activa, mantenerla visible
                            sel_id = getattr(self, 'selected_pending_id', None)
                            if sel_id:
                                # encontrar display asociado
                                display_to_set = None
                                for d in vals:
                                    r = pendientes_map.get(d)
                                    try:
                                        if r and (r.get('_id') == sel_id or r.get('id') == sel_id):
                                            display_to_set = d
                                            break
                                    except Exception:
                                        continue
                                if display_to_set:
                                    self.combo_folios_pendientes.set(display_to_set)
                                else:
                                    self.combo_folios_pendientes.set("")
                            else:
                                self.combo_folios_pendientes.set("")
                        except Exception:
                            self.combo_folios_pendientes.set(vals[0])
                    else:
                        self.combo_folios_pendientes.set("")
                except Exception:
                    pass
        except Exception as e:
            print(f"Error refrescando folios pendientes: {e}")

    def _seleccionar_folio_pendiente(self, seleccionado_text):
        """Al seleccionar un folio pendiente, cargar sus datos en el formulario y marcar como En proceso."""
        try:
            if not seleccionado_text:
                return
            registro = getattr(self, '_pendientes_map', {}).get(seleccionado_text)

            # DEBUG
            try:
                print(f"[DEBUG] _seleccionar_folio_pendiente seleccionado_text='{seleccionado_text}' registro_found={bool(registro)}")
            except Exception:
                pass

            # Si el usuario escribió solo el folio (p. ej. "CP0001") o el texto no coincide
            # intentar hacer una búsqueda por folio_visita, folio_acta o substring en los valores
            if not registro:
                try:
                    pendientes = self._get_folios_pendientes()
                    # búsqueda exacta por folio
                    for r in pendientes:
                        if seleccionado_text == r.get('folio_visita') or seleccionado_text == r.get('folio_acta'):
                            registro = r
                            break
                    # búsqueda por substring en la representación mostrada
                    if not registro:
                        for k, r in getattr(self, '_pendientes_map', {}).items():
                            try:
                                if seleccionado_text.strip() and seleccionado_text.strip().lower() in str(k).lower():
                                    registro = r
                                    break
                            except Exception:
                                continue
                except Exception:
                    registro = None

            if not registro:
                messagebox.showwarning("No encontrado", "No se encontró el folio pendiente seleccionado.")
                return

            # Cargar datos en la sección Información de Visita
            try:
                self.entry_folio_visita.configure(state='normal')
                self.entry_folio_visita.delete(0, 'end')
                folio_to_set = registro.get('folio_visita','')
                # Asegurar formato CP/AC si viene sin prefijo
                if folio_to_set and not (folio_to_set.upper().startswith('CP') or folio_to_set.upper().startswith('AC')):
                    # intentar deducir si corresponde a CP
                    if folio_to_set.isdigit():
                        folio_to_set = f"CP{folio_to_set.zfill(6)[-6:]}"
                self.entry_folio_visita.insert(0, folio_to_set)
                self.entry_folio_visita.configure(state='readonly')

                self.entry_folio_acta.configure(state='normal')
                self.entry_folio_acta.delete(0, 'end')
                self.entry_folio_acta.insert(0, registro.get('folio_acta',''))
                self.entry_folio_acta.configure(state='readonly')

                self.entry_fecha_inicio.delete(0, 'end')
                self.entry_fecha_inicio.insert(0, registro.get('fecha_inicio',''))
                self.entry_fecha_termino.delete(0, 'end')
                self.entry_fecha_termino.insert(0, registro.get('fecha_termino',''))

                try:
                    self.entry_hora_inicio.configure(state='normal')
                    self.entry_hora_inicio.delete(0, 'end')
                    self.entry_hora_inicio.insert(0, registro.get('hora_inicio',''))
                    self.entry_hora_inicio.configure(state='readonly')
                except Exception:
                    pass

                try:
                    self.entry_hora_termino.configure(state='normal')
                    self.entry_hora_termino.delete(0, 'end')
                    self.entry_hora_termino.insert(0, registro.get('hora_termino',''))
                    self.entry_hora_termino.configure(state='readonly')
                except Exception:
                    pass

                cliente = registro.get('cliente')
                if cliente and hasattr(self, 'combo_cliente'):
                    try:
                        self.combo_cliente.set(cliente)
                        try:
                            self.actualizar_cliente_seleccionado(cliente)
                        except Exception:
                            pass
                    except Exception:
                        pass

            except Exception as e:
                print(f"Error cargando registro pendiente en formulario: {e}")

            # No cambiar estatus en disco todavía: solo marcar en memoria que el usuario
            # seleccionó este folio pendiente. La visita permanecerá como 'Pendiente'
            # hasta que el usuario genere los documentos (o confirme su uso).
            try:
                rid = registro.get('_id') or registro.get('id')
                # Marcar folio como seleccionado para uso posterior (no persistir estatus)
                try:
                    fv = registro.get('folio_visita','') or ''
                    num = fv
                    if fv.upper().startswith('CP') or fv.upper().startswith('AC'):
                        num = fv[2:]
                    # Mantener formato de `current_folio` como 6 dígitos sin prefijo
                    num_only = ''.join([c for c in str(num) if c.isdigit()]) or ''
                    if num_only:
                        self.current_folio = num_only.zfill(6)
                    else:
                        # si no contiene dígitos, mantener el valor existente
                        pass
                    self.selected_pending_id = rid
                    self.usando_folio_reservado = True
                    try:
                        print(f"[DEBUG] seleccionado pending id={rid} folio={fv} current_folio set to {self.current_folio}")
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception as e:
                print(f"Error preparando registro seleccionado: {e}")

            try:
                self._refresh_pending_folios_dropdown()
            except Exception:
                pass

        except Exception as e:
            print(f"Error al seleccionar folio pendiente: {e}")

    def _desmarcar_folio_seleccionado(self):
        """Desmarca la selección actual sin eliminar la reserva persistente."""
        try:
            self.selected_pending_id = None
            self.usando_folio_reservado = False
            # Restaurar current_folio al cálculo normal (recalcular)
            try:
                self.cargar_ultimo_folio()
            except Exception:
                pass
            # refrescar UI
            try:
                self._refresh_pending_folios_dropdown()
            except Exception:
                pass
            messagebox.showinfo("Desmarcado", "La selección del folio reservado ha sido desactivada. La reserva se mantiene hasta que se utilice o se elimine.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo desmarcar la selección: {e}")

        except Exception as e:
            print(f"Error al seleccionar folio pendiente: {e}")

    def _eliminar_folio_pendiente(self):
        """Eliminar el folio pendiente seleccionado (con confirmación)."""
        try:
            seleccionado_text = None
            try:
                seleccionado_text = self.combo_folios_pendientes.get()
            except Exception:
                pass
            if not seleccionado_text:
                messagebox.showwarning("Seleccionar folio", "Seleccione un folio pendiente para eliminar.")
                return

            registro = getattr(self, '_pendientes_map', {}).get(seleccionado_text)
            # Intento fallback si el usuario solo escribió el folio
            if not registro:
                try:
                    pendientes = self._get_folios_pendientes()
                    for r in pendientes:
                        if seleccionado_text == r.get('folio_visita') or seleccionado_text == r.get('folio_acta'):
                            registro = r
                            break
                    if not registro:
                        for k, r in getattr(self, '_pendientes_map', {}).items():
                            if seleccionado_text.strip() and seleccionado_text.strip() in k:
                                registro = r
                                break
                except Exception:
                    registro = None

            if not registro:
                messagebox.showwarning("No encontrado", "No se encontró el folio seleccionado.")
                return

            if not messagebox.askyesno("Confirmar eliminación", f"¿Eliminar el folio pendiente {registro.get('folio_visita')}? Esta acción no se puede deshacer."):
                return

            try:
                self.hist_eliminar_registro(registro)
            except Exception:
                try:
                    self.historial['visitas'] = [v for v in self.historial.get('visitas', []) if v.get('folio_visita') != registro.get('folio_visita')]
                    self._guardar_historial()
                except Exception as e:
                    print(f"Error eliminando folio pendiente manualmente: {e}")

            # También eliminar de archivo de reservas si existe
            try:
                pf_path = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
                if os.path.exists(pf_path):
                    with open(pf_path, 'r', encoding='utf-8') as f:
                        arr = json.load(f) or []
                    arr = [p for p in arr if p.get('folio_visita') != registro.get('folio_visita')]
                    with open(pf_path, 'w', encoding='utf-8') as f:
                        json.dump(arr, f, ensure_ascii=False, indent=2)
                    # actualizar memoria
                    self.pending_folios = arr
            except Exception:
                pass

            try:
                self._refresh_pending_folios_dropdown()
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar el folio pendiente: {e}")

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

    # ----------------- PERSISTENCIA DE RESERVAS -----------------
    def _load_pending_folios(self):
        """Carga las reservas desde data/pending_folios.json en `self.pending_folios`."""
        try:
            pf = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
            if os.path.exists(pf):
                with open(pf, 'r', encoding='utf-8') as f:
                    arr = json.load(f) or []
                    self.pending_folios = [r for r in arr if isinstance(r, dict)]
            else:
                self.pending_folios = []
        except Exception as e:
            print(f"[WARN] Error cargando pending_folios.json: {e}")
            self.pending_folios = []

    def _save_pending_folios(self):
        """Guarda `self.pending_folios` en data/pending_folios.json."""
        try:
            pf = os.path.join(os.path.dirname(__file__), 'data', 'pending_folios.json')
            with open(pf, 'w', encoding='utf-8') as f:
                json.dump(self.pending_folios or [], f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[WARN] Error guardando pending_folios.json: {e}")

    # ----------------- Config y export persistente -----------------
    def _cargar_config_exportacion(self):
        """Carga o crea la configuración persistente para las exportaciones Excel."""
        try:
            data_folder = os.path.join(os.path.dirname(__file__), "data")
            os.makedirs(data_folder, exist_ok=True)
            cfg_path = os.path.join(data_folder, 'excel_export_config.json')
            if not os.path.exists(cfg_path):
                # Contenido por defecto
                default = {
                    "tabla_de_relacion": os.path.join(data_folder, 'tabla_de_relacion.json'),
                    "tabla_backups_dir": os.path.join(data_folder, 'tabla_relacion_backups'),
                    "clientes": os.path.join(data_folder, 'Clientes.json'),
                    "export_cache": os.path.join(data_folder, 'excel_export_data.json')
                }
                with open(cfg_path, 'w', encoding='utf-8') as f:
                    json.dump(default, f, ensure_ascii=False, indent=2)
                self.excel_export_config = default
            else:
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    self.excel_export_config = json.load(f)
            # Ensure directories exist
            os.makedirs(os.path.dirname(self.excel_export_config.get('tabla_de_relacion','') or data_folder), exist_ok=True)
            os.makedirs(self.excel_export_config.get('tabla_backups_dir', data_folder), exist_ok=True)
        except Exception as e:
            print(f"Error cargando config exportacion: {e}")
            self.excel_export_config = {}

    def _guardar_config_exportacion(self):
        try:
            data_folder = os.path.join(os.path.dirname(__file__), "data")
            cfg_path = os.path.join(data_folder, 'excel_export_config.json')
            with open(cfg_path, 'w', encoding='utf-8') as f:
                json.dump(self.excel_export_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error guardando config exportacion: {e}")

    def _generar_datos_exportable(self):
        """Genera y persiste un JSON consolidado que será la fuente para las exportaciones EMA y anual."""
        try:
            data_folder = os.path.join(os.path.dirname(__file__), "data")
            tabla_path = self.excel_export_config.get('tabla_de_relacion') or os.path.join(data_folder, 'tabla_de_relacion.json')
            clientes_path = self.excel_export_config.get('clientes') or os.path.join(data_folder, 'Clientes.json')
            export_cache = self.excel_export_config.get('export_cache') or os.path.join(data_folder, 'excel_export_data.json')

            # Cargar tabla de relación
            tabla = []
            if os.path.exists(tabla_path):
                try:
                    with open(tabla_path, 'r', encoding='utf-8') as f:
                        tabla = json.load(f)
                except Exception:
                    tabla = []

            # Cargar historial (ya en self.historial_data)
            visitas = getattr(self, 'historial_data', [])

            # Cargar clientes para enriquecer
            clientes = {}
            if os.path.exists(clientes_path):
                try:
                    with open(clientes_path, 'r', encoding='utf-8') as f:
                        cl = json.load(f)
                        if isinstance(cl, list):
                            for c in cl:
                                clientes[c.get('CLIENTE','').upper()] = c
                except Exception:
                    pass

            # Preparar estructura
            ema_rows = []
            for r in tabla:
                try:
                    cliente = r.get('EMPRESA','') or r.get('EMPRESA_VISITADA', r.get('CLIENTE',''))
                    cliente_key = (cliente or '').upper()
                    cliente_info = clientes.get(cliente_key, {})
                    # Enriquecer como en generar_reporte_ema
                    solicitud_full = r.get('ENCABEZADO', '') or r.get('SOLICITUD_ENCABEZADO', '') or r.get('SOLICITUD','')
                    sol_parts = str(solicitud_full).split()[-1] if solicitud_full else ''
                    ema_rows.append({
                        'NUMERO_SOLICITUD': sol_parts,
                        'CLIENTE': cliente,
                        'NUMERO_CONTRATO': cliente_info.get('NÚMERO_DE_CONTRATO',''),
                        'RFC': cliente_info.get('RFC',''),
                        'CURP': cliente_info.get('CURP','N/A') or 'N/A',
                        'PRODUCTO_VERIFICADO': r.get('DESCRIPCION',''),
                        'MARCAS': r.get('MARCA',''),
                        'NOM': r.get('CLASIF UVA') or r.get('CLASIF_UVA') or r.get('NOM',''),
                        'TIPO_DOCUMENTO': r.get('TIPO DE DOCUMENTO') or r.get('TIPO_DE_DOCUMENTO',''),
                        'DOCUMENTO_EMITIDO': solicitud_full,
                        'FECHA_DOCUMENTO_EMITIDO': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'VERIFICADOR': r.get('VERIFICADOR') or r.get('INSPECTOR',''),
                        'PEDIMENTO_IMPORTACION': r.get('PEDIMENTO',''),
                        'FECHA_DESADUANAMIENTO': r.get('FECHA DE ENTRADA') or r.get('FECHA_ENTRADA',''),
                        'MODELOS': r.get('CODIGO',''),
                        'FOLIO_EMA': str(r.get('FOLIO','')).zfill(6) if str(r.get('FOLIO','')).strip() else ''
                    })
                except Exception:
                    continue

            anual_rows = []
            for v in visitas:
                try:
                    anual_rows.append({
                        'FECHA_VISITA': v.get('fecha_termino') or v.get('fecha_inicio'),
                        'FOLIO_VISITA': v.get('folio_visita',''),
                        'CLIENTE': v.get('cliente',''),
                        'SOLICITUD': v.get('solicitud',''),
                        'FOLIOS_USADOS': v.get('folios_utilizados',''),
                        'NUM_SOLICITUDES': v.get('num_solicitudes',''),
                        'NORMAS': v.get('norma','')
                    })
                except Exception:
                    continue

            export_data = {
                'ema': ema_rows,
                'anual': anual_rows,
                'generated_at': datetime.now().isoformat()
            }

            # Guardar cache exportable
            try:
                with open(export_cache, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"Error guardando export cache: {e}")

            return export_data
        except Exception as e:
            print(f"Error generando datos exportable: {e}")
            return {}
    
    def descargar_excel_ema(self, registro=None):
        """Descarga el reporte EMA en Excel"""
        try:
            # Cargar el módulo control_folios_anual dinámicamente
            import importlib.util
            
            excel_gen_file = os.path.join(self.documentos_dir, 'control_folios_anual.py')
            
            if not os.path.exists(excel_gen_file):
                messagebox.showerror("Error", f"No se encontró el archivo generador de Excel: {excel_gen_file}")
                return
            
            spec = importlib.util.spec_from_file_location('control_folios_anual', excel_gen_file)
            excel_mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(excel_mod)
            
            # Preparar rutas (usar config persistente si existe)
            tabla_de_relacion_path = self.excel_export_config.get('tabla_de_relacion') if hasattr(self, 'excel_export_config') else os.path.join(self.documentos_dir, 'tabla_de_relacion.json')
            
            # Pedir ruta de guardado
            file_path = filedialog.asksaveasfilename(
                title="Guardar Reporte EMA",
                defaultextension=".xlsx",
                filetypes=[
                    ("Archivos Excel", "*.xlsx"),
                    ("Archivos Excel 97-2003", "*.xls"),
                    ("Todos los archivos", "*.*")
                ],
                initialfile=f"Reporte_EMA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            if not file_path:
                return
            
            # Si existe cache exportable, usar su sección 'ema' para generar el archivo
            export_cache = None
            if hasattr(self, 'excel_export_config'):
                export_cache = self.excel_export_config.get('export_cache')

            if export_cache and os.path.exists(export_cache):
                try:
                    with open(export_cache, 'r', encoding='utf-8') as f:
                        ec = json.load(f)
                    ema_list = ec.get('ema') if isinstance(ec, dict) else None
                    if ema_list is not None:
                        tmp_path = os.path.join(os.path.dirname(export_cache), f"_tmp_ema_{int(datetime.now().timestamp())}.json")
                        with open(tmp_path, 'w', encoding='utf-8') as tf:
                            json.dump(ema_list, tf, ensure_ascii=False, indent=2)
                        tabla_de_relacion_path_to_use = tmp_path
                    else:
                        tabla_de_relacion_path_to_use = tabla_de_relacion_path
                except Exception:
                    tabla_de_relacion_path_to_use = tabla_de_relacion_path
            else:
                tabla_de_relacion_path_to_use = tabla_de_relacion_path

            excel_mod.generar_reporte_ema(
                tabla_de_relacion_path_to_use,
                self.historial_path,
                file_path,
                export_cache=export_cache if hasattr(self, 'excel_export_config') else None
            )
            
            messagebox.showinfo("Éxito", f"Reporte EMA generado exitosamente:\n{file_path}")
            
            # Preguntar si abrir el archivo
            if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo descargado?"):
                self._abrir_archivo(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el reporte EMA:\n{str(e)}")
    
    def descargar_excel_anual(self, registro=None):
        """Descarga el reporte de control de folios anual en Excel"""
        try:
            # Cargar el módulo control_folios_anual dinámicamente
            import importlib.util
            
            excel_gen_file = os.path.join(self.documentos_dir, 'control_folios_anual.py')
            
            if not os.path.exists(excel_gen_file):
                messagebox.showerror("Error", f"No se encontró el archivo generador de Excel: {excel_gen_file}")
                return
            
            spec = importlib.util.spec_from_file_location('control_folios_anual', excel_gen_file)
            excel_mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(excel_mod)
            
            # Obtener el año actual
            year = datetime.now().year
            # Mostrar diálogo para seleccionar rango de fechas (opcional)
            def _pedir_rango_fechas_default():
                # Modal para pedir start/end date
                modal = ctk.CTkToplevel(self)
                modal.title("Rango para Control Anual")
                modal.geometry("420x180")
                modal.transient(self)
                modal.grab_set()

                ctk.CTkLabel(modal, text="Seleccione el rango de fechas (dd/mm/YYYY)\nDejar vacío para año completo:", anchor="w").pack(pady=(12,6), padx=12)

                frame = ctk.CTkFrame(modal, fg_color="transparent")
                frame.pack(fill="x", padx=12)

                ctk.CTkLabel(frame, text="Fecha inicio:").grid(row=0, column=0, sticky="w", padx=(0,6))
                ent_start = ctk.CTkEntry(frame, width=180)
                ent_start.grid(row=0, column=1, pady=6)

                ctk.CTkLabel(frame, text="Fecha fin:").grid(row=1, column=0, sticky="w", padx=(0,6))
                ent_end = ctk.CTkEntry(frame, width=180)
                ent_end.grid(row=1, column=1, pady=6)

                # Pre-fill with year bounds
                ent_start.insert(0, f"01/01/{year}")
                ent_end.insert(0, f"31/12/{year}")

                result = {"start": None, "end": None}

                def _aceptar():
                    s = ent_start.get().strip()
                    e = ent_end.get().strip()
                    result['start'] = s if s else None
                    result['end'] = e if e else None
                    modal.destroy()

                def _cancelar():
                    result['start'] = None
                    result['end'] = None
                    modal.destroy()

                btn_frame = ctk.CTkFrame(modal, fg_color="transparent")
                btn_frame.pack(fill="x", pady=8, padx=12)
                ctk.CTkButton(btn_frame, text="Aceptar", command=_aceptar, width=100).pack(side="right", padx=6)
                ctk.CTkButton(btn_frame, text="Cancelar", command=_cancelar, width=100).pack(side="right", padx=6)

                self.wait_window(modal)
                return result['start'], result['end']

            start_date, end_date = _pedir_rango_fechas_default()

            # Pedir ruta de guardado
            file_path = filedialog.asksaveasfilename(
                title="Guardar Control de Folios Anual",
                defaultextension=".xlsx",
                filetypes=[
                    ("Archivos Excel", "*.xlsx"),
                    ("Archivos Excel 97-2003", "*.xls"),
                    ("Todos los archivos", "*.*")
                ],
                initialfile=f"Control_Folios_Anual_{year}_{datetime.now().strftime('%H%M%S')}.xlsx"
            )
            
            if not file_path:
                return
            
            # Generar el reporte anual (usar backups configurados si existen)
            tabla_backups_dir = self.excel_export_config.get('tabla_backups_dir') if hasattr(self, 'excel_export_config') else os.path.join(self.documentos_dir, 'tabla_relacion_backups')

            # Si existe cache exportable, usar su sección 'anual' para alimentar el generador
            export_cache = None
            if hasattr(self, 'excel_export_config'):
                export_cache = self.excel_export_config.get('export_cache')

            if export_cache and os.path.exists(export_cache):
                try:
                    with open(export_cache, 'r', encoding='utf-8') as f:
                        ec = json.load(f)
                    anual_list = ec.get('anual') if isinstance(ec, dict) else None
                    if anual_list is not None:
                        # No escribir archivo temporal: pasamos la lista directamente
                        historial_path_to_use = self.historial_path
                        historial_list_to_pass = anual_list
                    else:
                        historial_path_to_use = self.historial_path
                        historial_list_to_pass = None
                except Exception:
                    historial_path_to_use = self.historial_path
                    historial_list_to_pass = None
            else:
                historial_path_to_use = self.historial_path
                historial_list_to_pass = None
            excel_mod.generar_control_folios_anual(
                historial_path_to_use,
                tabla_backups_dir,
                file_path,
                year,
                start_date=start_date,
                end_date=end_date,
                export_cache=export_cache,
                historial_list=historial_list_to_pass
            )
            
            messagebox.showinfo("Éxito", f"Control de Folios Anual generado exitosamente:\n{file_path}")
            
            # Preguntar si abrir el archivo
            if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo descargado?"):
                self._abrir_archivo(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el control de folios anual:\n{str(e)}")

    def hist_editar_registro(self, registro):
        """Abre el formulario para editar un registro del historial"""
        self._crear_formulario_visita(registro)

    def hist_buscar_general(self, event=None):
        """Buscar en el historial por cualquier campo"""
        try:
            # resetear paginado al buscar
            self.HISTORIAL_PAGINA_ACTUAL = 1
            # Asegurarse de que los datos estén cargados
            if not hasattr(self, 'historial_data') or not self.historial_data:
                self._cargar_historial()
                
            # Guardar copia original si no existe
            if not hasattr(self, 'historial_data_original') or not self.historial_data_original:
                self.historial_data_original = self.historial_data.copy()
            
            busqueda_raw = self.entry_buscar_general.get().strip()
            # Normalizar (quitar acentos) y bajar a minúsculas para comparaciones
            def _norm(s):
                try:
                    s2 = str(s)
                    s2 = unicodedata.normalize('NFKD', s2).encode('ASCII', 'ignore').decode('ASCII')
                    return s2.lower()
                except Exception:
                    return str(s).lower()

            busqueda = _norm(busqueda_raw)

            if not busqueda_raw:
                # Si no hay búsqueda, mostrar todos los datos
                self.historial_data = self.historial_data_original.copy()
            else:
                # Filtrar datos
                resultados = []
                for registro in self.historial_data_original:
                    # Buscar en todos los campos relevantes (añadir supervisor y tipo de documento)
                    campos_busqueda = [
                        registro.get('folio_visita', ''),
                        registro.get('folio_acta', ''),
                        registro.get('fecha_inicio', ''),
                        registro.get('fecha_termino', ''),
                        registro.get('cliente', ''),
                        registro.get('estatus', ''),
                        registro.get('folios_utilizados', ''),
                        registro.get('nfirma1', ''),
                        registro.get('nfirma2', ''),
                        registro.get('supervisor', ''),
                        registro.get('tipo_documento', '')
                    ]

                    matched = False
                    # búsqueda tradicional (substring en texto)
                    for campo in campos_busqueda:
                        try:
                            if busqueda in _norm(campo):
                                matched = True
                                break
                        except Exception:
                            continue

                    # Si no coincidió, intentar comparar solo dígitos (útil para folios con padding)
                    if not matched:
                        digits_search = ''.join([c for c in busqueda_raw if c.isdigit()])
                        if digits_search:
                            for campo in campos_busqueda:
                                campo_digits = ''.join([c for c in str(campo) if c.isdigit()])
                                if campo_digits and digits_search in campo_digits:
                                    matched = True
                                    break

                    if matched:
                        resultados.append(registro)
                
                self.historial_data = resultados
            
            self._poblar_historial_ui()
            
        except Exception as e:
            print(f"Error en búsqueda general: {e}")

    def hist_limpiar_busqueda(self):
        """Limpiar todas las búsquedas y mostrar todos los registros"""
        self.entry_buscar_general.delete(0, 'end')
        self.entry_buscar_folio.delete(0, 'end')
        
        # Recargar datos originales y resetear paginado
        self.HISTORIAL_PAGINA_ACTUAL = 1
        if hasattr(self, 'historial_data_original'):
            self.historial_data = self.historial_data_original.copy()
        else:
            self._cargar_historial()
            
        self._poblar_historial_ui()

    def _eliminar_archivos_asociados_folio(self, folio):
        """Elimina todos los archivos asociados a un folio de forma segura"""
        resultados = {"eliminados": [], "errores": []}
        
        try:
            # 1. ELIMINAR ARCHIVOS DE FOLIOS VISITA
            folios_visita_dir = self.folios_visita_path
            if os.path.exists(folios_visita_dir):
                folio_file = os.path.join(folios_visita_dir, f"folios_{folio}.json")
                if os.path.exists(folio_file):
                    try:
                        os.remove(folio_file)
                        resultados["eliminados"].append(f"Archivo de folios: folios_{folio}.json")
                    except Exception as e:
                        resultados["errores"].append(f"Error eliminando folios_{folio}.json: {str(e)}")
                
                # Eliminar backup si existe
                backup_dir = os.path.join(folios_visita_dir, "backups")
                if os.path.exists(backup_dir):
                    try:
                        for archivo in os.listdir(backup_dir):
                            if folio in archivo:
                                ruta_archivo = os.path.join(backup_dir, archivo)
                                try:
                                    os.remove(ruta_archivo)
                                    resultados["eliminados"].append(f"Backup: {archivo}")
                                except Exception as e:
                                    resultados["errores"].append(f"Error eliminando backup {archivo}: {str(e)}")
                    except Exception as e:
                        resultados["errores"].append(f"Error accediendo a backups: {str(e)}")
            
            # 2. ELIMINAR ARCHIVOS DE TABLA RELACIÓN BACKUP
            tabla_relacion_backup_dir = os.path.join(
                os.path.dirname(__file__), "data", "tabla_relacion_backups"
            )
            if os.path.exists(tabla_relacion_backup_dir):
                try:
                    for archivo in os.listdir(tabla_relacion_backup_dir):
                        if folio in archivo:
                            ruta_archivo = os.path.join(tabla_relacion_backup_dir, archivo)
                            try:
                                if os.path.isfile(ruta_archivo):
                                    os.remove(ruta_archivo)
                                    resultados["eliminados"].append(f"Tabla backup: {archivo}")
                            except Exception as e:
                                resultados["errores"].append(f"Error eliminando {archivo}: {str(e)}")
                except Exception as e:
                    resultados["errores"].append(f"Error accediendo a tabla_relacion_backups: {str(e)}")
            
        except Exception as e:
            resultados["errores"].append(f"Error general en eliminación: {str(e)}")
        
        return resultados

    def _validar_integridad_historial(self):
        """Valida la integridad del historial y repara si es necesario"""
        try:
            # Verificar que historial_data_original esté sincronizado
            if len(self.historial_data) != len(self.historial_data_original):
                print("⚠️ Resincronizando historial_data_original...")
                self.historial_data_original = self.historial_data.copy()
            
            # Verificar que el archivo JSON existe y es válido
            if os.path.exists(self.historial_path):
                with open(self.historial_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                    json_visitas = json_data.get('visitas', [])
                    
                    # Si hay desincronización, sincronizar
                    if len(json_visitas) != len(self.historial_data):
                        print(f"⚠️ Desincronización detectada. JSON: {len(json_visitas)}, Memoria: {len(self.historial_data)}")
                        self._sincronizar_historial()
            
            return True
        except Exception as e:
            print(f"❌ Error en validación: {e}")
            return False

    def _garantizar_persistencia(self, folio):
        """Garantiza que un folio no exista en ninguna parte del sistema después de eliminación"""
        try:
            # Verificar JSON
            if os.path.exists(self.historial_path):
                with open(self.historial_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                folio_existe = any(v.get('folio_visita') == folio for v in data.get('visitas', []))
                
                if folio_existe:
                    # Eliminar y guardar de nuevo
                    data['visitas'] = [v for v in data.get('visitas', []) if v.get('folio_visita') != folio]
                    with open(self.historial_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    print(f"✅ Folio {folio} eliminado del JSON")
            
            # Verificar carpetas
            carpetas = [
                os.path.join(self.folios_visita_path, f"folios_{folio}.json"),
                os.path.join(os.path.dirname(__file__), "data", "tabla_relacion_backups")
            ]
            
            for carpeta in carpetas:
                if os.path.exists(carpeta):
                    if os.path.isfile(carpeta):
                        os.remove(carpeta)
                        print(f"✅ Archivo eliminado: {carpeta}")
                    elif os.path.isdir(carpeta):
                        for archivo in os.listdir(carpeta):
                            if folio in archivo:
                                ruta = os.path.join(carpeta, archivo)
                                os.remove(ruta)
                                print(f"✅ Archivo de backup eliminado: {archivo}")
            
            return True
        except Exception as e:
            print(f"⚠️ Error en garantía de persistencia: {e}")
            return False

    def _registrar_operacion(self, tipo_operacion, folio, status, detalles=""):
        """Registra todas las operaciones para auditoría y persistencia"""
        try:
            log_path = os.path.join(os.path.dirname(__file__), "data", "operaciones_log.json")
            
            # Cargar log existente o crear uno nuevo
            if os.path.exists(log_path):
                with open(log_path, 'r', encoding='utf-8') as f:
                    log_data = json.load(f)
            else:
                log_data = {"operaciones": []}
            
            # Agregar nueva operación
            operacion = {
                "timestamp": datetime.now().isoformat(),
                "tipo": tipo_operacion,
                "folio": folio,
                "status": status,
                "detalles": detalles
            }
            
            log_data["operaciones"].append(operacion)
            
            # Guardar log
            with open(log_path, 'w', encoding='utf-8') as f:
                json.dump(log_data, f, ensure_ascii=False, indent=2)
            
            return True
        except Exception as e:
            print(f"⚠️ Error registrando operación: {e}")
            return False

    def hist_eliminar_registro(self, registro):
        """Eliminar un registro del historial con persistencia completa"""
        try:
            folio = registro.get('folio_visita', '')
            confirmacion = messagebox.askyesno(
                "Confirmar eliminación", 
                f"¿Está seguro de que desea eliminar el registro del folio {folio}?\n\nSe eliminarán todos los archivos asociados."
            )
            
            if not confirmacion:
                return
            
            # Eliminar archivos asociados de forma segura (folios file + backups)
            resultados = self._eliminar_archivos_asociados_folio(folio)

            # Eliminar SOLO la fila correspondiente en memoria (preferir _id cuando exista)
            registro_id = registro.get('_id')
            if registro_id:
                # Eliminar únicamente el registro con ese _id
                original_len = len(self.historial_data)
                self.historial_data = [r for r in self.historial_data if r.get('_id') != registro_id]
                self.historial_data_original = [r for r in self.historial_data_original if r.get('_id') != registro_id]
                if len(self.historial_data) < original_len:
                    resultados['eliminados'].append(f"Entrada en historial (id={registro_id})")
            else:
                # Fallback por folio si no hay _id
                self.historial_data = [r for r in self.historial_data if r.get('folio_visita') != folio]
                self.historial_data_original = [r for r in self.historial_data_original if r.get('folio_visita') != folio]
                resultados['eliminados'].append("Entrada en historial (por folio)")
            
            # Sincronizar con el archivo JSON (esto actualiza self.historial y guarda)
            sincronizacion_exitosa = self._sincronizar_historial()
            
            if sincronizacion_exitosa:
                resultados["eliminados"].append("✅ Entrada en historial visitas")
            else:
                resultados["errores"].append("⚠️ Error al sincronizar historial")
            
            # Adicional: eliminar entradas de data/tabla_de_relacion.json que pertenezcan a estos folios
            try:
                # Determinar folios numéricos asociados a esta visita (preferir archivo folios_{folio}.json)
                folios_asociados = set()
                folios_visita_file = os.path.join(self.folios_visita_path, f"folios_{folio}.json")
                if os.path.exists(folios_visita_file):
                    try:
                        with open(folios_visita_file, 'r', encoding='utf-8') as f:
                            fv = json.load(f)
                            if isinstance(fv, list):
                                for v in fv:
                                    try:
                                        folios_asociados.add(int(v))
                                    except Exception:
                                        pass
                    except Exception as e:
                        resultados['errores'].append(f"Error leyendo folios de {folios_visita_file}: {e}")
                # Fallback: extraer números del campo 'folios_utilizados' del registro
                if not folios_asociados:
                    posibles = []
                    raw = registro.get('folios_utilizados', '')
                    import re
                    posibles = re.findall(r"\d{1,6}", str(raw))
                    for p in posibles:
                        try:
                            folios_asociados.add(int(p))
                        except Exception:
                            pass

                tabla_relacion_path = os.path.join(os.path.dirname(__file__), 'data', 'tabla_de_relacion.json')
                if os.path.exists(tabla_relacion_path):
                    # Hacer backup antes de modificar
                    try:
                        backup_dir = os.path.join(os.path.dirname(__file__), 'data', 'tabla_relacion_backups')
                        os.makedirs(backup_dir, exist_ok=True)
                        ts = datetime.now().strftime('%Y%m%d%H%M%S')
                        shutil.copyfile(tabla_relacion_path, os.path.join(backup_dir, f"tabla_de_relacion_{folio}_{ts}.json"))
                    except Exception as e:
                        resultados['errores'].append(f"No se pudo crear backup de tabla_de_relacion: {e}")

                    try:
                        with open(tabla_relacion_path, 'r', encoding='utf-8') as f:
                            tabla = json.load(f)

                        # Filtrar filas cuya columna 'FOLIO' coincida con alguno de los folios asociados
                        nueva_tabla = []
                        for row in tabla:
                            try:
                                val = row.get('FOLIO', None)
                                if val is None:
                                    nueva_tabla.append(row)
                                    continue
                                # Normalizar y comparar
                                try:
                                    val_int = int(float(val))
                                except Exception:
                                    val_int = None

                                if val_int is not None and val_int in folios_asociados:
                                    # Saltar -> será eliminado
                                    continue
                                # Si val no es numérico, comparar como string con alguno de los folios formateados
                                if str(val).strip() in {str(f).zfill(6) for f in folios_asociados}:
                                    continue
                                nueva_tabla.append(row)
                            except Exception:
                                nueva_tabla.append(row)

                        # Guardar nueva tabla
                        with open(tabla_relacion_path, 'w', encoding='utf-8') as f:
                            json.dump(nueva_tabla, f, ensure_ascii=False, indent=2)
                        resultados['eliminados'].append('Entradas en tabla_de_relacion.json')
                    except Exception as e:
                        resultados['errores'].append(f"Error modificando tabla_de_relacion.json: {e}")

            except Exception as e:
                resultados['errores'].append(f"Error eliminando entradas de tabla de relación: {e}")

            # Garantizar persistencia completa (verificar que no queden rastros)
            persistencia_garantizada = self._garantizar_persistencia(folio)
            
            if persistencia_garantizada:
                resultados["eliminados"].append("✅ Persistencia verificada y garantizada")
            
            # Registrar la operación para auditoría
            detalles_eliminacion = f"Archivos: {len(resultados['eliminados'])}, Errores: {len(resultados['errores'])}"
            self._registrar_operacion("eliminar_registro", folio, "exitosa", detalles_eliminacion)
            
            # Actualizar la UI
            self._poblar_historial_ui()
            # Recalcular folio actual en memoria y UI para que tome efecto inmediato
            try:
                self.cargar_ultimo_folio()
            except Exception:
                pass
            
            # Mostrar resumen de eliminación
            mensaje = f"✅ Registro del folio {folio} eliminado correctamente\n\n"
            
            if resultados["eliminados"]:
                mensaje += "📁 Elementos eliminados:\n"
                for item in resultados["eliminados"]:
                    mensaje += f"  {item}\n"
            
            if resultados["errores"]:
                mensaje += "\n⚠️ Advertencias:\n"
                for error in resultados["errores"]:
                    mensaje += f"  {error}\n"
            
            messagebox.showinfo("Eliminación completada", mensaje)
            print(f"✅ Folio {folio} eliminado exitosamente con persistencia")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar registro:\n{str(e)}")
            print(f"❌ Error en hist_eliminar_registro: {e}")
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
            
            # Para evitar problemas al actualizar widgets desde hilos en background,
            # realizamos la mutación del historial y la actualización de la UI en el
            # hilo principal usando `self.after(0, ...)`.
            def _apply_and_refresh():
                try:
                    if "visitas" not in self.historial:
                        self.historial["visitas"] = []

                    # Buscar registro existente por folio_visita
                    existing_idx = None
                    try:
                        for idx, v in enumerate(self.historial.get('visitas', [])):
                            if str(v.get('folio_visita', '')).strip().lower() == str(payload.get('folio_visita', '')).strip().lower():
                                existing_idx = idx
                                break
                    except Exception:
                        existing_idx = None

                    # Normalizar campos de dirección en el payload antes de mezclar/añadir
                    for k in ('direccion','calle_numero','colonia','municipio','ciudad_estado','cp'):
                        if k not in payload:
                            payload[k] = ''
                    # Asegurar que cp sea string (preservar ceros a la izquierda si los hay)
                    try:
                        if payload.get('cp') is not None and payload.get('cp') != '':
                            payload['cp'] = str(payload.get('cp'))
                    except Exception:
                        payload['cp'] = str(payload.get('cp') or '')

                    # Si faltan datos de dirección, intentar poblar desde data/Clientes.json
                    try:
                        need_addr = not payload.get('direccion') or not payload.get('calle_numero')
                        cliente_nombre = payload.get('cliente') or ''
                        if need_addr and cliente_nombre:
                            clientes_path = os.path.join(os.path.dirname(__file__), 'data', 'Clientes.json')
                            if os.path.exists(clientes_path):
                                try:
                                    with open(clientes_path, 'r', encoding='utf-8') as cf:
                                        clientes = json.load(cf)
                                    needle = str(cliente_nombre).strip().upper()
                                    for c in (clientes or []):
                                        try:
                                            name = (c.get('CLIENTE') or c.get('RAZÓN SOCIAL ') or c.get('RAZON SOCIAL') or c.get('RAZON_SOCIAL') or '')
                                            if not name:
                                                continue
                                            if str(name).strip().upper() == needle or needle in str(name).strip().upper() or str(name).strip().upper() in needle:
                                                direcciones = c.get('DIRECCIONES') or []
                                                first = None
                                                if isinstance(direcciones, list) and direcciones:
                                                    first = direcciones[0]
                                                if first and isinstance(first, dict):
                                                    payload['calle_numero'] = payload.get('calle_numero') or (first.get('CALLE Y NO') or first.get('CALLE') or '')
                                                    payload['colonia'] = payload.get('colonia') or (first.get('COLONIA O POBLACION') or first.get('COLONIA') or '')
                                                    payload['municipio'] = payload.get('municipio') or (first.get('MUNICIPIO O ALCADIA') or first.get('MUNICIPIO') or '')
                                                    payload['ciudad_estado'] = payload.get('ciudad_estado') or (first.get('CIUDAD O ESTADO') or first.get('CIUDAD') or '')
                                                    cpval = first.get('CP') or first.get('cp')
                                                    if cpval is not None and cpval != '':
                                                        payload['cp'] = str(cpval)
                                                else:
                                                    payload['calle_numero'] = payload.get('calle_numero') or (c.get('CALLE Y NO') or c.get('CALLE') or '')
                                                    payload['colonia'] = payload.get('colonia') or (c.get('COLONIA O POBLACION') or c.get('COLONIA') or '')
                                                    payload['municipio'] = payload.get('municipio') or (c.get('MUNICIPIO O ALCADIA') or c.get('MUNICIPIO') or '')
                                                    payload['ciudad_estado'] = payload.get('ciudad_estado') or (c.get('CIUDAD O ESTADO') or c.get('CIUDAD') or '')
                                                    cpval = c.get('CP') or c.get('cp')
                                                    if cpval is not None and cpval != '':
                                                        payload['cp'] = str(cpval)
                                                break
                                        except Exception:
                                            continue
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    # Construir campo 'direccion' canónico a partir de componentes de dirección
                    try:
                        calle_val = (payload.get('calle_numero') or '').strip()
                        colonia_val = (payload.get('colonia') or '').strip()
                        municipio_val = (payload.get('municipio') or '').strip()
                        ciudad_estado_val = (payload.get('ciudad_estado') or '').strip()
                        cp_val = (payload.get('cp') or '').strip()
                        partes = [p for p in [calle_val, colonia_val, municipio_val, ciudad_estado_val] if p]
                        direccion_comp = ', '.join(partes)
                        if cp_val:
                            direccion_comp = (f"{direccion_comp}, C.P. {cp_val}" if direccion_comp else f"C.P. {cp_val}")
                        if direccion_comp:
                            payload['direccion'] = direccion_comp
                    except Exception:
                        pass

                    if existing_idx is not None:
                        # Mergear campos (no sobrescribir metadatos existentes innecesariamente)
                        existing = self.historial['visitas'][existing_idx]
                        for k, val in payload.items():
                            if k == '_id':
                                continue
                            if val is not None and val != '':
                                existing[k] = val
                        existing.setdefault('estatus', payload.get('estatus', 'En proceso'))
                    else:
                        # Append payload ensuring address fields exist
                        self.historial["visitas"].append(payload)

                    # Actualizar datos en memoria
                    self.historial_data = self.historial.get("visitas", [])

                    # Guardar y refrescar UI
                    self._guardar_historial()
                    self._poblar_historial_ui()

                    # Recalcular folio actual inmediatamente
                    try:
                        self.cargar_ultimo_folio()
                    except Exception:
                        pass

                    # Refrescar dropdown de folios pendientes al añadir una visita
                    try:
                        if hasattr(self, '_refresh_pending_folios_dropdown'):
                            self._refresh_pending_folios_dropdown()
                    except Exception:
                        pass

                    if not es_automatica:
                        messagebox.showinfo("OK", f"Visita {payload.get('folio_visita','-')} guardada correctamente")

                    # DEBUG: mostrar resumen mínimo del historial después de añadir
                    try:
                        print(f"[DEBUG] hist_create_visita: total registros = {len(self.historial.get('visitas', []))}")
                    except Exception:
                        pass
                except Exception as e:
                    print(f"❌ Error aplicando visita en hilo principal: {e}")
                    # Refrescar dropdown de folios pendientes cuando se aplica una modificación
                    try:
                        if hasattr(self, '_refresh_pending_folios_dropdown'):
                            self._refresh_pending_folios_dropdown()
                    except Exception:
                        pass

            # Programar la aplicación en el hilo principal
            try:
                self.after(0, _apply_and_refresh)
            except Exception:
                # Como fallback, intentar aplicar inmediatamente
                _apply_and_refresh()
                
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def hist_buscar_por_folio(self):
        """Buscar en el historial por folio de visita"""
        try:
            # resetear paginado al buscar
            self.HISTORIAL_PAGINA_ACTUAL = 1
            folio_busqueda_raw = self.entry_buscar_folio.get().strip()
            folio_busqueda = folio_busqueda_raw.lower()

            if not folio_busqueda_raw:
                # Si no hay búsqueda, mostrar todos los datos
                self.historial_data = self.historial_data_original.copy() if hasattr(self, 'historial_data_original') else self.historial_data
            else:
                # Filtrar datos por folio con normalización de dígitos
                resultados = []
                fuente = (self.historial_data_original if hasattr(self, 'historial_data_original') else self.historial_data)
                digits_search = ''.join([c for c in folio_busqueda_raw if c.isdigit()])
                for registro in fuente:
                    folio_actual = str(registro.get('folio_visita', '') or '')
                    # coincidencia directa (substring)
                    if folio_busqueda in folio_actual.lower():
                        resultados.append(registro)
                        continue
                    # coincidencia por dígitos (ignora padding)
                    if digits_search:
                        folio_digits = ''.join([c for c in folio_actual if c.isdigit()])
                        if folio_digits and digits_search in folio_digits:
                            resultados.append(registro)
                            continue

                self.historial_data = resultados
            
            self._poblar_historial_ui()
            
        except Exception as e:
            print(f"Error en búsqueda por folio: {e}")

    def hist_borrar_por_folio(self):
        """Borrar una visita usando el folio ingresado en la barra de búsqueda.
        Busca el registro por folio (coincidencia exacta primero, luego parcial)
        y delega en `hist_eliminar_registro` para la eliminación con confirmación.
        """
        try:
            folio = self.entry_buscar_folio.get().strip()
            if not folio:
                messagebox.showwarning("Advertencia", "Ingrese el folio en la barra de búsqueda para eliminar una visita.")
                return

            fuente = self.historial_data_original if hasattr(self, 'historial_data_original') else self.historial_data

            # Buscar coincidencia exacta (case-insensitive)
            matches = [r for r in fuente if str(r.get('folio_visita', '')).strip().lower() == folio.lower()]

            # Si no hay exactas, buscar por contains
            if not matches:
                matches = [r for r in fuente if folio.lower() in str(r.get('folio_visita', '')).lower()]

            if not matches:
                messagebox.showinfo("No encontrado", f"No se encontró ningún registro con folio '{folio}'.")
                return

            if len(matches) > 1:
                # Informar que se encontró más de una coincidencia y proceder con la primera
                confirmar = messagebox.askyesno(
                    "Confirmar eliminación",
                    f"Se encontraron {len(matches)} registros que coinciden con '{folio}'.\n\nSe eliminará el primer registro encontrado: {matches[0].get('folio_visita')}\n\n¿Desea continuar?"
                )
                if not confirmar:
                    return

            # Delegar en la función existente para eliminar (esta función pedirá su propia confirmación también)
            # Llamamos a hist_eliminar_registro con el primer match
            self.hist_eliminar_registro(matches[0])

        except Exception as e:
            messagebox.showerror("Error", f"Error al intentar eliminar por folio:\n{e}")

    def hist_update_visita(self, id_, nuevos):
        """Actualiza una visita existente"""
        try:
            # Buscar la visita a actualizar y mezclar (merge) los campos nuevos
            visitas = self.historial.get("visitas", [])
            encontrado = False
            for i, v in enumerate(visitas):
                try:
                    if v.get("_id") == id_ or v.get("id") == id_:
                        encontrado = True
                    else:
                        # permitir búsquedas por folio_visita o folio_acta si se pasó un folio
                        if id_ and isinstance(id_, str):
                            if id_.strip() and (id_.strip() == (v.get('folio_visita','') or '').strip() or id_.strip() == (v.get('folio_acta','') or '').strip()):
                                encontrado = True
                except Exception:
                    continue

                if encontrado:
                    actualizado = v.copy()
                    # Mezclar claves de 'nuevos' sobre el registro existente
                    for k, val in (nuevos or {}).items():
                        if k == "_id":
                            continue
                        actualizado[k] = val

                    # Normalizar y asegurar campos de dirección persisten
                    for k in ('direccion','calle_numero','colonia','municipio','ciudad_estado','cp'):
                        if k not in actualizado:
                            actualizado[k] = ''
                    # Si 'calle_numero' no existe pero 'direccion' sí, sincronizar
                    if not actualizado.get('calle_numero') and actualizado.get('direccion'):
                        actualizado['calle_numero'] = actualizado.get('direccion')
                    # Forzar cp como string
                    try:
                        if actualizado.get('cp') is not None and actualizado.get('cp') != '':
                            actualizado['cp'] = str(actualizado.get('cp'))
                    except Exception:
                        actualizado['cp'] = str(actualizado.get('cp') or '')

                    # Construir campo 'direccion' canónico a partir de componentes de dirección
                    try:
                        calle_val = (actualizado.get('calle_numero') or '').strip()
                        colonia_val = (actualizado.get('colonia') or '').strip()
                        municipio_val = (actualizado.get('municipio') or '').strip()
                        ciudad_estado_val = (actualizado.get('ciudad_estado') or '').strip()
                        cp_val = (actualizado.get('cp') or '').strip()
                        partes = [p for p in [calle_val, colonia_val, municipio_val, ciudad_estado_val] if p]
                        direccion_comp = ', '.join(partes)
                        if cp_val:
                            direccion_comp = (f"{direccion_comp}, C.P. {cp_val}" if direccion_comp else f"C.P. {cp_val}")
                        if direccion_comp:
                            actualizado['direccion'] = direccion_comp
                    except Exception:
                        pass
                    # Si faltan campos de dirección, intentar poblar desde data/Clientes.json
                    try:
                        need_addr = not actualizado.get('direccion') or not actualizado.get('calle_numero')
                        cliente_nombre = actualizado.get('cliente') or ''
                        if need_addr and cliente_nombre:
                            clientes_path = os.path.join(os.path.dirname(__file__), 'data', 'Clientes.json')
                            if os.path.exists(clientes_path):
                                try:
                                    with open(clientes_path, 'r', encoding='utf-8') as cf:
                                        clientes = json.load(cf)
                                    needle = str(cliente_nombre).strip().upper()
                                    for c in (clientes or []):
                                        try:
                                            name = (c.get('CLIENTE') or c.get('RAZÓN SOCIAL ') or c.get('RAZON SOCIAL') or c.get('RAZON_SOCIAL') or '')
                                            if not name:
                                                continue
                                            if str(name).strip().upper() == needle or needle in str(name).strip().upper() or str(name).strip().upper() in needle:
                                                # try DIRECCIONES first
                                                direcciones = c.get('DIRECCIONES') or []
                                                first = None
                                                if isinstance(direcciones, list) and direcciones:
                                                    first = direcciones[0]
                                                if first and isinstance(first, dict):
                                                    actualizado['calle_numero'] = actualizado.get('calle_numero') or (first.get('CALLE Y NO') or first.get('CALLE') or '')
                                                    actualizado['colonia'] = actualizado.get('colonia') or (first.get('COLONIA O POBLACION') or first.get('COLONIA') or '')
                                                    actualizado['municipio'] = actualizado.get('municipio') or (first.get('MUNICIPIO O ALCADIA') or first.get('MUNICIPIO') or '')
                                                    actualizado['ciudad_estado'] = actualizado.get('ciudad_estado') or (first.get('CIUDAD O ESTADO') or first.get('CIUDAD') or '')
                                                    cpval = first.get('CP') or first.get('cp')
                                                    if cpval is not None and cpval != '':
                                                        actualizado['cp'] = str(cpval)
                                                else:
                                                    # try top-level keys
                                                    actualizado['calle_numero'] = actualizado.get('calle_numero') or (c.get('CALLE Y NO') or c.get('CALLE') or '')
                                                    actualizado['colonia'] = actualizado.get('colonia') or (c.get('COLONIA O POBLACION') or c.get('COLONIA') or '')
                                                    actualizado['municipio'] = actualizado.get('municipio') or (c.get('MUNICIPIO O ALCADIA') or c.get('MUNICIPIO') or '')
                                                    actualizado['ciudad_estado'] = actualizado.get('ciudad_estado') or (c.get('CIUDAD O ESTADO') or c.get('CIUDAD') or '')
                                                    cpval = c.get('CP') or c.get('cp')
                                                    if cpval is not None and cpval != '':
                                                        actualizado['cp'] = str(cpval)
                                                break
                                        except Exception:
                                            continue
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    # Reemplazar el registro en la lista
                    try:
                        self.historial['visitas'][i] = actualizado
                    except Exception:
                        pass

                    # Actualizar vistas en memoria y persistir
                    self.historial_data = self.historial.get("visitas", [])
                    self._guardar_historial()
                    self._poblar_historial_ui()
                    try:
                        self.cargar_ultimo_folio()
                    except Exception:
                        pass
                    # Debug: confirmar en consola los campos de dirección guardados
                    try:
                        print(f"[DEBUG] visita actualizada _id={actualizado.get('_id')} folio={actualizado.get('folio_visita')} direccion={actualizado.get('direccion')} calle_numero={actualizado.get('calle_numero')} colonia={actualizado.get('colonia')} municipio={actualizado.get('municipio')} cp={actualizado.get('cp')}")
                    except Exception:
                        pass
                    messagebox.showinfo("OK", f"Visita {actualizado.get('folio_visita','-')} actualizada")
                    # Refrescar dropdown de pendientes por si cambió estatus/tipo
                    try:
                        if hasattr(self, '_refresh_pending_folios_dropdown'):
                            self._refresh_pending_folios_dropdown()
                    except Exception:
                        pass

                    # También actualizar en memoria/pending_folios si existe
                    try:
                        if hasattr(self, 'pending_folios') and isinstance(self.pending_folios, list):
                            for j, p in enumerate(self.pending_folios):
                                try:
                                    pid = p.get('_id') or p.get('id')
                                    if pid == id_ or p.get('folio_visita') == id_ or p.get('folio_acta') == id_:
                                        self.pending_folios[j].update(nuevos or {})
                                except Exception:
                                    continue
                            # persistir cambios
                            try:
                                self._save_pending_folios()
                            except Exception:
                                pass
                    except Exception:
                        pass

                    return

            # Si no encontramos coincidencias, mostrar advertencia (no lanzar excepción)
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
            # Leer supervisor de forma segura (proviene de la tabla de relación; el campo UI fue removido)
            safe_supervisor_widget = getattr(self, 'entry_supervisor', None)
            try:
                supervisor = safe_supervisor_widget.get().strip() if safe_supervisor_widget and safe_supervisor_widget.winfo_exists() else ""
            except Exception:
                supervisor = ""

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

            # CARGAR DATOS DE TABLA DE RELACIÓN SI EXISTEN
            datos_tabla = []
            if self.archivo_json_generado and os.path.exists(self.archivo_json_generado):
                with open(self.archivo_json_generado, 'r', encoding='utf-8') as f:
                    datos_tabla = json.load(f)
                    
                # Guardar folios específicos para esta visita
                self.guardar_folios_visita(folio_visita, datos_tabla)

            # ===== EXTRACCIÓN DE FOLIOS NUMÉRICOS DE LA TABLA DE RELACIÓN =====
            folios_numericos = []
            folios_totales = 0
            
            if datos_tabla:
                for registro in datos_tabla:
                    # Contar registros totales
                    folios_totales += 1
                    
                    # Extraer folio numérico si existe y no es NaN
                    if "FOLIO" in registro:
                        folio_valor = registro["FOLIO"]
                        
                        # Verificar que no sea NaN, None o vacío
                        if (folio_valor is not None and 
                            str(folio_valor).strip() != "" and 
                            str(folio_valor).lower() != "nan" and
                            str(folio_valor).lower() != "none"):
                            
                            try:
                                # Convertir a entero
                                folio_int = int(float(folio_valor))
                                folios_numericos.append(folio_int)
                            except (ValueError, TypeError):
                                # Si no se puede convertir a número, ignorar
                                pass
            
            # Ordenar folios numéricos
            folios_numericos_ordenados = sorted(folios_numericos)
            
            # Formatear información de folios para mostrar
            if folios_numericos_ordenados:
                if len(folios_numericos_ordenados) == 1:
                    folios_str = f"Folio: {folios_numericos_ordenados[0]:06d}"
                else:
                    # Mostrar solo rango
                    folios_str = f"{folios_numericos_ordenados[0]:06d} - {folios_numericos_ordenados[-1]:06d}"
            else:
                # Si no hay folios numéricos
                if folios_totales > 0:
                    folios_str = f"Total: {folios_totales} folios"
                else:
                    folios_str = "No se encontraron folios"

            # ===== EXTRACCIÓN DE NORMAS DE LA TABLA DE RELACIÓN =====
            normas_encontradas = set()  # Usamos set para evitar duplicados
            
            if datos_tabla:
                # Cargar el archivo de normas
                normas_path = os.path.join(os.path.dirname(__file__), "data", "Normas.json")
                
                if os.path.exists(normas_path):
                    with open(normas_path, 'r', encoding='utf-8') as f:
                        normas_data = json.load(f)
                    
                    # Crear un diccionario para mapear números de norma a códigos NOM completos
                    normas_mapeadas = {}
                    for norma_obj in normas_data:
                        if isinstance(norma_obj, dict) and "NOM" in norma_obj:
                            nom_code = norma_obj["NOM"]
                            # Extraer el número del código NOM
                            try:
                                import re
                                match = re.search(r'NOM-(\d+)-', nom_code)
                                if match:
                                    num_norma = int(match.group(1))
                                    normas_mapeadas[num_norma] = nom_code
                            except (ValueError, AttributeError):
                                pass
                
                # Buscar normas UVA en la tabla de relación
                for registro in datos_tabla:
                    if "NORMA UVA" in registro:
                        norma_uva = registro["NORMA UVA"]
                        # Verificar que no sea NaN o vacío
                        if norma_uva is not None and str(norma_uva).strip() != "" and str(norma_uva).lower() != "nan":
                            try:
                                # Convertir a entero (puede venir como string "4" o float 4.0)
                                norma_num = int(float(norma_uva))
                                
                                # Buscar la NOM correspondiente en el mapeo
                                if norma_num in normas_mapeadas:
                                    normas_encontradas.add(normas_mapeadas[norma_num])
                                else:
                                    # Si no encontramos mapeo, agregar como "NORMA UVA X"
                                    normas_encontradas.add(f"NORMA UVA {norma_num}")
                            except (ValueError, TypeError):
                                # Si no se puede convertir a número, agregar el valor tal cual
                                if str(norma_uva).strip():
                                    normas_encontradas.add(str(norma_uva).strip())
            
            # Crear cadena de normas (ordenar alfabéticamente para consistencia)
            normas_str = ", ".join(sorted(normas_encontradas)) if normas_encontradas else ""

            # ===== EXTRACCIÓN DE FIRMAS (SUPERVISORES) DE LA TABLA DE RELACIÓN =====
            supervisores_encontrados = set()  # Usamos set para evitar duplicados
            firmas_originales = set()  # Para guardar las firmas originales también
            
            if datos_tabla:
                # Cargar el archivo de firmas
                firmas_path = os.path.join(os.path.dirname(__file__), "data", "Firmas.json")
                
                # Prepare mapping dict even if file missing
                firmas_mapeadas = {}
                if os.path.exists(firmas_path):
                    with open(firmas_path, 'r', encoding='utf-8') as f:
                        firmas_data = json.load(f)

                    # Crear un diccionario para mapear firmas (normalizadas) a nombres completos
                    for inspector_obj in firmas_data:
                        if isinstance(inspector_obj, dict) and "FIRMA" in inspector_obj and "NOMBRE DE INSPECTOR" in inspector_obj:
                            raw_firma = inspector_obj["FIRMA"]
                            nombre_completo = inspector_obj["NOMBRE DE INSPECTOR"]
                            if raw_firma is None:
                                continue
                            key = str(raw_firma).strip().upper()
                            firmas_mapeadas[key] = nombre_completo
                
                # Buscar firmas en la tabla de relación
                for registro in datos_tabla:
                    if "FIRMA" in registro:
                        firma = registro["FIRMA"]
                        # Verificar que no sea NaN o vacío
                        if firma is not None and str(firma).strip() != "" and str(firma).lower() != "nan":
                            firma_str = str(firma).strip()
                            firmas_originales.add(firma_str)

                            # Normalizar para búsqueda
                            buscar_clave = firma_str.upper()
                            if buscar_clave in firmas_mapeadas:
                                supervisores_encontrados.add(firmas_mapeadas[buscar_clave])
                            else:
                                # Si no encontramos mapeo, agregar la firma original
                                supervisores_encontrados.add(firma_str)
            
            # Crear cadena de supervisores (ordenar alfabéticamente)
            supervisores_str = ", ".join(sorted(supervisores_encontrados)) if supervisores_encontrados else ""
            
            # Determinar qué supervisor mostrar en el campo principal
            # Prioridad: 1. Supervisores de la tabla, 2. Supervisor del formulario
            supervisor_mostrar = supervisores_str if supervisores_str else supervisor

            # Determinar tipo de documento para visitas automáticas (usar selección si existe)
            tipo_documento = (self.combo_tipo_documento.get().strip()
                               if hasattr(self, 'combo_tipo_documento') else "Dictamen")

            # Crear payload para visita automática con información de folios
            payload = {
                "folio_visita": folio_visita,
                "folio_acta": folio_acta or f"AC{self.current_folio}",
                "fecha_inicio": fecha_inicio or datetime.now().strftime("%d/%m/%Y"),
                "fecha_termino": fecha_termino,
                "hora_inicio_24h": hora_inicio_24h or datetime.now().strftime("%H:%M"),
                "hora_termino_24h": hora_termino_24h or datetime.now().strftime("%H:%M"),
                "hora_inicio": hora_inicio_formateada or self._formatear_hora_12h(datetime.now().strftime("%H:%M")),
                "hora_termino": hora_termino_formateada,
                "norma": normas_str,  # Normas encontradas
                "cliente": self.cliente_seleccionado['CLIENTE'],
                "nfirma1": supervisor_mostrar or " ",  # Supervisor principal (prioridad a los de la tabla)
                "nfirma2": "",
                "estatus": "Completada",
                "tipo_documento": tipo_documento,
                "folios_utilizados": folios_str,  # Información formateada de folios
                "total_folios": folios_totales,
                "total_folios_numericos": len(folios_numericos),
                "supervisores_tabla": supervisores_str,  # Todos los supervisores de la tabla
                "supervisor_formulario": supervisor  # Supervisor del formulario (por si se necesita)
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
        """Crea un formulario modal para editar visitas con disposición organizada
        Añade un combobox de domicilios dependiente del cliente para permitir
        seleccionar la dirección registrada y guardarla en la visita.
        """
        datos = datos or {}
        modal = ctk.CTkToplevel(self)
        modal.title("Editar Visita")
        modal.geometry("1200x600")  # Aumentado altura para mejor visibilidad
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
        
        # Configurar 4 columnas (última para inspectores)
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_columnconfigure(2, weight=1)
        content_frame.grid_columnconfigure(3, weight=3)
        
        entries = {}
        # Variable closure para almacenar la dirección raw seleccionada
        selected_address_raw = {}
        # Cargar listas de normas e inspectores para helpers del modal
        try:
            normas_path = os.path.join(os.path.dirname(__file__), 'data', 'Normas.json')
            if os.path.exists(normas_path):
                with open(normas_path, 'r', encoding='utf-8') as nf:
                    normas_data = json.load(nf)
                    normas_list = [n.get('NOM') or n.get('NOMBRE') or str(n) for n in (normas_data or [])]
            else:
                normas_list = []
        except Exception:
            normas_list = []

        try:
            firmas_path = os.path.join(os.path.dirname(__file__), 'data', 'Firmas.json')
            if os.path.exists(firmas_path):
                with open(firmas_path, 'r', encoding='utf-8') as ff:
                    firmas_data = json.load(ff)
                    inspectores_list = [f.get('NOMBRE DE INSPECTOR') or f.get('NOMBRE') or '' for f in (firmas_data or [])]
                    # Mapa rápido nombre -> normas acreditadas
                    try:
                        firmas_map = {}
                        for f in (firmas_data or []):
                            name = f.get('NOMBRE DE INSPECTOR') or f.get('NOMBRE') or ''
                            normas_ac = f.get('Normas acreditadas') or f.get('Normas Acreditadas') or f.get('Normas') or []
                            firmas_map[name] = normas_ac or []
                    except Exception:
                        firmas_map = {}
            else:
                inspectores_list = []
        except Exception:
            inspectores_list = []
            firmas_map = {}
        
        # Definir campos organizados por columnas
        campos_por_columna = [
            [  # Columna 0: Información básica
                ("fecha_inicio", "Fecha Inicio"),
                ("fecha_termino", "Fecha Termino"),
                ("tipo_documento", "Tipo de documento"),
                ("folio_visita", "Folio Visita"),
                ("folio_acta", "Folio Acta"),
                ("folios_utilizados", "Folios Utilizados"),
            ],
            [  # Columna 1: normas (normas se mostrará aquí)
                
                ("norma", ""),
            ],
            [  # Columna 2: Cliente y estatus
                ("cliente", "Cliente"),
                ("direccion", "Domicilio registrado"),
                ("estatus", "Estatus"),
            ],
            [  # Columna 3: Inspectores (UI)
                # esta columna se llenará con la lista de inspectores (checkboxes)
            ]
        ]
        
        # Crear campos para cada columna
        col_frames = []
        for col_idx, campos in enumerate(campos_por_columna):
            col_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            col_frame.grid(row=0, column=col_idx, padx=10, pady=0, sticky="nsew")
            content_frame.grid_columnconfigure(col_idx, weight=1)
            col_frames.append(col_frame)
            
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
                
                if key == "tipo_documento":
                    # ComboBox para tipo de documento
                    opciones_tipo = ["Dictamen", "Negación de dictamen", "Constancia", "Negación de constancia"]
                    ent = ctk.CTkComboBox(field_frame, values=opciones_tipo, font=FONT_SMALL, state="readonly", height=35, corner_radius=8)
                    ent.pack(fill="x")
                    ent.set(datos.get("tipo_documento", "Dictamen"))
                    entries[key] = ent
                    continue
                if key in ("fecha_inicio", "fecha_termino"):
                    # intentar usar DateEntry de tkcalendar si está disponible
                    DateEntry = None
                    try:
                        from tkcalendar import DateEntry as _DateEntry
                        DateEntry = _DateEntry
                    except Exception:
                        DateEntry = None

                    if DateEntry is not None:
                        try:
                            # crear un estilo ttk para que el DateEntry visualmente encaje con CTk
                            try:
                                style = ttk.Style()
                                style_name = f"CTkDate.{key}.TEntry"
                                style.configure(style_name, fieldbackground=STYLE.get('surface'), background=STYLE.get('surface'), foreground=STYLE.get('texto_oscuro'))
                            except Exception:
                                style_name = None
                            kwargs = {'date_pattern': 'dd/MM/yyyy', 'width': 16}
                            if style_name:
                                kwargs['style'] = style_name
                            ent = DateEntry(field_frame, **kwargs)
                            ent.pack(fill='x')
                            if datos and key in datos and datos.get(key):
                                try:
                                    ent.set_date(datos.get(key))
                                except Exception:
                                    try:
                                        ent.set_date(datetime.strptime(datos.get(key), '%d/%m/%Y'))
                                    except Exception:
                                        pass
                        except Exception:
                            ent = ctk.CTkEntry(field_frame, height=35, corner_radius=8, font=FONT_SMALL)
                            ent.pack(fill='x')
                            if datos and key in datos:
                                ent.insert(0, str(datos.get(key, '')))
                    else:
                        ent = ctk.CTkEntry(field_frame, height=35, corner_radius=8, font=FONT_SMALL)
                        ent.pack(fill='x')
                        if datos and key in datos:
                            ent.insert(0, str(datos.get(key, '')))
                    entries[key] = ent
                    continue
                if key == "cliente":
                    # Obtener lista de clientes
                    clientes_lista = ["Seleccione un cliente..."]
                    if hasattr(self, 'clientes_data') and self.clientes_data:
                        for cliente in self.clientes_data:
                            if not isinstance(cliente, dict):
                                continue
                            name = cliente.get('CLIENTE') or cliente.get('RAZÓN SOCIAL ') or cliente.get('RAZON SOCIAL') or cliente.get('RAZON_SOCIAL') or cliente.get('RFC') or cliente.get('NÚMERO_DE_CONTRATO')
                            if name:
                                clientes_lista.append(name)
                    
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
                    # Callback cuando se seleccione un cliente en el modal
                    def _on_cliente_modal_select(val):
                        nombre = val
                        # buscar dict del cliente
                        encontrado = None
                        try:
                            for c in (self.clientes_data or []):
                                if not isinstance(c, dict):
                                    continue
                                name = c.get('CLIENTE') or c.get('RAZÓN SOCIAL ') or c.get('RAZON SOCIAL') or c.get('RAZON_SOCIAL') or c.get('RFC') or c.get('NÚMERO_DE_CONTRATO')
                                if name and name == nombre:
                                    encontrado = c
                                    break
                        except Exception:
                            encontrado = None

                        # Construir lista de domicilios (mismo heurístico usado en actualizar_cliente_seleccionado)
                        domicilios = []
                        raw = []
                        if encontrado:
                            try:
                                direcciones = encontrado.get('DIRECCIONES')
                                if isinstance(direcciones, list) and direcciones:
                                    for d in direcciones:
                                        if not isinstance(d, dict):
                                            continue
                                        parts = []
                                        for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO', 'CALLE_Y_NRO', 'NUMERO'):
                                            v = d.get(k) or d.get(k.upper()) if isinstance(d, dict) else None
                                            if v:
                                                parts.append(str(v))
                                        for k in ('COLONIA O POBLACION', 'COLONIA'):
                                            v = d.get(k)
                                            if v:
                                                parts.append(str(v))
                                        for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                                            v = d.get(k)
                                            if v:
                                                parts.append(str(v))
                                        if d.get('CIUDAD O ESTADO'):
                                            parts.append(str(d.get('CIUDAD O ESTADO')))
                                        if d.get('CP'):
                                            parts.append(str(d.get('CP')))
                                        addr = ", ".join(parts).strip()
                                        if addr:
                                            domicilios.append(addr)
                                            raw.append(d)

                                # fallback: intentar con campos a nivel superior
                                if not domicilios:
                                    parts = []
                                    for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO'):
                                        v = encontrado.get(k) or encontrado.get(k.upper())
                                        if v:
                                            parts.append(str(v))
                                    for k in ('COLONIA O POBLACION', 'COLONIA'):
                                        v = encontrado.get(k)
                                        if v:
                                            parts.append(str(v))
                                    for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                                        v = encontrado.get(k)
                                        if v:
                                            parts.append(str(v))
                                    if encontrado.get('CIUDAD O ESTADO'):
                                        parts.append(str(encontrado.get('CIUDAD O ESTADO')))
                                    if encontrado.get('CP') is not None:
                                        parts.append(str(encontrado.get('CP')))
                                    addr = ", ".join(parts).strip()
                                    if addr:
                                        domicilios.append(addr)
                                        raw.append({
                                            'CALLE Y NO': encontrado.get('CALLE Y NO') or encontrado.get('CALLE') or encontrado.get('CALLE_Y_NO') or '',
                                            'COLONIA O POBLACION': encontrado.get('COLONIA O POBLACION') or encontrado.get('COLONIA') or '',
                                            'MUNICIPIO O ALCADIA': encontrado.get('MUNICIPIO O ALCADIA') or encontrado.get('MUNICIPIO') or '',
                                            'CIUDAD O ESTADO': encontrado.get('CIUDAD O ESTADO') or encontrado.get('CIUDAD') or '',
                                            'CP': encontrado.get('CP')
                                        })
                            except Exception:
                                domicilios = []

                        if not domicilios:
                            domicilios = ["Domicilio no disponible"]
                            raw = [{'CALLE Y NO': '', 'COLONIA O POBLACION': '', 'MUNICIPIO O ALCADIA': '', 'CIUDAD O ESTADO': '', 'CP': ''}]

                        # configurar combobox de domicilios del modal
                        try:
                            vals = ['Seleccione un domicilio...'] + domicilios
                            if 'direccion' in entries and isinstance(entries['direccion'], ctk.CTkComboBox):
                                entries['direccion'].configure(values=vals, state='readonly')
                                entries['direccion'].set('Seleccione un domicilio...')
                                # almacenar raw list alineada y display list para referencias
                                entries['_domicilios_modal_raw'] = raw
                                entries['_domicilios_modal_display'] = domicilios

                                # handler cuando se selecciona un domicilio en el combobox
                                def _on_domicilio_select(val):
                                    try:
                                        display = entries.get('_domicilios_modal_display', []) or []
                                        rawlist = entries.get('_domicilios_modal_raw', []) or []
                                        if not display or not rawlist:
                                            selected_address_raw.clear()
                                            return
                                        if val == 'Seleccione un domicilio...':
                                            selected_address_raw.clear()
                                            return
                                        if val in display:
                                            idx = display.index(val)
                                            if idx < len(rawlist):
                                                selected_address_raw.clear()
                                                selected_address_raw.update(rawlist[idx])
                                    except Exception:
                                        pass

                                entries['direccion'].configure(command=_on_domicilio_select)

                                # si la visita ya tiene una dirección guardada, intentar seleccionarla
                                # considerar tanto 'direccion' como 'calle_numero' como posibles fuentes
                                saved_vals = []
                                if datos:
                                    if datos.get('direccion'):
                                        saved_vals.append(str(datos.get('direccion')))
                                    if datos.get('calle_numero'):
                                        saved_vals.append(str(datos.get('calle_numero')))

                                matched = False
                                for saved in saved_vals:
                                    if not saved:
                                        continue
                                    if saved in domicilios:
                                        try:
                                            entries['direccion'].set(saved)
                                            idx = domicilios.index(saved)
                                            selected_address_raw.clear()
                                            selected_address_raw.update(raw[idx])
                                            matched = True
                                            break
                                        except Exception:
                                            pass

                                # si no hubo match exacto, intentar empatar por componentes
                                if not matched and datos:
                                    parts = []
                                    # preferir campos disponibles en datos (soporta varias claves)
                                    for k in ('direccion', 'calle_numero', 'CALLE Y NO', 'CALLE'):
                                        v = datos.get(k)
                                        if v:
                                            parts.append(str(v))
                                            break
                                    if datos.get('colonia'):
                                        parts.append(str(datos.get('colonia')))
                                    if datos.get('municipio'):
                                        parts.append(str(datos.get('municipio')))
                                    if datos.get('ciudad_estado'):
                                        parts.append(str(datos.get('ciudad_estado')))
                                    if datos.get('cp'):
                                        parts.append(str(datos.get('cp')))
                                    built = ", ".join(parts).strip()
                                    if built and built in domicilios:
                                        try:
                                            entries['direccion'].set(built)
                                            idx = domicilios.index(built)
                                            selected_address_raw.clear()
                                            selected_address_raw.update(raw[idx])
                                        except Exception:
                                            pass
                                # Si sigue sin empatar exactamente, intentar empatar por fragmento de 'calle y no' o 'calle_numero'
                                if not matched and datos:
                                    fragment = None
                                    for k in ('calle_numero', 'CALLE Y NO', 'CALLE', 'direccion'):
                                        v = datos.get(k)
                                        if v:
                                            fragment = str(v).strip()
                                            break
                                    if fragment:
                                        for i, disp in enumerate(domicilios):
                                            try:
                                                if fragment and fragment.lower() in disp.lower():
                                                    entries['direccion'].set(disp)
                                                    selected_address_raw.clear()
                                                    if i < len(raw):
                                                        selected_address_raw.update(raw[i])
                                                    break
                                            except Exception:
                                                continue
                        except Exception:
                            pass

                    # enlazar callback
                    ent.configure(command=_on_cliente_modal_select)

                    # Establecer cliente si existe en datos
                    if datos and "cliente" in datos:
                        cliente_actual = datos.get("cliente", "")
                        if cliente_actual in clientes_lista:
                            ent.set(cliente_actual)
                            # forzar población de domicilios al abrir modal
                            try:
                                _on_cliente_modal_select(cliente_actual)
                            except Exception:
                                pass
                        else:
                            # intentar encontrar coincidencia en self.clientes_data por nombre
                            encontrado_cliente = None
                            try:
                                needle = cliente_actual.strip().lower()
                                for c in (self.clientes_data or []):
                                    try:
                                        name = c.get('CLIENTE') or c.get('RAZÓN SOCIAL ') or c.get('RAZON SOCIAL') or c.get('RAZON_SOCIAL') or c.get('RFC') or c.get('NÚMERO_DE_CONTRATO')
                                        if not name:
                                            continue
                                        name_s = str(name).strip().lower()
                                        if name_s == needle or needle in name_s or name_s in needle:
                                            encontrado_cliente = c
                                            break
                                    except Exception:
                                        continue
                            except Exception:
                                encontrado_cliente = None

                            if encontrado_cliente:
                                # poblar domicilios directamente desde el dict encontrado
                                try:
                                    parts_list = []
                                    domicilios = []
                                    raw = []
                                    direcciones = encontrado_cliente.get('DIRECCIONES')
                                    if isinstance(direcciones, list) and direcciones:
                                        for d in direcciones:
                                            if not isinstance(d, dict):
                                                continue
                                            parts = []
                                            for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO', 'CALLE_Y_NRO', 'NUMERO'):
                                                v = d.get(k) or d.get(k.upper()) if isinstance(d, dict) else None
                                                if v:
                                                    parts.append(str(v))
                                            for k in ('COLONIA O POBLACION', 'COLONIA'):
                                                v = d.get(k)
                                                if v:
                                                    parts.append(str(v))
                                            for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                                                v = d.get(k)
                                                if v:
                                                    parts.append(str(v))
                                            if d.get('CIUDAD O ESTADO'):
                                                parts.append(str(d.get('CIUDAD O ESTADO')))
                                            if d.get('CP') is not None:
                                                parts.append(str(d.get('CP')))
                                            addr = ", ".join(parts).strip()
                                            if addr:
                                                domicilios.append(addr)
                                                raw.append(d)

                                    if not domicilios:
                                        # fallback a nivel cliente
                                        parts = []
                                        for k in ('CALLE Y NO', 'CALLE', 'CALLE_Y_NO'):
                                            v = encontrado_cliente.get(k) or encontrado_cliente.get(k.upper())
                                            if v:
                                                parts.append(str(v))
                                        for k in ('COLONIA O POBLACION', 'COLONIA'):
                                            v = encontrado_cliente.get(k)
                                            if v:
                                                parts.append(str(v))
                                        for k in ('MUNICIPIO O ALCADIA', 'MUNICIPIO'):
                                            v = encontrado_cliente.get(k)
                                            if v:
                                                parts.append(str(v))
                                        if encontrado_cliente.get('CIUDAD O ESTADO'):
                                            parts.append(str(encontrado_cliente.get('CIUDAD O ESTADO')))
                                        if encontrado_cliente.get('CP') is not None:
                                            parts.append(str(encontrado_cliente.get('CP')))
                                        addr = ", ".join(parts).strip()
                                        if addr:
                                            domicilios.append(addr)
                                            raw.append({
                                                'CALLE Y NO': encontrado_cliente.get('CALLE Y NO') or encontrado_cliente.get('CALLE') or encontrado_cliente.get('CALLE_Y_NO') or '',
                                                'COLONIA O POBLACION': encontrado_cliente.get('COLONIA O POBLACION') or encontrado_cliente.get('COLONIA') or '',
                                                'MUNICIPIO O ALCADIA': encontrado_cliente.get('MUNICIPIO O ALCADIA') or encontrado_cliente.get('MUNICIPIO') or '',
                                                'CIUDAD O ESTADO': encontrado_cliente.get('CIUDAD O ESTADO') or encontrado_cliente.get('CIUDAD') or '',
                                                'CP': encontrado_cliente.get('CP')
                                            })

                                    vals = ['Seleccione un domicilio...'] + (domicilios or ['Domicilio no disponible'])
                                    if 'direccion' in entries and isinstance(entries['direccion'], ctk.CTkComboBox):
                                        entries['direccion'].configure(values=vals, state='readonly')
                                        entries['direccion'].set('Seleccione un domicilio...')
                                        entries['_domicilios_modal_raw'] = raw
                                        entries['_domicilios_modal_display'] = domicilios
                                except Exception:
                                    pass
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
                    # Campo de texto normal, excepto casos especiales: 'direccion' y 'norma'
                    if key == 'direccion':
                        # combobox que se rellenará según cliente seleccionado
                        ent = ctk.CTkComboBox(
                            field_frame,
                            values=['Seleccione un domicilio...'],
                            font=FONT_SMALL,
                            dropdown_font=FONT_SMALL,
                            state='disabled',
                            height=35,
                            corner_radius=8
                        )
                        ent.pack(fill='x')
                        # si ya hay direccion en datos, mostrarla inmediatamente y crear mapping raw
                        if datos and datos.get('direccion'):
                            try:
                                v = str(datos.get('direccion'))
                                # preparar raw mapping a partir de campos en 'datos' cuando estén disponibles
                                raw_item = {
                                    'CALLE Y NO': datos.get('calle_numero') or (v.split(',')[0].strip() if v else ''),
                                    'COLONIA O POBLACION': datos.get('colonia') or (v.split(',')[1].strip() if len(v.split(','))>1 else ''),
                                    'MUNICIPIO O ALCADIA': datos.get('municipio') or (v.split(',')[2].strip() if len(v.split(','))>2 else ''),
                                    'CIUDAD O ESTADO': datos.get('ciudad_estado') or (v.split(',')[3].strip() if len(v.split(','))>3 else ''),
                                    'CP': datos.get('cp') or datos.get('CP') or (v.split(',')[-1].strip() if len(v.split(','))>0 else '')
                                }
                                entries['_domicilios_modal_raw'] = [raw_item]
                                entries['_domicilios_modal_display'] = [v]
                                ent.configure(values=[v], state='readonly')
                                ent.set(v)
                                # also set selected_address_raw so save picks it up
                                try:
                                    selected_address_raw.clear()
                                    selected_address_raw.update(raw_item)
                                except Exception:
                                    pass
                            except Exception:
                                pass
                    elif key == 'norma':
                        # Crear un contenedor aquí (justo debajo de Fecha Termino)
                        # y usarlo más adelante como padre del listado de normas.
                        try:
                            norma_container = ctk.CTkFrame(field_frame, fg_color='transparent')
                            norma_container.pack(fill='both', expand=True, pady=(-8, 0))
                            entries['_norma_container'] = norma_container
                        except Exception:
                            entries['_norma_container'] = field_frame
                        ent = None
                    else:
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
        
        # Helper: inserción rápida de inspectores y selección múltiple de normas
        try:
            # Colocar inspectores en la columna 0, normas en la columna 1 para mejor visibilidad
            # Inspectores: columna 3 con checkboxes; cada click muestra normas acreditadas
            insp_frame = ctk.CTkFrame(col_frames[3], fg_color='transparent')
            insp_frame.pack(fill='both', expand=True)
            ctk.CTkLabel(insp_frame, text='Listado de Inspectores', font=FONT_SMALL, text_color=STYLE['texto_oscuro']).pack(anchor='w')
            # Reducir la altura del listado de inspectores a la mitad
            scroll_insp = ctk.CTkScrollableFrame(insp_frame, height=100, fg_color='transparent')
            scroll_insp.pack(fill='both', expand=True, pady=(6,6), padx=(6,0))

            insp_checks = []
            # Determinar inspectores seleccionados a partir de los datos de la visita
            try:
                existing_raw_insp = (datos.get('supervisores_tabla') or datos.get('nfirma1') or '') if datos else ''
                selected_inspectores = [s.strip() for s in str(existing_raw_insp).split(',') if s.strip()]
            except Exception:
                selected_inspectores = []
            last_insp_normas_label = ctk.CTkLabel(insp_frame, text='', font=("Inter", 11), text_color=STYLE['texto_oscuro'])
            last_insp_normas_label.pack(anchor='w', pady=(6,4))

            # mapping para labels de estado (marca verde)
            inspector_status_labels = {}

            def _on_insp_click(nombre, var):
                try:
                    normas = firmas_map.get(nombre, []) if 'firmas_map' in locals() or 'firmas_map' in globals() else []
                    if normas:
                        lines = [f"{i}. {n}" for i, n in enumerate(normas, start=1)]
                        last_insp_normas_label.configure(text="\n".join(lines))
                    else:
                        last_insp_normas_label.configure(text='(Sin normas acreditadas)')
                except Exception:
                    try:
                        last_insp_normas_label.configure(text='')
                    except Exception:
                        pass

            def update_inspector_statuses():
                try:
                    norma_checks_local = entries.get('_norma_checks') or []
                    selected_norms = [nm for nm, v in norma_checks_local if getattr(v, 'get', lambda: '0')() in ('1', 'True', 'true')]
                    for nombre, lbl in inspector_status_labels.items():
                        try:
                            acc = set(firmas_map.get(nombre, []) or [])
                            ok = False
                            if selected_norms:
                                ok = set(selected_norms).issubset(acc)
                            else:
                                ok = False
                            if ok:
                                lbl.configure(text='✓', text_color=STYLE['exito'])
                            else:
                                lbl.configure(text='', text_color=STYLE['texto_oscuro'])
                        except Exception:
                            try:
                                lbl.configure(text='', text_color=STYLE['texto_oscuro'])
                            except Exception:
                                pass
                except Exception:
                    pass

            for nombre in (inspectores_list or []):
                try:
                    var = ctk.StringVar(value='0')
                    # fila contenedora para checkbox + estado
                    row = ctk.CTkFrame(scroll_insp, fg_color='transparent')
                    row.pack(fill='x', pady=(2,2), padx=(2,0))
                    row.grid_columnconfigure(0, weight=1)
                    chk = ctk.CTkCheckBox(row, text=nombre, variable=var, onvalue='1', offvalue='0', command=lambda n=nombre, v=var: _on_insp_click(n, v), font=("Inter", 11))
                    chk.grid(row=0, column=0, sticky='w')
                    status_lbl = ctk.CTkLabel(row, text='', font=("Inter", 11), text_color=STYLE['exito'])
                    status_lbl.grid(row=0, column=1, sticky='w', padx=(6,0))
                    inspector_status_labels[nombre] = status_lbl
                    # trace para actualizar estado cuando cambien normas o el propio checkbox
                    try:
                        var.trace_add('write', lambda *a, _n=nombre: update_inspector_statuses())
                    except Exception:
                        pass
                    # Si en los datos de la visita ya viene este inspector, marcarlo
                    try:
                        if nombre in selected_inspectores:
                            try:
                                var.set('1')
                            except Exception:
                                pass
                    except Exception:
                        pass
                    insp_checks.append((nombre, var))
                except Exception:
                    continue
            entries['_inspectores_checks'] = insp_checks

            # Normas: lista de checkboxes para selección múltiple (columna central)
            # Usar el contenedor creado en el lugar del campo 'norma' para
            # que el listado quede justo debajo de Fecha Termino.
            parent_container = entries.get('_norma_container') if entries.get('_norma_container') else col_frames[1]
            normas_frame = ctk.CTkFrame(parent_container, fg_color='transparent')
            normas_frame.pack(fill='both', expand=False, pady=(0, 0))
            ctk.CTkLabel(normas_frame, text='Listado de Normas', font=FONT_SMALL, text_color=STYLE['texto_oscuro']).pack(anchor='w', pady=(0,4))
            # Quitar el scroll del listado de normas para mostrarlas directamente
            normas_list_frame = ctk.CTkFrame(normas_frame, fg_color='transparent')
            normas_list_frame.pack(fill='both', expand=True, pady=(0,2))
            # crear checkvariables
            norma_checks = []
            selected_normas = [n.strip() for n in (datos.get('norma') or '').split(',') if n.strip()] if datos else []
            for nm in (normas_list or []):
                try:
                    var = ctk.StringVar(value='0')
                    chk = ctk.CTkCheckBox(normas_list_frame, text=nm, variable=var, onvalue='1', offvalue='0')
                    chk.pack(anchor='w', fill='x')
                    # cuando cambie una norma, actualizar estados de inspectores
                    try:
                        var.trace_add('write', lambda *a: update_inspector_statuses())
                    except Exception:
                        pass
                    if nm in selected_normas:
                        try:
                            var.set('1')
                        except Exception:
                            pass
                    norma_checks.append((nm, var))
                except Exception:
                    continue
            entries['_norma_checks'] = norma_checks
        except Exception:
            pass

        # Frame para botones
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 20), padx=25)
        
        def _guardar():
            # Recoger datos de todos los campos
            payload = {}
            for key, entry in entries.items():
                # Ignorar helpers internos (prefijo _ ) o valores que no sean widgets
                if str(key).startswith('_'):
                    continue
                # Si el objeto no tiene método get(), omitimos
                if not hasattr(entry, 'get'):
                    continue
                if key in ["cliente", "estatus"]:
                    # Para combobox, obtener el valor seleccionado
                    raw_value = entry.get()
                    if key == "cliente" and raw_value == "Seleccione un cliente...":
                        raw_value = ""
                    value = raw_value
                else:
                    raw_value = entry.get()
                    value = raw_value.strip() if isinstance(raw_value, str) else raw_value
                payload[key] = value
            
                # (normas) -- se recogen MÁS ABAJO fuera del bucle para evitar
                # recolecciones múltiples/duplicadas mientras iteramos `entries`.

                # Asegurar que la lista de supervisores se guarda a partir de los checkboxes de inspectores (si existen)
                try:
                    # Preferir checkboxes de inspectores si existen
                    insp_checks = entries.get('_inspectores_checks') or []
                    selected = [name for name, var in insp_checks if getattr(var, 'get', lambda: '0')() in ('1', 'True', 'true')]
                    if selected:
                        joined = ', '.join(selected)
                        payload['supervisores_tabla'] = joined
                        payload['nfirma1'] = joined
                    else:
                        # fallback: combinar inspectores previos con los del payload.nfirma1 (si el usuario escribió/insertó)
                        existing = []
                        try:
                            existing_raw = (datos.get('supervisores_tabla') or datos.get('nfirma1') or '') if datos else ''
                            existing = [s.strip() for s in str(existing_raw).split(',') if s.strip()]
                        except Exception:
                            existing = []
                        new_list = []
                        try:
                            nf = (payload.get('nfirma1') or '')
                            new_list = [s.strip() for s in str(nf).split(',') if s.strip()]
                        except Exception:
                            new_list = []
                        merged = []
                        for s in (existing + new_list):
                            if s and s not in merged:
                                merged.append(s)
                        if merged:
                            joined = ', '.join(merged)
                            payload['supervisores_tabla'] = joined
                            payload['nfirma1'] = joined
                except Exception:
                    pass

                # Recolectar normas seleccionadas (fuera del bucle) y deduplicar
            try:
                norma_checks = entries.get('_norma_checks') or []
                sel = []
                for nm, var in norma_checks:
                    try:
                        v = getattr(var, 'get', lambda: '0')()
                    except Exception:
                        v = '0'
                    if str(v).lower() in ('1', 'true', 'yes'):
                        if nm not in sel:
                            sel.append(nm)
                if sel:
                    payload['norma'] = ', '.join(sel)
                else:
                    # si no hay selección, dejar vacío (o mantener lo que ya exista)
                    payload['norma'] = payload.get('norma', '') or ''
            except Exception:
                pass

                # Validaciones
            if not payload.get("cliente"):
                messagebox.showwarning("Validación", "Por favor seleccione un cliente")
                return
            
            if not payload.get("estatus"):
                payload["estatus"] = "En proceso"
            
            # Conservar horas originales si el formulario no las incluye
            try:
                for h in ("hora_inicio", "hora_termino", "hora_inicio_24h", "hora_termino_24h"):
                    if (h not in payload or payload.get(h) in (None, "", " ")) and datos.get(h):
                        payload[h] = datos.get(h)
            except Exception:
                pass

            # Añadir componentes de dirección desde la selección modal si existen
            try:
                # Determinar valor desplegado del domicilio (si existe)
                dir_widget = entries.get('direccion') if isinstance(entries.get('direccion', None), ctk.CTkComboBox) else None
                dir_display = None
                if dir_widget:
                    try:
                        dir_display = dir_widget.get()
                    except Exception:
                        dir_display = None

                # Si existe la lista raw en entries, intentar mapear por índice
                raw_mapped = None
                try:
                    display_list = entries.get('_domicilios_modal_display') or []
                    raw_list = entries.get('_domicilios_modal_raw') or []
                    if dir_display and display_list and raw_list and dir_display in display_list:
                        idx = display_list.index(dir_display)
                        if idx < len(raw_list):
                            raw_mapped = raw_list[idx]
                except Exception:
                    raw_mapped = None

                # Priorizar raw_mapped, luego selected_address_raw (si fue seteado por el handler),
                # luego intentar parsear dir_display como respaldo
                source_raw = raw_mapped or (selected_address_raw if selected_address_raw else None)

                if source_raw:
                    # Guardar display completo en 'direccion' pero asegurar que 'calle_numero'
                    # almacene únicamente la 'CALLE Y NO' (o la mejor alternativa)
                    payload['direccion'] = dir_display or payload.get('direccion') or source_raw.get('CALLE Y NO') or source_raw.get('CALLE')
                    payload['calle_numero'] = source_raw.get('CALLE Y NO') or source_raw.get('CALLE') or payload.get('direccion') or payload.get('calle_numero')
                    payload['colonia'] = source_raw.get('COLONIA O POBLACION') or source_raw.get('COLONIA') or source_raw.get('colonia') or payload.get('colonia')
                    payload['municipio'] = source_raw.get('MUNICIPIO O ALCADIA') or source_raw.get('MUNICIPIO') or source_raw.get('municipio') or payload.get('municipio')
                    payload['ciudad_estado'] = source_raw.get('CIUDAD O ESTADO') or source_raw.get('CIUDAD') or source_raw.get('ciudad_estado') or payload.get('ciudad_estado')
                    payload['cp'] = str(source_raw.get('CP') or source_raw.get('cp') or payload.get('cp') or '')
                else:
                    # Si no hay raw, pero hay texto desplegado, intentar descomponerlo por comas
                    if dir_display and dir_display not in (None, '', 'Seleccione un domicilio...'):
                        payload['direccion'] = dir_display
                        parts = [p.strip() for p in dir_display.split(',') if p.strip()]
                        # heurística: último token suele ser CP o ciudad; asignar por posición
                        if parts:
                            payload['colonia'] = parts[1] if len(parts) > 1 else payload.get('colonia')
                            payload['municipio'] = parts[2] if len(parts) > 2 else payload.get('municipio')
                            payload['ciudad_estado'] = parts[3] if len(parts) > 3 else payload.get('ciudad_estado')
                            # intentar extraer CP numérico
                            for p in parts[::-1]:
                                s = ''.join(ch for ch in p if ch.isdigit())
                                if s:
                                    payload['cp'] = s
                                    break
            except Exception:
                pass
            # Antes de actualizar, sincronizar también atributos de instancia para UI principal
            try:
                if payload.get('direccion'):
                    self.direccion_seleccionada = payload.get('direccion')
                # sincronizar alias
                if payload.get('calle_numero') and not getattr(self, 'direccion_seleccionada', None):
                    self.direccion_seleccionada = payload.get('calle_numero')
                if payload.get('colonia'):
                    self.colonia_seleccionada = payload.get('colonia')
                if payload.get('municipio'):
                    self.municipio_seleccionado = payload.get('municipio')
                if payload.get('ciudad_estado'):
                    self.ciudad_seleccionada = payload.get('ciudad_estado')
                if payload.get('cp'):
                    self.cp_seleccionado = payload.get('cp')
                # mantener domicilio_seleccionado como display
                if payload.get('direccion'):
                    self.domicilio_seleccionado = payload.get('direccion')
            except Exception:
                pass

            # DEBUG: mostrar payload recogido del modal
            try:
                print(f"[DEBUG] modal _guardar payload: {json.dumps(payload, ensure_ascii=False)}")
            except Exception:
                print(f"[DEBUG] modal _guardar payload: {payload}")

            # usar id seguro (soporta _id, id o folio) para actualizar
            target_id = datos.get('_id') or datos.get('id') or datos.get('folio_visita') or datos.get('folio_acta')
            if not target_id:
                messagebox.showerror("Error", "No se pudo determinar el identificador de la visita para actualizar")
                return
            self.hist_update_visita(target_id, payload)
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

    # ------------------ EVIDENCIAS: carga y persistencia ------------------
    def _load_evidence_paths(self):
        """Carga el archivo `data/evidence_paths.json` si existe."""
        try:
            data_file = os.path.join(os.path.dirname(__file__), "data", "evidence_paths.json")
            if os.path.exists(data_file):
                with open(data_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            return {}
        except Exception:
            return {}

    def _save_evidence_path(self, group, path):
        """Guarda la ruta `path` bajo la clave `group` en `data/evidence_paths.json`."""
        data_file = os.path.join(os.path.dirname(__file__), "data", "evidence_paths.json")
        os.makedirs(os.path.dirname(data_file), exist_ok=True)
        data = self._load_evidence_paths() or {}
        existing = data.get(group, [])
        if path not in existing:
            existing.append(path)
        data[group] = existing
        with open(data_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def configurar_carpeta_evidencias(self):
        """Abre un modal para elegir a qué grupo se guardará la ruta y seleccionar carpeta."""
        modal = ctk.CTkToplevel(self)
        modal.title("Configurar Carpeta de Evidencias")
        modal.geometry("520x260")
        modal.transient(self)
        modal.grab_set()

        ctk.CTkLabel(modal, text="Seleccione el grupo y luego elija la carpeta de evidencias:",
                     font=FONT_SMALL, text_color=STYLE["texto_oscuro"]).pack(anchor="w", padx=16, pady=(12, 8))

        var_grupo = ctk.StringVar(value="grupo_axo")
        opciones = [
            ("Grupo Axo (varios clientes)", "grupo_axo"),
            ("Bosch", "bosch"),
            ("Unilever", "unilever"),
        ]

        for texto, valor in opciones:
            ctk.CTkRadioButton(modal, text=texto, variable=var_grupo, value=valor).pack(anchor="w", padx=20, pady=6)

        # Mostrar rutas actualmente guardadas
        rutas_frame = ctk.CTkFrame(modal, fg_color="transparent")
        rutas_frame.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        lbl_actual = ctk.CTkLabel(rutas_frame, text="Rutas guardadas:", font=FONT_SMALL, text_color=STYLE["texto_oscuro"]) 
        lbl_actual.pack(anchor="w")

        lista_rutas = ctk.CTkLabel(rutas_frame, text="(ninguna)", font=("Inter", 10), text_color=STYLE["texto_claro"], wraplength=480)
        lista_rutas.pack(anchor="w", pady=(4, 0))

        def _refrescar_rutas():
            data = self._load_evidence_paths() or {}
            lines = []
            for g, lst in data.items():
                lines.append(f"{g}: \n  " + "\n  ".join(lst))
            lista_rutas.configure(text="\n\n".join(lines) if lines else "(ninguna)")

        _refrescar_rutas()

        def elegir_carpeta():
            grp = var_grupo.get()
            carpeta = filedialog.askdirectory(title="Seleccionar carpeta de evidencias")
            if not carpeta:
                return
            try:
                self._save_evidence_path(grp, carpeta)
                messagebox.showinfo("Guardado", f"Ruta guardada para '{grp}':\n{carpeta}")
                _refrescar_rutas()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar la ruta:\n{e}")

        btn_frame = ctk.CTkFrame(modal, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)

        ctk.CTkButton(btn_frame, text="Seleccionar carpeta y guardar", command=elegir_carpeta,
                      fg_color=STYLE["primario"], text_color=STYLE["secundario"], height=36).pack(side="left", padx=12)
        ctk.CTkButton(btn_frame, text="Cerrar", command=modal.destroy, height=36).pack(side="right", padx=12)

    # ------------------ PEGADO EVIDENCIAS (botones UI) ------------------
    def _run_script_and_notify(self, fn):
        try:
            fn()
            try:
                messagebox.showinfo("Pegado", "Proceso de pegado finalizado. Revise el registro de fallos si corresponde.")
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror("Error pegado", f"Error al ejecutar el proceso de pegado:\n{e}")
            except Exception:
                pass

    def _call_pegado_script(self, script_filename, func_name, ruta_docs=None, ruta_imgs=None):
        """Carga dinámicamente los módulos de Pegado sin persistir rutas del usuario.
        Si se proveen `ruta_docs` y `ruta_imgs`, inyecta una implementación de
        `obtener_rutas()` que devuelve esas rutas y convierte `guardar_config` en no-op.
        """
        try:
            base = os.path.join(os.path.dirname(__file__), "Pegado de Evidenvia Fotografica")
            main_path = os.path.join(base, "main.py")

            # Asegurar que la carpeta del pegado esté en sys.path para que
            # los imports relativos como `from registro_fallos import ...`
            # funcionen cuando el script se ejecute.
            added_to_path = False
            if base not in sys.path:
                sys.path.insert(0, base)
                added_to_path = True

            # Cargar module 'main' desde archivo y parchear
            spec_main = importlib.util.spec_from_file_location("main_pegado", main_path)
            mod_main = importlib.util.module_from_spec(spec_main)
            spec_main.loader.exec_module(mod_main)

            # Evitar persistir config
            try:
                mod_main.guardar_config = lambda x: None
            except Exception:
                pass

            if ruta_docs and ruta_imgs:
                mod_main.obtener_rutas = lambda: (ruta_docs, ruta_imgs)

            # Hacer disponible como 'main' para que los scripts que hacen `from main import ...` funcionen
            old_main = sys.modules.get('main')
            sys.modules['main'] = mod_main

            # Cargar y ejecutar el script solicitado
            script_path = os.path.join(base, script_filename)
            spec = importlib.util.spec_from_file_location("pegado_script", script_path)
            mod_script = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod_script)

            fn = getattr(mod_script, func_name, None)
            if not callable(fn):
                messagebox.showerror("Error", f"Función {func_name} no encontrada en {script_filename}")
                return

            # Ejecutar en hilo para no bloquear la UI
            threading.Thread(target=lambda: self._run_script_and_notify(fn), daemon=True).start()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar el pegado:\n{e}")
        finally:
            # Restaurar 'main' anterior si existía
            try:
                if old_main is not None:
                    sys.modules['main'] = old_main
                else:
                    sys.modules.pop('main', None)
                # Quitar la ruta añadida al sys.path
                try:
                    if added_to_path and base in sys.path:
                        sys.path.remove(base)
                except Exception:
                    pass
            except Exception:
                pass

    def handle_pegado_simple(self):
        # Guardar únicamente la ruta de imágenes para usarla posteriormente
        ruta_imgs = filedialog.askdirectory(title="Seleccionar carpeta de imágenes para evidencias (se guardará)")
        if not ruta_imgs:
            return
        try:
            # Guardar la ruta bajo un grupo genérico 'manual_pegado'
            self._save_evidence_path('manual_pegado', ruta_imgs)
            messagebox.showinfo("Pegado guardado", "Ruta de imágenes guardada. Cuando genere los dictámenes, se buscarán evidencias en esta carpeta.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la ruta de evidencias:\n{e}")

    def handle_pegado_carpetas(self):
        ruta_imgs = filedialog.askdirectory(title="Seleccionar carpeta raíz de carpetas por código (se guardará)")
        if not ruta_imgs:
            return
        try:
            self._save_evidence_path('manual_pegado', ruta_imgs)
            messagebox.showinfo("Pegado guardado", "Ruta de carpetas guardada. Cuando genere los dictámenes, se buscarán evidencias en estas carpetas.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la ruta de evidencias:\n{e}")

    def handle_pegado_indice(self):
        ruta_imgs = filedialog.askdirectory(title="Seleccionar carpeta de imágenes para índice (se guardará)")
        if not ruta_imgs:
            return
        try:
            self._save_evidence_path('manual_pegado', ruta_imgs)
            messagebox.showinfo("Pegado guardado", "Ruta de imágenes guardada. Al generar dictámenes y subir el Excel de índice, se usarán estas imágenes.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la ruta de evidencias:\n{e}")

# ================== EJECUCIÓN ================== #
if __name__ == "__main__":
    app = SistemaDictamenesVC()
    app.mainloop()
