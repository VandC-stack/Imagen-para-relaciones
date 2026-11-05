"""Tema visual del Sistema V&C: colores, fuentes, iconos y estilo de widgets.

`aplicar_tema()` debe ejecutarse antes de crear cualquier widget, porque
CustomTkinter lee sus colores por defecto al construir cada control.
"""
import os

import customtkinter as ctk
from PIL import Image, ImageChops, ImageDraw, ImageFont
from tkinter import ttk

FONT_FAMILY = "Segoe UI"

STYLE = {
    # Marca
    "primario": "#ECD925",
    "primario_hover": "#D4BF22",
    "secundario": "#282828",
    "secundario_hover": "#1A1A1A",
    # Estados
    "exito": "#008D53",
    "advertencia": "#B45309",
    "peligro": "#C8372B",
    # Superficies
    "fondo": "#F3F4F6",
    "surface": "#FFFFFF",
    "neutro": "#E5E7EB",
    "neutro_hover": "#D1D5DB",
    # Texto
    "texto_oscuro": "#282828",
    "texto_claro": "#FFFFFF",
    "texto_secundario": "#5F6673",
    # Bordes
    "borde": "#E5E7EB",
    "borde_panel": "#E5E7EB",
    "borde_fuerte": "#C4C9D1",
    # Barra lateral
    "sidebar_hover": "#3A3A3A",
    "sidebar_texto": "#D8DBE0",
    "sidebar_tenue": "#9CA3AF",
    # Selección en tablas
    "seleccion": "#D9F0FF",
}

FONT_TITLE = (FONT_FAMILY, 22, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 16, "bold")
FONT_SECCION = (FONT_FAMILY, 13, "bold")
FONT_LABEL = (FONT_FAMILY, 13)
FONT_SMALL = (FONT_FAMILY, 12)

# Glifos de la fuente "Segoe MDL2 Assets", incluida en Windows 10 y 11.
ICONOS = {
    "principal": "\uE80F",
    "historial": "\uE81C",
    "clientes": "\uE716",
    "inspectores": "\uE7EE",
    "reportes": "\uE9D2",
    "descarga": "\uE896",
    "usuarios": "\uE7EF",
    "salir": "\uF3B1",
}


# Opciones de tkcalendar.Calendar. El tamaño va en píxeles (negativo) porque la
# aplicación fija `tk scaling` en 1.0 y eso distorsiona los tamaños en puntos.
ESTILO_CALENDARIO = {
    "font": (FONT_FAMILY, -15),
    "showweeknumbers": False,
    "borderwidth": 0,
    "background": STYLE["secundario"],
    "foreground": STYLE["texto_claro"],
    "bordercolor": STYLE["borde"],
    "headersbackground": STYLE["fondo"],
    "headersforeground": STYLE["texto_secundario"],
    "normalbackground": STYLE["surface"],
    "normalforeground": STYLE["texto_oscuro"],
    "weekendbackground": STYLE["surface"],
    "weekendforeground": STYLE["texto_oscuro"],
    "othermonthbackground": STYLE["fondo"],
    "othermonthforeground": STYLE["sidebar_tenue"],
    "othermonthwebackground": STYLE["fondo"],
    "othermonthweforeground": STYLE["sidebar_tenue"],
    "selectbackground": STYLE["primario"],
    "selectforeground": STYLE["secundario"],
}


def _par(color):
    """CustomTkinter espera un color para modo claro y otro para modo oscuro."""
    return [color, color]


def aplicar_tema():
    """Sustituye los colores y la fuente por defecto de CustomTkinter."""
    ctk.set_appearance_mode("light")
    # CustomTkinter oculta y vuelve a mostrar cada ventana para pintar la barra de
    # título, lo que se ve como un parpadeo. En modo claro no hace falta.
    ctk.CTk._deactivate_windows_window_header_manipulation = True
    ctk.CTkToplevel._deactivate_windows_window_header_manipulation = True

    tema = ctk.ThemeManager.theme
    s = STYLE

    tema["CTkFont"]["family"] = FONT_FAMILY
    tema["CTk"]["fg_color"] = _par(s["fondo"])
    tema["CTkToplevel"]["fg_color"] = _par(s["fondo"])
    tema["CTkFrame"].update(
        fg_color=_par(s["surface"]),
        top_fg_color=_par(s["surface"]),
        border_color=_par(s["borde_panel"]),
    )
    tema["CTkButton"].update(
        fg_color=_par(s["secundario"]),
        hover_color=_par(s["secundario_hover"]),
        border_color=_par(s["borde_fuerte"]),
        text_color=_par(s["texto_claro"]),
        text_color_disabled=_par(s["sidebar_tenue"]),
    )
    tema["CTkLabel"]["text_color"] = _par(s["texto_oscuro"])
    tema["CTkEntry"].update(
        fg_color=_par(s["surface"]),
        border_color=_par(s["borde_fuerte"]),
        text_color=_par(s["texto_oscuro"]),
        placeholder_text_color=_par(s["sidebar_tenue"]),
        border_width=1,
    )
    tema["CTkComboBox"].update(
        fg_color=_par(s["surface"]),
        border_color=_par(s["borde_fuerte"]),
        button_color=_par(s["borde_fuerte"]),
        button_hover_color=_par(s["sidebar_tenue"]),
        text_color=_par(s["texto_oscuro"]),
        text_color_disabled=_par(s["sidebar_tenue"]),
        border_width=1,
    )
    tema["CTkOptionMenu"].update(
        fg_color=_par(s["surface"]),
        button_color=_par(s["borde_fuerte"]),
        button_hover_color=_par(s["sidebar_tenue"]),
        text_color=_par(s["texto_oscuro"]),
    )
    tema["DropdownMenu"].update(
        fg_color=_par(s["surface"]),
        hover_color=_par(s["neutro"]),
        text_color=_par(s["texto_oscuro"]),
    )
    for control in ("CTkCheckBox", "CTkRadioButton"):
        tema[control].update(
            fg_color=_par(s["exito"]),
            border_color=_par(s["borde_fuerte"]),
            hover_color=_par(s["exito"]),
            text_color=_par(s["texto_oscuro"]),
            text_color_disabled=_par(s["texto_secundario"]),
        )
    tema["CTkCheckBox"]["border_width"] = 2
    tema["CTkProgressBar"].update(
        fg_color=_par(s["neutro"]),
        progress_color=_par(s["primario"]),
    )
    tema["CTkScrollbar"].update(
        fg_color="transparent",
        button_color=_par(s["borde_fuerte"]),
        button_hover_color=_par(s["sidebar_tenue"]),
    )
    tema["CTkTextbox"].update(
        fg_color=_par(s["surface"]),
        border_color=_par(s["borde_fuerte"]),
        text_color=_par(s["texto_oscuro"]),
        scrollbar_button_color=_par(s["borde_fuerte"]),
        scrollbar_button_hover_color=_par(s["sidebar_tenue"]),
    )


class BotonAccion(ctk.CTkButton):
    """Botón principal que se ve apagado mientras está deshabilitado."""

    def __init__(self, *args, fg_color, state="normal", **kwargs):
        self._color_activo = fg_color
        super().__init__(*args, fg_color=self._color_segun(state), state=state, **kwargs)

    def _color_segun(self, state):
        return self._color_activo if state == "normal" else STYLE["neutro"]

    def configure(self, require_redraw=False, **kwargs):
        if "state" in kwargs:
            kwargs["fg_color"] = self._color_segun(kwargs["state"])
        super().configure(require_redraw=require_redraw, **kwargs)


def configurar_tablas(ventana):
    """Aplica el estilo común de las tablas ttk (Historial, Clientes, Inspectores).

    ttk no escala con CustomTkinter, así que fuentes y filas se calculan en
    píxeles según la escala de pantalla de `ventana`.
    """
    s = STYLE
    escala = ctk.ScalingTracker.get_widget_scaling(ventana)
    estilo = ttk.Style()
    try:
        # 'clam' es el tema de ttk que permite colorear los encabezados.
        estilo.theme_use("clam")
    except Exception:
        pass
    estilo.configure(
        "clientes.Treeview",
        font=(FONT_FAMILY, -round(12 * escala)),
        rowheight=round(30 * escala),
        background=s["surface"],
        fieldbackground=s["surface"],
        foreground=s["texto_oscuro"],
        borderwidth=0,
    )
    estilo.layout("clientes.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    estilo.configure(
        "clientes.Treeview.Heading",
        font=(FONT_FAMILY, -round(12 * escala), "bold"),
        background=s["fondo"],
        foreground=s["texto_secundario"],
        relief="flat",
        borderwidth=0,
        padding=(6, 7),
    )
    estilo.map(
        "clientes.Treeview.Heading",
        background=[("active", s["neutro"])],
        foreground=[("active", s["texto_oscuro"])],
    )
    estilo.map(
        "clientes.Treeview",
        background=[("selected", s["seleccion"])],
        foreground=[("selected", s["texto_oscuro"])],
    )
    estilo.configure(
        "Vertical.TScrollbar",
        troughcolor=s["surface"],
        background=s["borde_fuerte"],
        arrowcolor=s["texto_oscuro"],
    )


def icono(nombre, color, tamano=18):
    """Devuelve un icono de `ICONOS` como CTkImage, o None si falta la fuente."""
    ruta = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segmdl2.ttf")
    lado = tamano * 4  # Se dibuja en grande para que se vea nítido al reducirlo.
    try:
        fuente = ImageFont.truetype(ruta, int(lado * 0.8))
    except OSError:
        return None
    imagen = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    ImageDraw.Draw(imagen).text((lado / 2, lado / 2), ICONOS[nombre], font=fuente, fill=color, anchor="mm")
    return ctk.CTkImage(imagen, size=(tamano, tamano))


def logo_monocromo(ruta, color, alto):
    """Devuelve el logo en un solo color como CTkImage, recortado a su contenido.

    El logo original es negro, así que no se vería sobre un fondo oscuro. Las
    zonas oscuras y opacas se conservan; las claras o transparentes se eliminan.
    """
    original = Image.open(ruta).convert("RGBA")
    oscuridad = ImageChops.invert(original.convert("L"))
    mascara = ImageChops.multiply(oscuridad, original.getchannel("A"))
    caja = mascara.getbbox()
    if caja:
        mascara = mascara.crop(caja)
    logo = Image.new("RGBA", mascara.size, color)
    logo.putalpha(mascara)
    ancho = max(1, round(logo.width * alto / logo.height))
    return ctk.CTkImage(logo, size=(ancho, alto))
