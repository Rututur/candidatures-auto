"""Light and dark appearance of the window, drawn with ttk's "clam" engine.

Only the standard library is used. Colours live in one Palette per mode; apply()
configures the ttk styles and restyle() handles the classic Tk widgets (Text,
Menu, windows), which ttk styles do not reach.
"""

from __future__ import annotations

import ctypes
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import font as tkfont
from tkinter import ttk

from .workbook import (
    STATUS_ACCEPTED,
    STATUS_FOLLOWED_UP,
    STATUS_INTERVIEW,
    STATUS_LETTER_READY,
    STATUS_NO_ANSWER,
    STATUS_REFUSED,
    STATUS_REPLIED,
    STATUS_SENT,
)

LIGHT_MODE = "light"
DARK_MODE = "dark"
FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
FONT_SIZE = 10

# Row colour of the spreadsheet for each status (key of Palette.status).
STATUS_GROUPS = {
    STATUS_SENT: "blue",
    STATUS_FOLLOWED_UP: "blue",
    STATUS_REPLIED: "green",
    STATUS_INTERVIEW: "green",
    STATUS_ACCEPTED: "green",
    STATUS_REFUSED: "red",
    STATUS_LETTER_READY: "amber",
    STATUS_NO_ANSWER: "grey",
}


@dataclass(frozen=True)
class Palette:
    dark: bool
    background: str
    surface: str
    sidebar: str
    heading: str
    text: str
    muted: str
    border: str
    hover: str
    accent: str
    accent_hover: str
    on_accent: str
    accent_soft: str
    accent_text: str
    success: str
    warning: str
    danger: str
    status: dict[str, str]


PALETTES = {
    LIGHT_MODE: Palette(
        dark=False,
        background="#F6F5FB",
        surface="#FFFFFF",
        sidebar="#EDEAF7",
        heading="#F1EEFA",
        text="#1F1B2E",
        muted="#625E76",
        border="#D9D4E8",
        hover="#E6E1F5",
        accent="#6D4AE8",
        accent_hover="#5A38D2",
        on_accent="#FFFFFF",
        accent_soft="#DED5FC",
        accent_text="#4A2BB8",
        success="#1E7B34",
        warning="#A8450A",
        danger="#B42318",
        status={"blue": "#E7F0FD", "green": "#E3F5E8", "red": "#FBE6E8", "amber": "#FDF2D8", "grey": "#EEEDF2"},
    ),
    DARK_MODE: Palette(
        dark=True,
        background="#16141F",
        surface="#211E2E",
        sidebar="#1B1927",
        heading="#2A2639",
        text="#ECEAF4",
        muted="#A9A6BB",
        border="#3A3650",
        hover="#2E2A40",
        accent="#9B7DFF",
        accent_hover="#AE95FF",
        on_accent="#16141F",
        accent_soft="#372D60",
        accent_text="#C9B8FF",
        success="#5CCB7A",
        warning="#F2A65A",
        danger="#F4807A",
        status={"blue": "#1E2B45", "green": "#1C3527", "red": "#3D2228", "amber": "#3B3119", "grey": "#2B2937"},
    ),
}


def apply(root: tk.Tk, mode: str) -> Palette:
    """Switch the whole window to a mode and return its palette."""
    palette = PALETTES.get(mode, PALETTES[LIGHT_MODE])
    style = ttk.Style(root)
    style.theme_use("clam")
    _configure_styles(style, palette, root)
    restyle(root, palette)
    return palette


def restyle(widget, palette: Palette) -> None:
    """Colour a window and the classic Tk widgets it contains."""
    if isinstance(widget, (tk.Tk, tk.Toplevel)):
        widget.configure(background=palette.background)
        _set_title_bar(widget, palette)
    elif isinstance(widget, tk.Text):
        widget.configure(
            background=palette.surface, foreground=palette.text, insertbackground=palette.text,
            selectbackground=palette.accent, selectforeground=palette.on_accent, inactiveselectbackground=palette.accent_soft,
            relief="flat", borderwidth=0, highlightthickness=1,
            highlightbackground=palette.border, highlightcolor=palette.accent,
        )
    elif isinstance(widget, tk.Menu):
        widget.configure(
            background=palette.surface, foreground=palette.text, relief="flat", borderwidth=1,
            activebackground=palette.accent, activeforeground=palette.on_accent,
        )
    elif isinstance(widget, ttk.Combobox):
        # The drop-down list is a Tcl-only child: winfo_children() would fail on it.
        popdown = widget.tk.call("ttk::combobox::PopdownWindow", widget)
        widget.tk.call(
            f"{popdown}.f.l", "configure", "-background", palette.surface, "-foreground", palette.text,
            "-selectbackground", palette.accent, "-selectforeground", palette.on_accent,
        )
        return
    for child in widget.winfo_children():
        restyle(child, palette)


def _set_title_bar(window, palette: Palette) -> None:
    """Follow the mode in the Windows title bar (ignored where it is not supported)."""
    if sys.platform != "win32":
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1 if palette.dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except (AttributeError, OSError):
        pass


def _flat(style: ttk.Style, name: str, color: str, **options) -> None:
    """Configure a style with no visible 3D edges."""
    style.configure(name, background=color, lightcolor=color, darkcolor=color, **options)


def _configure_styles(style: ttk.Style, p: Palette, root: tk.Tk) -> None:
    font = (FONT_FAMILY, FONT_SIZE)
    bold = (FONT_FAMILY, FONT_SIZE, "bold")
    row_height = tkfont.Font(root=root, family=FONT_FAMILY, size=FONT_SIZE).metrics("linespace") + 14

    style.configure(
        ".", background=p.background, foreground=p.text, font=font, bordercolor=p.border,
        lightcolor=p.background, darkcolor=p.background, troughcolor=p.background, focuscolor=p.background,
        selectbackground=p.accent, selectforeground=p.on_accent, insertcolor=p.text, fieldbackground=p.surface,
    )
    style.map(".", foreground=[("disabled", p.muted)])

    # Text
    style.configure("Muted.TLabel", foreground=p.muted)
    style.configure("Title.TLabel", font=(FONT_FAMILY, 13, "bold"))
    style.configure("Page.TLabel", font=(FONT_FAMILY, 18, "bold"))
    style.configure("Sidebar.TFrame", background=p.sidebar)
    style.configure("Brand.TLabel", background=p.sidebar, foreground=p.accent_text, font=(FONT_FAMILY, 15, "bold"))
    style.configure("SidebarMuted.TLabel", background=p.sidebar, foreground=p.muted)

    # Buttons
    _flat(style, "TButton", p.surface, padding=(14, 7), relief="flat", anchor="center")
    style.map("TButton",
              background=[("disabled", p.background), ("pressed", p.accent_soft), ("active", p.hover)],
              lightcolor=[("disabled", p.background), ("pressed", p.accent_soft), ("active", p.hover)],
              darkcolor=[("disabled", p.background), ("pressed", p.accent_soft), ("active", p.hover)],
              bordercolor=[("focus", p.accent)])
    _flat(style, "Accent.TButton", p.accent, foreground=p.on_accent, bordercolor=p.accent, font=bold, padding=(16, 9))
    style.map("Accent.TButton",
              background=[("disabled", p.border), ("pressed", p.accent_hover), ("active", p.accent_hover)],
              lightcolor=[("disabled", p.border), ("pressed", p.accent_hover), ("active", p.accent_hover)],
              darkcolor=[("disabled", p.border), ("pressed", p.accent_hover), ("active", p.accent_hover)],
              foreground=[("disabled", p.muted)],
              bordercolor=[("disabled", p.border), ("focus", p.text)])
    _flat(style, "TMenubutton", p.surface, padding=(14, 7), relief="flat", arrowcolor=p.text)
    style.map("TMenubutton",
              background=[("pressed", p.accent_soft), ("active", p.hover)],
              lightcolor=[("pressed", p.accent_soft), ("active", p.hover)],
              darkcolor=[("pressed", p.accent_soft), ("active", p.hover)],
              bordercolor=[("focus", p.accent)])

    # Side menu
    _flat(style, "Nav.TButton", p.sidebar, bordercolor=p.sidebar, foreground=p.text, anchor="w", padding=(12, 10))
    _flat(style, "NavActive.TButton", p.accent_soft, bordercolor=p.accent_soft, foreground=p.accent_text,
          font=bold, anchor="w", padding=(12, 10))
    for name, color in (("Nav.TButton", p.hover), ("NavActive.TButton", p.accent_soft)):
        style.map(name, background=[("active", color)], lightcolor=[("active", color)],
                  darkcolor=[("active", color)], bordercolor=[("focus", p.accent), ("active", color)])

    # Fields
    for name in ("TEntry", "TCombobox"):
        style.configure(name, fieldbackground=p.surface, foreground=p.text, bordercolor=p.border,
                        lightcolor=p.surface, darkcolor=p.surface, padding=5, arrowcolor=p.text, background=p.surface)
        style.map(name, bordercolor=[("focus", p.accent)], lightcolor=[("focus", p.accent)],
                  darkcolor=[("focus", p.accent)], fieldbackground=[("readonly", p.surface)],
                  foreground=[("readonly", p.text)], selectbackground=[("readonly", p.surface)],
                  selectforeground=[("readonly", p.text)], background=[("active", p.hover)])

    # Check boxes
    for name, color in (("TCheckbutton", p.text), ("Test.TCheckbutton", p.warning), ("Real.TCheckbutton", p.success)):
        style.configure(name, background=p.background, foreground=color, indicatorbackground=p.surface,
                        indicatorforeground=p.accent, upperbordercolor=p.border, lowerbordercolor=p.border,
                        indicatormargin=(0, 2, 6, 2), font=font if name == "TCheckbutton" else bold)
        style.map(name, background=[("active", p.background)], indicatorbackground=[("active", p.hover)])

    # Tables
    style.configure("Treeview", background=p.surface, fieldbackground=p.surface, foreground=p.text,
                    bordercolor=p.border, lightcolor=p.surface, darkcolor=p.surface, rowheight=row_height)
    style.map("Treeview", background=[("selected", p.accent)], foreground=[("selected", p.on_accent)])
    _flat(style, "Treeview.Heading", p.heading, foreground=p.muted, bordercolor=p.border, relief="flat",
          font=bold, padding=(8, 7))
    style.map("Treeview.Heading", background=[("active", p.hover)], lightcolor=[("active", p.hover)],
              darkcolor=[("active", p.hover)])

    # Scrollbars without arrows
    for orient, sticky in (("Vertical", "ns"), ("Horizontal", "ew")):
        name = f"{orient}.TScrollbar"
        style.layout(name, [(f"{orient}.Scrollbar.trough", {"sticky": sticky, "children": [
            (f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        style.configure(name, background=p.border, troughcolor=p.background, bordercolor=p.background,
                        lightcolor=p.border, darkcolor=p.border, arrowsize=10, gripcount=0, relief="flat")
        style.map(name, background=[("active", p.muted)], lightcolor=[("active", p.muted)], darkcolor=[("active", p.muted)])

    # Other widgets
    style.configure("Horizontal.TProgressbar", background=p.accent, troughcolor=p.border, bordercolor=p.border,
                    lightcolor=p.accent, darkcolor=p.accent)
    style.configure("TPanedwindow", background=p.background)
    style.configure("Sash", sashthickness=10, gripcount=0, background=p.background, lightcolor=p.background,
                    darkcolor=p.background, bordercolor=p.background)
    style.configure("TNotebook", background=p.background, bordercolor=p.border, tabmargins=(0, 0, 0, 0))
    _flat(style, "TNotebook.Tab", p.heading, foreground=p.muted, bordercolor=p.border, padding=(14, 7))
    style.map("TNotebook.Tab",
              background=[("selected", p.surface), ("active", p.hover)],
              lightcolor=[("selected", p.surface), ("active", p.hover)],
              foreground=[("selected", p.accent_text)],
              padding=[("selected", (14, 7))],
              expand=[("selected", (0, 0, 0, 0))])

    # The main pages are chosen from the side menu: no tabs and no frame.
    style.configure("Pages.TNotebook", background=p.background, borderwidth=0, bordercolor=p.background,
                    lightcolor=p.background, darkcolor=p.background, tabmargins=0)
    style.layout("Pages.TNotebook.Tab", [])
