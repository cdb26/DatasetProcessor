"""
ui/widgets.py
Reusable themed widget factories.
"""

import customtkinter as ctk
from ui.theme import *


def make_label(parent, text, font=FONT_LABEL, color=TEXT_MUTED, **kw):
    return ctk.CTkLabel(parent, text=text, font=font, text_color=color, **kw)


def make_value_label(parent, text="--", font=FONT_VALUE, color=ACCENT, **kw):
    return ctk.CTkLabel(parent, text=text, font=font, text_color=color, **kw)


def make_entry(parent, textvariable, width=130, **kw):
    return ctk.CTkEntry(
        parent,
        textvariable=textvariable,
        width=width,
        height=ENTRY_HEIGHT,
        corner_radius=CORNER_RADIUS,
        fg_color=BG_SURFACE,
        border_color=BORDER,
        border_width=1,
        text_color=TEXT_PRIMARY,
        font=FONT_MONO,
        **kw,
    )


def make_option_menu(parent, values, variable, width=120, **kw):
    return ctk.CTkOptionMenu(
        parent,
        values=values,
        variable=variable,
        width=width,
        height=ENTRY_HEIGHT,
        corner_radius=CORNER_RADIUS,
        fg_color=BG_SURFACE,
        button_color=BG_SURFACE,
        button_hover_color=ACCENT_DIM,
        dropdown_fg_color=BG_PANEL,
        dropdown_hover_color=ACCENT_DIM,
        text_color=TEXT_PRIMARY,
        font=FONT_MONO,
        **kw,
    )


def make_button(parent, text, command=None, width=160, color=ACCENT, hover=ACCENT_DIM, text_color="#000000", **kw):
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        height=BUTTON_HEIGHT,
        corner_radius=CORNER_RADIUS,
        fg_color=color,
        hover_color=hover,
        text_color=text_color,
        font=FONT_LABEL,
        **kw,
    )


def make_card(parent, **kw):
    return ctk.CTkFrame(
        parent,
        fg_color=BG_PANEL,
        corner_radius=CORNER_RADIUS,
        border_width=1,
        border_color=BORDER,
        **kw,
    )


def section_title(parent, text):
    return ctk.CTkLabel(
        parent,
        text=text,
        font=FONT_TITLE,
        text_color=TEXT_MUTED,
        anchor="w",
    )
