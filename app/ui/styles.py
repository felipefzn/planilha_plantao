from __future__ import annotations

from tkinter import ttk

import customtkinter as ctk


def setup_theme(root, tema: dict) -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    cor_primaria = tema.get("cor_primaria", "#1F5AA6")
    cor_fundo = tema.get("cor_fundo", "#F3F7FC")
    cor_card = tema.get("cor_card", "#FFFFFF")
    cor_texto = tema.get("cor_texto", "#183153")
    cor_texto_secundario = tema.get("cor_texto_secundario", "#5B6B82")

    style.configure(
        "Corporate.Treeview",
        background=cor_card,
        fieldbackground=cor_card,
        foreground=cor_texto,
        borderwidth=0,
        rowheight=31,
        font=("Segoe UI", 10),
    )
    style.map(
        "Corporate.Treeview",
        background=[("selected", "#D9E8FB")],
        foreground=[("selected", cor_texto)],
    )
    style.configure(
        "Corporate.Treeview.Heading",
        background=cor_primaria,
        foreground="#FFFFFF",
        relief="flat",
        borderwidth=0,
        font=("Segoe UI Semibold", 10),
        padding=(10, 8),
    )
    style.map("Corporate.Treeview.Heading", background=[("active", cor_primaria)])
    style.configure(
        "TScrollbar",
        gripcount=0,
        background="#CFDCF0",
        troughcolor=cor_fundo,
        bordercolor=cor_fundo,
        arrowcolor=cor_texto_secundario,
    )


def card_kwargs(tema: dict) -> dict:
    return {
        "fg_color": tema.get("cor_card", "#FFFFFF"),
        "corner_radius": 18,
        "border_width": 1,
        "border_color": "#D8E4F4",
    }
