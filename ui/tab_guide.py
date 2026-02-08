"""
Guide Tab Module
================

This module defines a simple Tkinter frame for the Guide tab.  The guide
tab provides users with instructions on how to prepare their data file,
compose messages and configure providers.  It should include sample CSV
and XLSX file formats, placeholder syntax explanations and common
troubleshooting tips.

In this skeleton implementation, the tab only displays a placeholder
message.  The actual content should be written in French to align with
the user's language preference.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class GuideTab(ttk.Frame):
    """A Tkinter frame containing the guide content."""

    def __init__(self, parent: tk.Widget, *args, **kwargs) -> None:
        super().__init__(parent, *args, **kwargs)
        self._build_widgets()

    def _build_widgets(self) -> None:
        """Construct the widgets for the guide tab.

        In the full implementation, this method will populate the tab with
        headings, paragraphs, lists and sample file images.  For now it
        displays a placeholder label.
        """
        placeholder = ttk.Label(
            self,
            text=(
                "Onglet Guide\n"
                "Ici vous trouverez des instructions sur la préparation de vos fichiers,\n"
                "l'utilisation des variables et la configuration de l'application."
            ),
            justify="left",
        )
        placeholder.pack(fill="both", expand=True, padx=10, pady=10)