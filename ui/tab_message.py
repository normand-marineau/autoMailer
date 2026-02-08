"""
Message Tab Module
==================

Defines the UI for composing the email subject and body.  The tab should
provide an editor for the subject, a multiline text box for the body,
a list of available variables (detected from the data file) and a preview
pane showing the rendered message for a selected row.  Clicking on a
variable in the list should insert it into the editor at the cursor
position.

This skeleton implementation creates placeholders for those widgets.  The
actual logic for populating available variables and rendering previews will
reside in the core modules and the UI state.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class MessageTab(ttk.Frame):
    """A Tkinter frame containing fields for the email message."""

    def __init__(self, parent: tk.Widget, *args, **kwargs) -> None:
        super().__init__(parent, *args, **kwargs)
        self._build_widgets()

    def _build_widgets(self) -> None:
        """Construct the widgets for the message tab."""
        # Subject entry
        subject_label = ttk.Label(self, text="Sujet :")
        subject_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.subject_var = tk.StringVar()
        subject_entry = ttk.Entry(self, textvariable=self.subject_var)
        subject_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # Body editor
        body_label = ttk.Label(self, text="Message :")
        body_label.grid(row=1, column=0, sticky="nw", padx=5, pady=5)
        self.body_text = tk.Text(self, wrap="word", height=15)
        self.body_text.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)

        # Available variables list (placeholder)
        variables_label = ttk.Label(self, text="Variables disponibles :")
        variables_label.grid(row=2, column=0, sticky="nw", padx=5, pady=5)
        self.variables_list = tk.Listbox(self, height=6)
        self.variables_list.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.variables_list.insert(
            tk.END,
            "(aucune variable chargée, importer un fichier dans Paramétrisation)",
        )

        # Preview area (placeholder)
        preview_label = ttk.Label(self, text="Aperçu :")
        preview_label.grid(row=3, column=0, sticky="nw", padx=5, pady=5)
        self.preview_text = tk.Text(self, wrap="word", height=10, state="disabled")
        self.preview_text.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)

        # Configure grid weights for resizing
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(3, weight=1)