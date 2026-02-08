"""
Parametrisation Tab Module
=========================

This module defines the UI for loading the data file, selecting the
recipient column, configuring the provider (Outlook or Gmail) and
running the email sending process.  It also displays the progress and
logs as the run progresses.  The design goal is to integrate all
operational controls into this single tab, replacing the multiple run
tabs used in the original script.

The current implementation provides placeholder widgets with no
functionality.  It will later interact with the core modules to perform
validation, run engine management and logging.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class ParamsTab(ttk.Frame):
    """A Tkinter frame containing parametrisation controls."""

    def __init__(self, parent: tk.Widget, *args, **kwargs) -> None:
        super().__init__(parent, *args, **kwargs)
        self._build_widgets()

    def _build_widgets(self) -> None:
        """Construct the widgets for the parametrisation tab."""
        # File loader
        file_frame = ttk.LabelFrame(self, text="Fichier de données")
        file_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        load_button = ttk.Button(file_frame, text="Choisir un fichier…")
        load_button.grid(row=0, column=0, padx=5, pady=5)
        self.filename_var = tk.StringVar()
        filename_entry = ttk.Entry(file_frame, textvariable=self.filename_var, state="readonly")
        filename_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        file_frame.columnconfigure(1, weight=1)

        # Recipient column selector (placeholder)
        recipient_frame = ttk.LabelFrame(self, text="Colonne destinataire")
        recipient_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        self.recipient_column_var = tk.StringVar()
        recipient_dropdown = ttk.Combobox(recipient_frame, textvariable=self.recipient_column_var, values=[])
        recipient_dropdown.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        recipient_frame.columnconfigure(0, weight=1)

        # Provider selection (Outlook or Gmail)
        provider_frame = ttk.LabelFrame(self, text="Fournisseur d'envoi")
        provider_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.provider_var = tk.StringVar(value="outlook")
        ttk.Radiobutton(provider_frame, text="Outlook", variable=self.provider_var, value="outlook").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(provider_frame, text="Gmail", variable=self.provider_var, value="gmail").grid(row=0, column=1, padx=5, pady=5)

        # Mode selection (draft, send now, send later) — Gmail supports only send now for now
        mode_frame = ttk.LabelFrame(self, text="Mode d'envoi")
        mode_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.mode_var = tk.StringVar(value="draft")
        ttk.Radiobutton(mode_frame, text="Brouillon", variable=self.mode_var, value="draft").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="Envoyer maintenant", variable=self.mode_var, value="send").grid(row=0, column=1, padx=5, pady=5)

        # Test run and throttle controls (placeholders)
        options_frame = ttk.LabelFrame(self, text="Options avancées")
        options_frame.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
        self.test_run_var = tk.IntVar(value=0)
        ttk.Checkbutton(options_frame, text="Exécuter un test (envoyer à N premières lignes)", variable=self.test_run_var).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.test_count_var = tk.IntVar(value=3)
        test_count_spin = ttk.Spinbox(options_frame, from_=1, to=1000, textvariable=self.test_count_var, width=5)
        test_count_spin.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        # Throttling: emails per minute
        ttk.Label(options_frame, text="Limite d'envoi (emails/min) :").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.throttle_var = tk.IntVar(value=60)
        throttle_spin = ttk.Spinbox(options_frame, from_=1, to=500, textvariable=self.throttle_var, width=5)
        throttle_spin.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        options_frame.columnconfigure(0, weight=1)

        # Start/Stop controls and log area
        run_frame = ttk.LabelFrame(self, text="Exécution")
        run_frame.grid(row=5, column=0, sticky="nsew", padx=5, pady=5)
        start_button = ttk.Button(run_frame, text="Démarrer", state="disabled")
        start_button.grid(row=0, column=0, padx=5, pady=5)
        pause_button = ttk.Button(run_frame, text="Pause", state="disabled")
        pause_button.grid(row=0, column=1, padx=5, pady=5)
        stop_button = ttk.Button(run_frame, text="Arrêter", state="disabled")
        stop_button.grid(row=0, column=2, padx=5, pady=5)

        # Log display
        self.log_text = tk.Text(run_frame, wrap="none", height=10, state="disabled")
        self.log_text.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        # Configure run_frame to expand log
        run_frame.rowconfigure(1, weight=1)
        run_frame.columnconfigure(0, weight=1)
        run_frame.columnconfigure(1, weight=1)
        run_frame.columnconfigure(2, weight=1)

        # Configure overall grid weights for this tab
        self.columnconfigure(0, weight=1)
        self.rowconfigure(5, weight=1)