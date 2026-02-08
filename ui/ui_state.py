"""
Shared UI State
===============

This module defines a simple class used to store state shared between the
different tabs of the Ulaval Mailer UI.  Examples of shared state include
the loaded data file path, the detected headers, the selected recipient
column, the current subject and body template, and the list of available
variables.

It is designed as a lightweight container rather than a business logic
handler.  Business logic should live in the core modules, while this
class can be mutated by the UI when the user interacts with widgets.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class UIState:
    """Holds shared data across the UI tabs."""

    # Path to the loaded data file (CSV or XLSX)
    filename: Optional[str] = None
    # Normalised headers as placeholder keys and their original labels
    headers: Dict[str, str] = field(default_factory=dict)
    # Candidate recipient columns detected by the schema module
    recipient_candidates: List[str] = field(default_factory=list)
    # Currently selected recipient column key
    recipient_column: Optional[str] = None
    # Subject and body templates
    subject_template: str = ""
    body_template: str = ""
    # Available variable keys (same as headers keys, but may be sorted)
    variables: List[str] = field(default_factory=list)

    def update_templates(self, subject: str, body: str) -> None:
        """Update the stored subject and body templates."""
        self.subject_template = subject
        self.body_template = body