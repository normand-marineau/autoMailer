"""
Validation Module
=================

This module defines routines for pre-flight checks and per-row validation
before sending emails.  It ensures that:

* A recipient column has been selected and exists in the data.
* All placeholders used in the subject and body exist in the header map.
* Each row contains valid and non-empty values for required fields.

Rows that fail validation are skipped with a human-readable reason.  The
pre-flight check returns a list of fatal errors that must be fixed
before any sending can occur.
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, List, Optional, Tuple

from .template_engine import find_placeholders


class PreflightError(Exception):
    """Raised when a blocking configuration error is detected."""


def preflight_check(
    headers: Dict[str, str],
    recipient_column: Optional[str],
    subject_template: str,
    body_template: str,
) -> List[str]:
    """Perform configuration checks before sending.

    :returns: A list of error messages.  If the list is non-empty, sending
        should not proceed.
    """
    errors: List[str] = []
    if not recipient_column:
        errors.append("Aucune colonne destinataire sélectionnée.")
    elif recipient_column not in headers:
        errors.append(
            f"La colonne destinataire '{recipient_column}' n'existe pas dans le fichier."
        )
    # Check that all placeholders used in the templates exist in the headers
    for template, label in [(subject_template, "sujet"), (body_template, "corps")]:
        missing = _missing_placeholders(template, headers)
        if missing:
            errors.append(
                f"Le {label} contient des variables inconnues: {', '.join(sorted(missing))}."
            )
    return errors


def _missing_placeholders(template: str, headers: Dict[str, str]) -> List[str]:
    """Return a list of placeholders that are not present in headers."""
    placeholders = find_placeholders(template)
    missing = [p for p in placeholders if p not in headers]
    return missing


def validate_row(
    row: Dict[str, str],
    required_keys: Iterable[str],
    recipient_key: str,
) -> Optional[str]:
    """Validate a single row.

    :param row: Normalised row dictionary keyed by placeholder keys.
    :param required_keys: Keys that must not be empty in addition to the recipient.
    :param recipient_key: Key for the recipient email address.
    :returns: None if the row is valid; otherwise a human-readable reason
        why it should be skipped.
    """
    # Check recipient presence
    recipient = row.get(recipient_key, "").strip()
    if not recipient:
        return "Le destinataire est vide ou invalide."
    # Basic email validation (very simple regex)
    if not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", recipient):
        return "Le destinataire n'est pas une adresse email valide."
    # Check additional required keys
    for key in required_keys:
        value = row.get(key, "").strip()
        if not value:
            return f"Le champ obligatoire '{key}' est vide."
    return None