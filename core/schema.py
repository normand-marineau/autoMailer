"""
Schema Detection Module
=======================

This module is responsible for analysing the raw headers from a data file
and determining how they map to placeholders used in templates.  It
normalises header names into safe keys (e.g. removing spaces and accents)
while preserving the original labels for display in the UI.  It also
identifies special columns such as the recipient email column and
optional CC/BCC fields.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Dict, List, Optional, Tuple


def normalise_header(header: str) -> str:
    """Convert a raw header into a safe placeholder key.

    The function removes leading/trailing whitespace, converts spaces to
    underscores, strips accents and non-alphanumeric characters and
    lowercases the result.  This ensures that placeholders like
    `{prenom}` are easy to type and consistent regardless of the source
    file.

    :param header: Raw header from the CSV/XLSX file.
    :returns: Normalised header key.
    """
    # Strip whitespace
    name = header.strip()
    # Remove accents
    name = unicodedata.normalize('NFKD', name)
    name = ''.join(c for c in name if not unicodedata.combining(c))
    # Replace spaces and hyphens with underscores
    name = re.sub(r"[\s\-]+", "_", name)
    # Remove non-alphanumeric/underscore characters
    name = re.sub(r"[^a-zA-Z0-9_]", "", name)
    # Lowercase
    return name.lower()


def build_schema(headers: List[str]) -> Tuple[Dict[str, str], List[str]]:
    """Build a mapping of normalised headers and detect recipient candidates.

    Given a list of raw header names, returns a tuple `(normalised, candidates)`
    where `normalised` maps the safe placeholder key to the original label
    and `candidates` is a list of normalised keys that are likely to
    correspond to the recipient email column.  Candidate detection is
    heuristic-based, matching common French and English words such as
    'courriel', 'email', 'to' and 'adresse'.

    :param headers: List of raw header strings.
    :returns: A tuple of (normalised_header_map, recipient_candidates).
    """
    norm_map: Dict[str, str] = {}
    candidates: List[str] = []
    for raw in headers:
        norm = normalise_header(raw)
        # Skip empty headers
        if not norm:
            continue
        norm_map[norm] = raw
        # Identify potential recipient columns based on keywords
        lowered = raw.lower()
        if any(
            keyword in lowered
            for keyword in ["courriel", "email", "e-mail", "to", "destinataire", "adresse"]
        ):
            candidates.append(norm)
    return norm_map, candidates