from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Dict, List

EMAIL_RE = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.IGNORECASE)
PLACEHOLDER_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")

def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def normalize_header(header: str) -> str:
    """
    Convert a raw header label into a safe placeholder key.
    Example:
      "Prénom (étudiant)" -> "prenom_etudiant"
    """
    h = (header or "").strip()
    h = unicodedata.normalize("NFKD", h)
    h = "".join(ch for ch in h if not unicodedata.combining(ch))
    h = h.lower()
    h = re.sub(r"[^a-z0-9]+", "_", h)
    h = re.sub(r"_+", "_", h).strip("_")
    if not h:
        h = "col"
    return h

def placeholders_used(subject: str, body: str) -> List[str]:
    used = set()
    for m in PLACEHOLDER_RE.finditer(subject or ""):
        used.add(m.group(1))
    for m in PLACEHOLDER_RE.finditer(body or ""):
        used.add(m.group(1))
    return sorted(used)

def render_template(tpl: str, row: Dict[str, str]) -> str:
    if not tpl:
        return ""
    def repl(m: re.Match) -> str:
        key = m.group(1)
        return str(row.get(key, ""))
    return PLACEHOLDER_RE.sub(repl, tpl)
