from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .text_utils import EMAIL_RE, normalize_header

try:
    import openpyxl  # type: ignore
except Exception:
    openpyxl = None

def sniff_csv_dialect(sample: str) -> csv.Dialect:
    try:
        return csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t"])
    except Exception:
        class _D(csv.Dialect):
            delimiter = ","
            quotechar = '"'
            escapechar = None
            doublequote = True
            skipinitialspace = False
            lineterminator = "\r\n"
            quoting = csv.QUOTE_MINIMAL
        return _D()

def read_csv(path: Path) -> Tuple[List[str], List[List[str]]]:
    raw = None
    for enc in ("utf-8-sig", "utf-8", "latin1"):
        try:
            raw = path.read_text(encoding=enc)
            break
        except Exception:
            continue
    if raw is None:
        raise RuntimeError("Impossible de lire le fichier CSV (encodage).")

    sample = raw[:4096]
    dialect = sniff_csv_dialect(sample)

    rows: List[List[str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=dialect.delimiter)
        for r in reader:
            rows.append([c.strip() if isinstance(c, str) else "" for c in r])

    if not rows:
        return [], []

    headers = rows[0]
    data = rows[1:]
    return headers, data

def read_xlsx(path: Path) -> Tuple[List[str], List[List[str]]]:
    if openpyxl is None:
        raise RuntimeError("openpyxl n'est pas installé. (pip install openpyxl)")

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows: List[List[str]] = []
    for row in ws.iter_rows(values_only=True):
        all_rows.append([("" if v is None else str(v)).strip() for v in row])

    if not all_rows:
        return [], []
    headers = all_rows[0]
    data = all_rows[1:]
    return headers, data

@dataclass
class LoadedData:
    path: Path
    raw_headers: List[str]
    key_to_header: Dict[str, str]
    header_to_key: Dict[str, str]
    rows: List[Dict[str, str]]  # keys -> values

    @property
    def nrows(self) -> int:
        return len(self.rows)

def build_key_maps(headers: List[str]) -> Tuple[Dict[str, str], Dict[str, str]]:
    key_to_header: Dict[str, str] = {}
    header_to_key: Dict[str, str] = {}
    used: Dict[str, int] = {}

    for h in headers:
        base = normalize_header(h)
        k = base
        if k in used:
            used[base] = used.get(base, 1) + 1
            k = f"{base}_{used[base]}"
        else:
            used[k] = 1
        key_to_header[k] = h
        header_to_key[h] = k
    return key_to_header, header_to_key

def rows_to_dicts(headers: List[str], data: List[List[str]], header_to_key: Dict[str, str]) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    for r in data:
        row_map: Dict[str, str] = {}
        for i, h in enumerate(headers):
            k = header_to_key.get(h, normalize_header(h))
            v = r[i].strip() if i < len(r) and isinstance(r[i], str) else ("" if i < len(r) else "")
            row_map[k] = v
        out.append(row_map)
    return out

def detect_recipient_key(key_to_header: Dict[str, str], rows: List[Dict[str, str]]) -> Optional[str]:
    priority_substrings = ["courriel", "email", "e_mail", "mail", "adresse", "to", "destin"]
    candidates = []
    for k, h in key_to_header.items():
        hl = (h or "").lower()
        if any(s in hl for s in priority_substrings):
            candidates.append(k)
    if candidates:
        return candidates[0]

    best_k = None
    best_score = 0
    for k in key_to_header.keys():
        score = 0
        for row in rows[:200]:
            val = (row.get(k, "") or "").strip()
            if val and EMAIL_RE.match(val):
                score += 1
        if score > best_score:
            best_score = score
            best_k = k
    return best_k if best_score > 0 else None
