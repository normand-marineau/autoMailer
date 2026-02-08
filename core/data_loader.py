"""
Data Loader Module
==================

This module defines functions to load data from CSV and XLSX files into
a unified internal representation (a list of dictionaries).  It also
handles simple type inference and whitespace trimming.  The result of
`load_data` should contain a list of rows where each row is a mapping
from the raw header name to the string value found in the file.

The module does not perform header normalisation or special column
detection.  Those responsibilities belong to the `schema` module.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Tuple, Union

try:
    import openpyxl  # type: ignore
except ImportError:
    openpyxl = None  # Optional dependency


def load_data(file_path: Union[str, Path]) -> Tuple[List[str], List[Dict[str, str]]]:
    """Load data from a CSV or XLSX file.

    Returns a tuple `(headers, rows)` where `headers` is a list of the raw
    column names and `rows` is a list of dictionaries mapping header names
    to their string values.  Leading and trailing whitespace is stripped
    from cell values.  If the file extension is not recognised or the
    optional `openpyxl` dependency is missing for XLSX files, an
    exception is raised.

    :param file_path: Path to the file on disk.
    :raises ValueError: If the file has an unsupported extension.
    :raises RuntimeError: If reading an XLSX file without openpyxl.
    :returns: A tuple of `(headers, rows)`.
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    if ext == ".csv":
        return _read_csv(path)
    if ext in {".xls", ".xlsx"}:
        return _read_xlsx(path)
    raise ValueError(f"Unsupported file extension: {ext}")


def _read_csv(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    """Read a CSV file into headers and rows."""
    rows: List[Dict[str, str]] = []
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        for row in reader:
            # Ensure all values are strings and strip whitespace
            cleaned = {h: (row.get(h, "") or "").strip() for h in headers}
            rows.append(cleaned)
    return headers, rows


def _read_xlsx(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    """Read an XLSX file into headers and rows using openpyxl."""
    if openpyxl is None:
        raise RuntimeError(
            "openpyxl is required to read Excel files; please install it via pip"
        )
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    headers: List[str] = []
    rows: List[Dict[str, str]] = []
    for i, row_cells in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0:
            # header row
            headers = [str(cell).strip() if cell is not None else "" for cell in row_cells]
            continue
        row_dict: Dict[str, str] = {}
        for header, cell in zip(headers, row_cells):
            value = "" if cell is None else str(cell).strip()
            row_dict[header] = value
        rows.append(row_dict)
    return headers, rows