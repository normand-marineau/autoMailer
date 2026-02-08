"""
Logging Module
==============

Provides classes for writing send logs and skipped rows to CSV files.  The
`LogWriter` class exposes a simple `write_result` method that can be
called from the run engine to record the status of each row.  Logs are
written to two files: one capturing all attempts (with statuses
"sent", "drafted", "scheduled", "skipped", "error") and another
containing only skipped rows along with the reason for skipping.

The logger uses Python's built-in CSV module and supports streaming
entries to the UI via an optional callback.
"""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .models import RowResult


class LogWriter:
    """Write results to CSV log files and optionally notify a callback."""

    def __init__(self, output_dir: Optional[Path] = None, callback: Optional[Callable[[RowResult], None]] = None) -> None:
        # Determine output directory
        self.output_dir = output_dir or Path.cwd()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = self.output_dir / f"send_log_{timestamp}.csv"
        self.skipped_path = self.output_dir / f"skipped_rows_{timestamp}.csv"
        self.callback = callback
        # Write headers immediately
        with self.log_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["row_number", "recipient", "status", "reason"])
        with self.skipped_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["row_number", "recipient", "reason"])

    def write_result(self, result: RowResult) -> None:
        """Append a row result to the log files and trigger callback."""
        # Write to main log
        with self.log_path.open("a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([result.row_number, result.recipient, result.status, result.reason])
        # Write to skipped log if applicable
        if result.status in {"skipped", "error"}:
            with self.skipped_path.open("a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([result.row_number, result.recipient, result.reason])
        # Emit callback to UI
        if self.callback is not None:
            self.callback(result)