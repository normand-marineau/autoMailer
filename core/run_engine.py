"""
Run Engine Module
=================

Orchestrates the process of sending emails to a batch of recipients based on
the user-selected configuration.  The run engine interacts with a
provider to actually send or create drafts, applies throttling, supports
pausing and stopping mid-run, and records results via the logger.

This module exposes a `RunEngine` class with methods to start the run,
pause/resume and stop.  It is intentionally decoupled from the UI so that
it can be tested in isolation and potentially reused in other contexts.

In this skeleton implementation, the actual sending logic is not
implemented.  Instead, placeholders illustrate where calls to the
provider should occur.
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from typing import Callable, Dict, Iterable, List, Optional

from .logger import LogWriter
from .models import MessageSpec, ProviderCaps, RowResult
from .validator import validate_row
from .template_engine import render_template


class RunEngine:
    """Controls the batch sending of emails using a provider."""

    def __init__(
        self,
        *,
        provider: "EmailProvider",  # type: ignore
        message_spec: MessageSpec,
        rows: List[Dict[str, str]],
        recipient_key: str,
        required_keys: Iterable[str],
        throttle: int,
        test_count: Optional[int] = None,
        log_writer: Optional[LogWriter] = None,
    ) -> None:
        self.provider = provider
        self.message_spec = message_spec
        self.rows = rows
        self.recipient_key = recipient_key
        self.required_keys = list(required_keys)
        self.throttle = max(1, throttle)
        self.test_count = test_count
        self.log_writer = log_writer or LogWriter()
        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()  # Initially not paused

    def start(self) -> None:
        """Start the run in a background thread."""
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_loop)
        self._thread.start()

    def pause(self) -> None:
        """Pause the sending process."""
        self._pause_event.clear()

    def resume(self) -> None:
        """Resume sending after a pause."""
        self._pause_event.set()

    def stop(self) -> None:
        """Signal the run to stop after the current message."""
        self._stop_event.set()
        self._pause_event.set()  # Unpause to allow thread to exit if paused
        if self._thread:
            self._thread.join()

    def _run_loop(self) -> None:
        """Internal loop iterating through rows and sending messages."""
        count = 0
        for row_number, row in enumerate(self.rows, start=1):
            if self.test_count is not None and count >= self.test_count:
                break
            if self._stop_event.is_set():
                break
            # Pause if requested
            self._pause_event.wait()
            # Validate row
            error = validate_row(row, self.required_keys, self.recipient_key)
            if error:
                # Row skipped
                result = RowResult(
                    row_number=row_number,
                    recipient=row.get(self.recipient_key, ""),
                    status="skipped",
                    reason=error,
                )
                self.log_writer.write_result(result)
                continue
            # Render subject and body
            try:
                subject = render_template(self.message_spec.subject, row)
                body = render_template(self.message_spec.body, row)
            except KeyError as ke:
                result = RowResult(
                    row_number=row_number,
                    recipient=row.get(self.recipient_key, ""),
                    status="skipped",
                    reason=str(ke),
                )
                self.log_writer.write_result(result)
                continue
            # Send via provider
            try:
                if self.message_spec.mode == "schedule":
                    send_time = self.message_spec.schedule_time
                    self.provider.schedule_send(row[self.recipient_key], subject, body, send_time)  # type: ignore[call-arg]
                    status = "scheduled"
                else:
                    self.provider.send_now(row[self.recipient_key], subject, body)
                    status = "sent"
            except Exception as exc:
                result = RowResult(
                    row_number=row_number,
                    recipient=row.get(self.recipient_key, ""),
                    status="error",
                    reason=str(exc),
                )
                self.log_writer.write_result(result)
                # decide whether to stop on provider error or continue
                continue
            # Log success
            result = RowResult(
                row_number=row_number,
                recipient=row[self.recipient_key],
                status=status,
                reason="",
            )
            self.log_writer.write_result(result)
            count += 1
            # Throttle: wait (60/throttle) seconds between sends
            time_per_email = 60.0 / self.throttle
            time.sleep(time_per_email)