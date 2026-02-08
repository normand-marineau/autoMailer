"""
Outlook Provider
================

Implements the EmailProvider interface using Microsoft Outlook via
COM automation.  This provider requires Windows and Outlook to be
installed.  It supports creating drafts, sending immediately and
scheduling via Outlook's DeferredDeliveryTime property.

At the moment, this implementation is a skeleton.  In a full
implementation you would import and use the `win32com.client` module to
create and configure Outlook mail items.  Exceptions should be caught
and re-raised as appropriate for the run engine.
"""

from __future__ import annotations

from datetime import datetime

try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # type: ignore

from ..core.models import ProviderCaps
from .base_provider import EmailProvider


class OutlookProvider(EmailProvider):
    """Email provider for Microsoft Outlook."""

    def __init__(self) -> None:
        # Outlook supports draft and schedule
        self.caps = ProviderCaps(supports_draft=True, supports_schedule=True)
        if win32com is None:
            raise RuntimeError(
                "win32com is required for OutlookProvider; please install pywin32"
            )
        # Defer connecting to Outlook until needed
        self._outlook = None

    def _ensure_outlook(self) -> None:
        if self._outlook is None:
            # Acquire the Outlook application COM object
            self._outlook = win32com.client.Dispatch("Outlook.Application")

    def send_now(self, recipient: str, subject: str, body: str) -> None:
        self._ensure_outlook()
        mail = self._outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Send()

    def create_draft(self, recipient: str, subject: str, body: str) -> None:
        self._ensure_outlook()
        mail = self._outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Save()  # Save to Drafts folder

    def schedule_send(self, recipient: str, subject: str, body: str, send_time: datetime) -> None:
        self._ensure_outlook()
        mail = self._outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        # Set DeferredDeliveryTime to schedule send
        mail.DeferredDeliveryTime = send_time
        mail.Send()