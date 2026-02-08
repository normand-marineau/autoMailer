"""
Base Provider Interface
=======================

Defines an abstract interface for email providers.  A provider must
implement methods to send an email immediately, optionally create a
draft, and optionally schedule a send.  Each provider also exposes
capabilities describing which of these features are available so that
the UI can disable unsupported options.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any

from ..core.models import ProviderCaps


class EmailProvider(ABC):
    """Abstract base class for email providers."""

    caps: ProviderCaps

    @abstractmethod
    def send_now(self, recipient: str, subject: str, body: str) -> Any:
        """Send an email immediately."""
        raise NotImplementedError

    def create_draft(self, recipient: str, subject: str, body: str) -> Any:
        """Create a draft email.  Default implementation is unsupported."""
        raise NotImplementedError("Draft mode not supported by this provider")

    def schedule_send(self, recipient: str, subject: str, body: str, send_time: datetime) -> Any:
        """Schedule an email to be sent later.  Default implementation is unsupported."""
        raise NotImplementedError("Scheduled send not supported by this provider")