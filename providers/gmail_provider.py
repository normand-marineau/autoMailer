"""
Gmail Provider
==============

Implements the EmailProvider interface using the Gmail API via OAuth2.
This provider currently supports sending emails immediately; draft mode
and scheduling will be added in a later version.  To use this provider,
the user must set up a Google Cloud project, enable the Gmail API, and
obtain OAuth client credentials.  The implementation should handle the
OAuth flow and token persistence, then build and send MIME messages via
the Gmail API.

This is a stub implementation that defines the interface only.  It
includes comments to indicate where the OAuth and API calls should be
placed.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from ..core.models import ProviderCaps
from .base_provider import EmailProvider


class GmailProvider(EmailProvider):
    """Email provider for Gmail using the Gmail API and OAuth2."""

    def __init__(self) -> None:
        # Gmail v2 supports send only; drafts and scheduling are not yet implemented
        self.caps = ProviderCaps(supports_draft=False, supports_schedule=False)
        # Placeholder for authorised Gmail service object
        self.service: Any = None
        # TODO: perform OAuth2 authentication and store the credentials
        # You would use google-auth and google-api-python-client libraries here.
        raise NotImplementedError(
            "GmailProvider OAuth flow is not yet implemented."
        )

    def send_now(self, recipient: str, subject: str, body: str) -> None:
        # TODO: use the authorised Gmail service to send an email
        # For example, build a MIME message and call
        # `service.users().messages().send(userId='me', body=message).execute()`
        raise NotImplementedError("GmailProvider send_now not implemented")

    def create_draft(self, recipient: str, subject: str, body: str) -> None:
        raise NotImplementedError("Draft mode is not supported for GmailProvider in v2")

    def schedule_send(self, recipient: str, subject: str, body: str, send_time: datetime) -> None:
        raise NotImplementedError("Scheduled send is not supported for GmailProvider in v2")