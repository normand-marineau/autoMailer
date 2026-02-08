"""
Providers package
=================

This package contains implementations of the provider interface for
different email sending backends.  Each provider class implements
methods to send an email immediately, create drafts if supported, or
schedule sending if supported.  Providers expose a capability object
to the UI so that unsupported features can be disabled.

Currently implemented providers:

* OutlookProvider – uses win32com to control Outlook on Windows.
* GmailProvider – uses the Gmail API and OAuth2; supports send now only.
"""

from __future__ import annotations

__all__ = [
    "base_provider",
    "outlook_provider",
    "gmail_provider",
]