"""
Data Models Module
==================

Defines simple data classes used throughout the core and provider
modules.  These classes facilitate type checking and improve
readability.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class MessageSpec:
    """Represents the message template and sending configuration."""

    subject: str
    body: str
    mode: str  # 'draft', 'send' or 'schedule'
    schedule_time: Optional[datetime] = None


@dataclass
class RowResult:
    """Represents the result of processing a single row."""

    row_number: int
    recipient: str
    status: str  # 'sent', 'drafted', 'scheduled', 'skipped', 'error'
    reason: str


@dataclass
class ProviderCaps:
    """Describes the capabilities of a provider."""

    supports_draft: bool
    supports_schedule: bool