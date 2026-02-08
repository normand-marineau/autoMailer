"""
UI package for Ulaval Mailer v2.

This package contains all Tkinter-related modules used to build the user
interface.  It is organised into separate modules for each tab and shared
state.  The intent is to keep UI concerns isolated from core logic and
providers.
"""

from __future__ import annotations

__all__ = [
    "tab_guide",
    "tab_message",
    "tab_params",
    "ui_state",
]