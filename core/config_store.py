"""
Configuration Storage Module
===========================

Provides a simple mechanism for persisting and retrieving user
configuration across sessions.  Settings such as the last-used
provider, throttle value, and selected recipient column can be saved
to a JSON file in the user's application data directory.  This module
abstracts away the file system paths and JSON handling.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional


DEFAULT_FILENAME = "ulaval_mailer_config.json"


def get_config_path() -> Path:
    """Return the path to the configuration file.

    The configuration file is stored in the user's home directory under
    `.ulaval_mailer`.  If that directory does not exist, it is created.
    """
    home = Path.home()
    config_dir = home / ".ulaval_mailer"
    config_dir.mkdir(exist_ok=True)
    return config_dir / DEFAULT_FILENAME


def load_config(path: Optional[Path] = None) -> Dict[str, Any]:
    """Load configuration data from disk.

    :param path: Override the default configuration file location.
    :returns: A dictionary of settings or an empty dict if no file exists.
    """
    config_file = path or get_config_path()
    if config_file.exists():
        with config_file.open("r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {}


def save_config(data: Dict[str, Any], path: Optional[Path] = None) -> None:
    """Persist configuration data to disk."""
    config_file = path or get_config_path()
    with config_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)