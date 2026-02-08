"""
Core package for Ulaval Mailer.

The core package contains all business logic that is independent of the
user interface and the specific provider implementations.  It deals with
reading input data, interpreting the schema, validating inputs, running
send loops, and logging results.  The UI should interact with this
package via clean, well-defined functions and classes.
"""

from __future__ import annotations

__all__ = [
    "data_loader",
    "schema",
    "template_engine",
    "validator",
    "run_engine",
    "logger",
    "config_store",
    "models",
]