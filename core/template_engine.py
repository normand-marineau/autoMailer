"""
Template Engine Module
======================

This module provides a very simple template engine for substituting
placeholders in a subject or body template with values from a row of
data.  Placeholders are denoted using curly braces around the
normalised header name, for example `{prenom}` or `{dateexam}`.  Only
exact matches are replaced; no expression syntax is supported.

If a placeholder references a key that does not exist in the row, the
function will either raise an error or leave the placeholder untouched
depending on the chosen behaviour.  By default missing keys result in a
KeyError.
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, Set


PLACEHOLDER_PATTERN = re.compile(r"\{([^{}]+)\}")


def find_placeholders(template: str) -> Set[str]:
    """Return the set of unique placeholder keys found in the template."""
    return set(PLACEHOLDER_PATTERN.findall(template))


def render_template(template: str, row: Dict[str, str], *, strict: bool = True) -> str:
    """Render a template by substituting placeholders with row values.

    :param template: The subject or body template containing placeholders.
    :param row: A dictionary mapping normalised header keys to their
        corresponding values for a single row.
    :param strict: If True, raise KeyError when a placeholder is not found
        in the row.  If False, leave unknown placeholders untouched.
    :returns: The rendered string with placeholders replaced.
    :raises KeyError: If `strict` is True and a placeholder key is missing.
    """

    def replace(match: re.Match) -> str:
        key = match.group(1)
        if key in row:
            return row[key]
        if strict:
            raise KeyError(f"Missing value for placeholder '{key}'")
        # Leave unknown placeholder unchanged
        return match.group(0)

    return PLACEHOLDER_PATTERN.sub(replace, template)