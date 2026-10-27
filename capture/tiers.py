"""Fast-tier selection, shared by capture and retier.

Rule: rows are visited in capture order; a row joins the fast tier when its output shape
(digits replaced by 9, letters by a) has not been seen for its (format, width) group and the
group has fewer than ``cap`` fast rows. ``cap`` 0 makes a suite full-tier only.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import replace

from capture.schema import Row

_DIGITS = re.compile(r"[0-9]")
_LETTERS = re.compile(r"[A-Za-z]")


def output_class(text: str) -> str:
    """Shape of an output: digits become 9, letters become a."""
    return _LETTERS.sub("a", _DIGITS.sub("9", text))


def rule_description(cap: int) -> str:
    """Text for the TSV header."""
    if cap == 0:
        return "full tier only"
    return f"first row per output shape per (format, width), at most {cap} per group"


def assign_tiers(rows: list[Row], cap: int) -> list[Row]:
    """Return rows with the ``tier`` column reassigned."""
    seen: set[tuple[str, str, str]] = set()
    counts: Counter[tuple[str, str]] = Counter()
    result: list[Row] = []
    for row in rows:
        group = (row.format, row.width)
        key = (*group, output_class(row.expected_text))
        fast = cap > 0 and key not in seen and counts[group] < cap
        seen.add(key)
        if fast:
            counts[group] += 1
        result.append(replace(row, tier="fast" if fast else "full"))
    return result
