"""Measure functions and the column model for the width decision (``ARCHITECTURE.md`` §8).

Character mode measures a text by its length and the budget is the width in characters. Excel
mode measures a text as the floored sum of glyph advances in points at the cell's font size and
the budget is the column's text area in points: Excel for Mac stores a column of ``w`` units as
``round(MDW * w + 5)`` points, MDW being the maximum digit width of the Normal font (5 points for
Calibri 9 and 10, 6 for 11 and 12, 7 for 14, 8 for 16), and compares text against those points
less the 5 of padding. Both modes feed
the same seam in ``_assemble``: fit when ``measure(text) <= budget``, ``#`` and ``*`` counts by
dividing the budget by the glyph's measure (CAPTURING.md section 2 items 24 and 25).
"""

from __future__ import annotations

import math
from collections.abc import Callable
from functools import lru_cache

from xlformat._glyphs_calibri import CALIBRI

Measure = Callable[[str], float]
"""Width of a text in the caller's units."""

FontSpec = str | tuple[str, float]
"""A font name, or a name and a point size; the name alone means 11 points."""

DEFAULT_FONT: tuple[str, float] = ("Calibri", 11)
"""openpyxl's default workbook font; sheetpeek passes the workbook's Normal style."""

_FONTS: dict[str, dict[int, dict[str, float]]] = {"calibri": CALIBRI}
_COLUMN_PADDING = 5.0  # points Excel adds to MDW * width for the column, not available to text


def character_measure(text: str) -> float:
    """Character mode: one unit per character."""
    return float(len(text))


def general_cost(text: str) -> float:
    """What ``General`` digit reduction counts against the column width in units.

    Excel reduces General by characters, not by glyph widths: one unit per character and half
    a unit for a leading minus sign (``-12345678`` fits width 8.5, ``-1.2E+07`` needs 8, while
    the exponent's minus is a full unit: ``-2E-07`` needs 6; CAPTURING.md section 2 item 25). In
    character mode this equals ``len`` at whole widths.
    """
    return len(text) - (0.5 if text.startswith("-") else 0.0)


def column_units(column_width: float, font: FontSpec | None = None) -> float:
    """The column width Excel keeps after snapping to whole points (8.43 becomes 8.5)."""
    return (column_budget(column_width, font)) / _advances(*resolve_font(font))["0"]


def resolve_font(font: FontSpec | None) -> tuple[str, float]:
    """Canonical (name, size); raises LookupError for a font without a bundled table."""
    if font is None:
        return DEFAULT_FONT
    name, size = (font, 11.0) if isinstance(font, str) else (font[0], float(font[1]))
    if name.lower() not in _FONTS:
        bundled = ", ".join(sorted(k.title() for k in _FONTS))
        msg = f"no glyph table for font {name!r}; bundled: {bundled}"
        raise LookupError(msg)
    return name.title(), size


@lru_cache(maxsize=16)
def _advances(name: str, size: float) -> dict[str, float]:
    tables = _FONTS[name.lower()]
    if size in tables:
        return tables[int(size)]
    # An unmeasured size: scale the nearest measured table (the larger one on a tie) and round
    # each advance to whole points, the way hinting does. An approximation; the measured sizes
    # (Calibri 9, 10, 11, 12, 14 and 16) are exact.
    nearest = min(tables, key=lambda measured: (abs(measured - size), -measured))
    base = tables[nearest]
    return {ch: float(round(adv * size / nearest)) for ch, adv in base.items()}


def glyph_measure(font: FontSpec | None = None) -> Measure:
    """Excel mode: floored sum of glyph advances in points at the font's size.

    Glyphs outside the table measure as a digit (the maximum digit width), which is what a
    column width unit is; the documentation says so.
    """
    name, size = resolve_font(font)
    adv = _advances(name, size)
    digit = adv["0"]

    def measure(text: str) -> float:
        return float(math.floor(sum(adv.get(c, digit) for c in text) + 1e-9))

    return measure


def column_budget(column_width: float, font: FontSpec | None = None) -> float:
    """Points available to text in a column of ``column_width`` units (Column Width dialog)."""
    name, size = resolve_font(font)
    mdw = _advances(name, size)["0"]
    return round(mdw * column_width + _COLUMN_PADDING) - _COLUMN_PADDING
