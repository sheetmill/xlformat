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
from dataclasses import dataclass
from functools import lru_cache

from xlformat._glyphs_calibri import CALIBRI

Measure = Callable[[str], float]
"""Width of a text in the caller's units."""


@dataclass(frozen=True, slots=True)
class Font:
    """A font for Excel mode: family, size in points and style.

    Attributes:
        name: The family, such as ``"Calibri"``; matched without regard to case.
        size: The size in points.
        bold: Bold.
        italic: Italic.
    """

    name: str
    size: float = 11
    bold: bool = False
    italic: bool = False

    @property
    def style(self) -> str:
        """``"regular"``, ``"bold"``, ``"italic"`` or ``"bold italic"``."""
        if self.bold and self.italic:
            return "bold italic"
        return "bold" if self.bold else ("italic" if self.italic else "regular")


FontSpec = str | tuple[str, float] | Font
"""A font name, a name and a point size, or a :class:`Font`; the name alone means 11 points."""

DEFAULT_FONT = Font("Calibri", 11)
"""openpyxl's default workbook font; sheetpeek passes the workbook's Normal style."""

# family (lower case) -> (display name, style -> size -> advances in points)
_FONTS: dict[str, tuple[str, dict[str, dict[int, dict[str, float]]]]] = {
    "calibri": ("Calibri", {"regular": CALIBRI}),
}
_COLUMN_PADDING = 5.0  # points Excel adds to MDW * width for the column, not available to text
_LEFT_INSET = 3.0  # where left-aligned text starts (sheetpeek's PDF capture; provisional)
_RIGHT_INSET = 2.0  # where right-aligned text ends, from the right gridline (provisional)
MAX_TEXT = 255  # Excel shows at most 255 characters: a longer filled text is a '#' run


def character_measure(text: str) -> float:
    """Character mode: one unit per character."""
    return float(len(text))


def general_cost(minus: float) -> Measure:
    """What ``General`` digit reduction counts against the column width in units.

    Excel reduces General by characters, not by glyph widths: one unit per character, and a
    leading minus sign costs its own advance in digits, ``minus`` (3 / 6 at Calibri 11:
    ``-12345678`` fits width 8.5, ``-1.2E+07`` needs 8), while the exponent's minus is a full
    unit (``-2E-07`` needs 6; CAPTURING.md section 2 items 25 and 27). In character mode
    ``minus`` is 1 and the cost is ``len``.
    """

    def cost(text: str) -> float:
        # rounded like the budget it meets: 4 / 7 of a minus ties a width of x.57 exactly
        return round(len(text) - (1.0 - minus if text.startswith("-") else 0.0), 9)

    return cost


def column_units(column_width: float, font: FontSpec | None = None) -> float:
    """The column width Excel keeps after snapping to whole points (8.43 becomes 8.5)."""
    return _units(column_points(column_width, font), _mdw(font))


def resolve_font(font: FontSpec | None) -> Font:
    """The canonical :class:`Font`; raises LookupError for a font or style without a table."""
    if font is None:
        return DEFAULT_FONT
    if isinstance(font, str):
        font = Font(font, 11.0)
    elif not isinstance(font, Font):
        font = Font(font[0], float(font[1]))
    entry = _FONTS.get(font.name.lower())
    if entry is None:
        bundled = ", ".join(display for display, _ in _FONTS.values())
        msg = f"no glyph table for font {font.name!r}; measured: {bundled}"
        raise LookupError(msg)
    display, styles = entry
    if font.style not in styles:
        msg = f"no glyph table for {display} {font.style}; measured: {', '.join(styles)}"
        raise LookupError(msg)
    return Font(display, float(font.size), font.bold, font.italic)


def _table(font: FontSpec | None) -> dict[str, float]:
    f = resolve_font(font)
    return _advances(f.name, f.size, f.style)


@lru_cache(maxsize=64)
def _advances(name: str, size: float, style: str = "regular") -> dict[str, float]:
    tables = _FONTS[name.lower()][1][style]
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
    adv = _table(font)
    digit = adv["0"]

    def measure(text: str) -> float:
        return float(math.floor(sum(adv.get(c, digit) for c in text) + 1e-9))

    return measure


def column_budget(column_width: float, font: FontSpec | None = None) -> float:
    """Points available to a number or date in a column of ``column_width`` dialog units."""
    return column_points(column_width, font) - _COLUMN_PADDING


def column_points(column_width: float, font: FontSpec | None = None) -> float:
    """A column's width in whole points.

    ``round(MDW * w + 5)``, and below one unit ``round(w * (MDW + 5))`` (CAPTURING.md section 2
    item 28).
    """
    mdw = _mdw(font)
    if column_width < 1:
        return float(round(column_width * (mdw + _COLUMN_PADDING)))
    return float(round(mdw * column_width + _COLUMN_PADDING))


def _mdw(font: FontSpec | None) -> float:
    return float(_table(font)["0"])


def _units(points: float, mdw: float) -> float:
    """Dialog units of a column of ``points``, exact (Excel shows them to two decimals)."""
    if points < mdw + _COLUMN_PADDING:
        return points / (mdw + _COLUMN_PADDING)
    return (points - _COLUMN_PADDING) / mdw


def file_points(file_width: float, font: FontSpec | None = None) -> float:
    """Points of a column whose ``<col width>`` is ``file_width``.

    ECMA-376 Part 1 section 18.3.1.13's pixel formula with the maximum digit width in points,
    the width first truncated to 1/256 as Excel stores it:
    ``trunc((256 * w + trunc(128 / MDW)) / 256 * MDW)`` (CAPTURING.md section 2 item 28).
    """
    mdw = _mdw(font)
    w = math.trunc(file_width * 256) / 256
    return float(math.trunc((256 * w + math.trunc(128 / mdw)) / 256 * mdw))


@dataclass(frozen=True, slots=True)
class ColumnMetrics:
    """A column as Excel for Mac lays it out, in whole points.

    Attributes:
        width: The width in Column Width dialog units, exact (Excel shows it to two decimals:
            8.43 becomes 8.5 at Calibri 11); pass it as ``column_width``.
        points: The column's width in points (53 for an openpyxl column without a width).
        budget: Points a number or date may take before it becomes a ``#`` run: ``points``
            less 5.
        left_inset: Where left-aligned text starts, in points from the left gridline.
        right_inset: Where right-aligned text ends, in points from the right gridline.
        digit_width: The maximum digit width of the Normal font: one column unit, one indent
            level.

    The insets are provisional: they come from Excel's print layout at Calibri 11 (a PDF
    export) and wait for a screen capture. ``points - left_inset`` is how far a left-aligned
    text may reach before it spills into the next column (50 of a 53-point column).
    """

    width: float
    points: float
    budget: float
    left_inset: float
    right_inset: float
    digit_width: float


def _metrics(points: float, font: FontSpec | None) -> ColumnMetrics:
    mdw = _mdw(font)
    return ColumnMetrics(
        width=_units(points, mdw),
        points=points,
        budget=points - _COLUMN_PADDING,
        left_inset=_LEFT_INSET,
        right_inset=_RIGHT_INSET,
        digit_width=mdw,
    )


def column_metrics(column_width: float, font: FontSpec | None = None) -> ColumnMetrics:
    """The column of ``column_width`` dialog units in the Normal font ``font``."""
    if column_width <= 0:
        msg = f"column_width must be positive, not {column_width!r}"
        raise ValueError(msg)
    return _metrics(column_points(column_width, font), font)


def file_column_metrics(
    file_width: float | None = None,
    *,
    base_col_width: float = 8,
    default_col_width: float | None = None,
    font: FontSpec | None = None,
) -> ColumnMetrics:
    """The column a file describes: ``<col width>``, or ``None`` for a column without one.

    A column without ``<col>`` takes ``<sheetFormatPr defaultColWidth>`` as a file width, and
    without that ``baseColWidth`` characters plus the padding: ``baseColWidth * MDW + 5``
    points, 53 at Calibri 11 for the ``baseColWidth="8"`` openpyxl writes, which the dialog
    shows as 8 (not 8.43). ``font`` is the workbook's Normal font (CAPTURING.md section 2 item
    28).
    """
    if file_width is None:
        if default_col_width is None:
            return _metrics(base_col_width * _mdw(font) + _COLUMN_PADDING, font)
        file_width = default_col_width
    if file_width < 0:
        msg = f"file_width must not be negative, not {file_width!r}"
        raise ValueError(msg)
    return _metrics(file_points(file_width, font), font)
