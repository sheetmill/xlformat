"""Measure functions and the column model for the width decision (``ARCHITECTURE.md`` §8).

Character mode measures a text by its length and the budget is the width in characters. Excel
mode measures a text as the floored sum of glyph advances in points at the cell's font size and
the budget is the column's text area in points: Excel for Mac stores a column of ``w`` units as
``round(MDW * w + P)`` points, MDW being the maximum digit width of the Normal font (5 points at
9 and 10, 6 at 11 and 12, 7 at 14, 8 at 16, 9 at 18, 10 at 20 and 12 at 24, in Calibri and Aptos
Narrow alike; Arial 5, 6, 6, 7, 8, 9, 10, 11 and 13) and ``P`` the padding of
:func:`column_padding` (5 points for a digit width up to 8, 7 for 9 to 12, 9 for 13), and
compares text against those points less the padding. Both modes feed the same seam in
``_assemble``: fit when ``measure(text) <= budget``, ``#`` and ``*`` counts by dividing the
budget by the glyph's measure (CAPTURING.md section 2 items 24 and 25).
"""

from __future__ import annotations

import math
import warnings
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from functools import lru_cache

from xlformat._glyphs_aptos_narrow import APTOS_NARROW, APTOS_NARROW_OVERHANG
from xlformat._glyphs_arial import ARIAL, ARIAL_OVERHANG
from xlformat._glyphs_calibri import CALIBRI, CALIBRI_OVERHANG

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
        underline: The underline as a file names it: ``None``, ``"single"``, ``"double"``,
            ``"singleAccounting"`` or ``"doubleAccounting"``. Only :func:`row_metrics` reads
            it: an accounting underline makes a row taller.
    """

    name: str
    size: float = 11
    bold: bool = False
    italic: bool = False
    underline: str | None = None

    def __post_init__(self) -> None:
        if self.underline not in _UNDERLINES:
            names = ", ".join(repr(u) for u in _UNDERLINES)
            msg = f"underline must be one of {names}, not {self.underline!r}"
            raise ValueError(msg)

    @property
    def style(self) -> str:
        """``"regular"``, ``"bold"``, ``"italic"`` or ``"bold italic"``."""
        if self.bold and self.italic:
            return "bold italic"
        return "bold" if self.bold else ("italic" if self.italic else "regular")


_UNDERLINES = (None, "single", "double", "singleAccounting", "doubleAccounting")
FontSpec = str | tuple[str, float] | Font
"""A font name, a name and a point size, or a :class:`Font`; the name alone means 11 points."""

DEFAULT_FONT = Font("Calibri", 11)
"""openpyxl's default workbook font; a reader of a workbook passes its Normal style instead."""

# family (lower case) -> (display name, style -> size -> advances in points)
_FONTS: dict[str, tuple[str, dict[str, dict[int, dict[str, float]]]]] = {
    "calibri": ("Calibri", CALIBRI),
    "aptos narrow": ("Aptos Narrow", APTOS_NARROW),
    "arial": ("Arial", ARIAL),
}
# family (lower case) -> style -> size -> points an italic face's text reaches past its advances
# on each side: the cell's budget is that much smaller (CAPTURING.md section 2 item 30)
_OVERHANG: dict[str, dict[str, dict[int, float]]] = {
    "calibri": CALIBRI_OVERHANG,
    "aptos narrow": APTOS_NARROW_OVERHANG,
    "arial": ARIAL_OVERHANG,
}


def column_padding(mdw: float) -> float:
    """Points Excel adds to ``MDW * width`` for a column, not available to text.

    ``2 * ceil(MDW / 4) + 1``: 5 for a maximum digit width of 5 to 8 points, 7 for 9 to 12, 9
    for 13 (Arial 24; CAPTURING.md section 2 items 48 and 55).
    """
    return 2.0 * math.ceil(mdw / 4) + 1.0


_LEFT_INSET = 3.0  # where left-aligned text starts (capture/alignment.py, PDF export)
_RIGHT_INSET = 2.0  # where right-aligned text ends, from the right gridline (same capture)
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
    return Font(display, float(font.size), font.bold, font.italic, font.underline)


def _table(font: FontSpec | None) -> dict[str, float]:
    f = resolve_font(font)
    return _whole_advances(f.name, f.size, f.style)


@lru_cache(maxsize=64)
def _whole_advances(name: str, size: float, style: str) -> dict[str, float]:
    """The advances fit and fill decisions use: whole points, without the tables' tenths.

    A tenth (``W`` is 10.1 at Calibri 11) is real only in autofit, where ten ``W`` measure 101;
    ten ``W`` and a digit fit a budget of 106, and ``*W0`` holds 10 of them there
    (CAPTURING.md section 2 item 41). A line feed and a tab take the space's advance and a
    carriage return none, so CR LF measures one space (item 51).
    """
    whole = {c: float(math.floor(a + 1e-9)) for c, a in _advances(name, size, style).items()}
    whole["\n"] = whole["\t"] = whole[" "]
    whole["\r"] = 0.0
    return whole


@lru_cache(maxsize=64)
def _advances(name: str, size: float, style: str = "regular") -> dict[str, float]:
    tables = _FONTS[name.lower()][1][style]
    if size in tables:
        return tables[int(size)]
    # An unmeasured size: scale the nearest measured table (the larger one on a tie) and round
    # each advance to whole points, the way hinting does. An approximation; the measured sizes
    # (9, 10, 11, 12, 14, 16, 18, 20 and 24) are exact.
    nearest = min(tables, key=lambda measured: (abs(measured - size), -measured))
    base = tables[nearest]
    return {ch: float(round(adv * size / nearest)) for ch, adv in base.items()}


def overhang(font: FontSpec | None = None) -> float:
    """Points an italic face reaches past its advances on each side; 0 for upright faces.

    The text area of a cell in that font is this much smaller: an italic Calibri 11 cell holds
    one fill space fewer than a regular one. An unmeasured size scales the nearest measured
    overhang to whole points.
    """
    f = resolve_font(font)
    return _overhang(f.name, f.size, f.style)


@lru_cache(maxsize=64)
def _overhang(name: str, size: float, style: str) -> float:
    table = _OVERHANG[name.lower()].get(style, {})
    if not table:
        return 0.0
    if size in table:
        return float(table[int(size)])
    nearest = min(table, key=lambda measured: (abs(measured - size), -measured))
    return float(round(table[nearest] * size / nearest))


def glyph_measure(font: FontSpec | None = None) -> Measure:
    """Excel mode: the sum of whole glyph advances in points at the font's size.

    A line feed and a tab measure as a space, a carriage return as nothing; other glyphs outside
    the table measure as a digit (the maximum digit width), which is what a column width unit
    is; ``text_width`` and the ``font`` argument of ``format_value`` state it.
    """
    adv = _table(font)
    digit = adv["0"]

    def measure(text: str) -> float:
        return sum(adv.get(c, digit) for c in text)

    return measure


def text_width(text: str, font: FontSpec | None = None) -> float:
    """The width of ``text`` in points as Excel for Mac measures it in ``font``.

    The sum of the whole glyph advances, the measure every Excel-mode decision uses (fit, ``#``
    count, fill count); a line feed and a tab measure as a space, a carriage return as nothing,
    and other glyphs outside the table as a digit. Like reportlab's
    ``stringWidth``, but whole points as Excel hints them. Excel's autofit alone counts a tenth
    on a few wide glyphs.

    Raises:
        LookupError: ``font`` has no measured table.
    """
    return glyph_measure(font)(text)


def glyph_widths(text: str, font: FontSpec | None = None) -> tuple[float, ...]:
    """The advance of each character of ``text`` in whole points.

    For measuring runs of a text (words, spaces): ``text_width`` of a run is
    ``sum(glyph_widths(run))``.
    """
    adv = _table(font)
    digit = adv["0"]
    return tuple(float(adv.get(c, digit)) for c in text)


def digit_width(font: FontSpec | None = None) -> float:
    """The maximum digit width (MDW) of ``font`` in points.

    One column unit of a workbook whose Normal font is ``font``, and one indent level.
    """
    return _mdw(font)


def measured_fonts() -> dict[str, dict[str, tuple[int, ...]]]:
    """The glyph tables xlformat bundles: family, then style, then the measured sizes.

    Other sizes of a measured family and style scale the nearest measured table; anything else
    raises ``LookupError``. A caller can test a workbook's fonts against this before choosing
    Excel mode.
    """
    return {
        display: {style: tuple(sorted(sizes)) for style, sizes in styles.items()}
        for display, styles in _FONTS.values()
    }


def column_budget(column_width: float, font: FontSpec | None = None) -> float:
    """Points available to a number or date in a column of ``column_width`` dialog units."""
    return column_points(column_width, font) - column_padding(_mdw(font))


def column_points(column_width: float, font: FontSpec | None = None) -> float:
    """A column's width in whole points.

    ``round(MDW * w + P)``, and below one unit ``round(w * (MDW + P))``, ``P`` being
    :func:`column_padding` (CAPTURING.md section 2 items 28 and 48).
    """
    mdw = _mdw(font)
    if column_width < 1:
        return float(round(column_width * (mdw + column_padding(mdw))))
    return float(round(mdw * column_width + column_padding(mdw)))


def _mdw(font: FontSpec | None) -> float:
    return float(_table(font)["0"])


def _units(points: float, mdw: float) -> float:
    """Dialog units of a column of ``points``, exact (Excel shows them to two decimals)."""
    pad = column_padding(mdw)
    if points < mdw + pad:
        return points / (mdw + pad)
    return (points - pad) / mdw


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
            8.43 becomes 8.5 at Calibri 11); pass it as ``column_width``, unless it is 0 (a
            column of no width, which ``column_width`` refuses with ``ValueError``).
        points: The column's width in points (53 for an openpyxl column without a width).
        budget: Points a number or date may take before it becomes a ``#`` run: ``points``
            less the column padding (5 points for a digit width up to 8, 7 for 9 to 12, 9 for
            13), and 0
            for a column narrower than that.
        left_inset: Where left-aligned text starts, in points from the left gridline.
        right_inset: Where right-aligned text ends, in points from the right gridline.
        digit_width: The maximum digit width of the Normal font: one column unit, one indent
            level.

    The insets are measured in Excel's print layout at Calibri 11 (a PDF export,
    ``capture/alignment.py``); other sizes and the screen are not captured.
    ``points - left_inset`` is how far a left-aligned text may reach before it spills into the
    next column (50 of a 53-point column).
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
        budget=max(points - column_padding(mdw), 0.0),
        left_inset=_LEFT_INSET,
        right_inset=_RIGHT_INSET,
        digit_width=mdw,
    )


def _normal_font_argument(normal_font: FontSpec | None, font: FontSpec | None) -> FontSpec | None:
    """The Normal font a column function was given, warning on the deprecated ``font``."""
    if font is None:
        return normal_font
    if normal_font is not None:
        msg = "pass normal_font or the deprecated font, not both"
        raise ValueError(msg)
    warnings.warn(
        "font is deprecated here and will be removed in 1.0: pass normal_font, the workbook's "
        "Normal font (in the format functions, font is the cell's font)",
        DeprecationWarning,
        stacklevel=3,
    )
    return font


def column_metrics(
    column_width: float,
    normal_font: FontSpec | None = None,
    *,
    font: FontSpec | None = None,
) -> ColumnMetrics:
    """The column of ``column_width`` dialog units in the workbook's Normal font.

    Args:
        column_width: The width in Column Width dialog units.
        normal_font: The workbook's Normal font, whose digit width sets the column's units.
        font: Deprecated name of ``normal_font``, removed in 1.0.

    Raises:
        ValueError: ``column_width`` is not positive, or both font arguments were passed.
    """
    normal_font = _normal_font_argument(normal_font, font)
    if column_width <= 0:
        msg = f"column_width must be positive, not {column_width!r}"
        raise ValueError(msg)
    return _metrics(column_points(column_width, normal_font), normal_font)


def column_metrics_from_points(points: float, normal_font: FontSpec | None = None) -> ColumnMetrics:
    """The column ``points`` wide, in the workbook's Normal font.

    For a merged cell, pass the sum of its columns' ``points``: Excel formats it as one column of
    that width, with a single padding (CAPTURING.md section 2 item 40), and a hidden column counts
    as 0. The result's ``width`` is what ``column_width`` takes, unless it is 0.

    Raises:
        ValueError: ``points`` is negative.
    """
    if points < 0:
        msg = f"points must not be negative, not {points!r}"
        raise ValueError(msg)
    return _metrics(float(points), normal_font)


def file_column_metrics(
    file_width: float | None = None,
    *,
    base_col_width: float = 8,
    default_col_width: float | None = None,
    normal_font: FontSpec | None = None,
    font: FontSpec | None = None,
) -> ColumnMetrics:
    """The column a file describes: ``<col width>``, or ``None`` for a column without one.

    A column without ``<col>`` takes ``<sheetFormatPr defaultColWidth>`` as a file width, and
    without that ``baseColWidth`` characters plus the padding: ``baseColWidth * MDW + P``
    points, ``P`` being ``2 * ceil(MDW / 4) + 1`` (5 for a digit width up to 8, 7 for 9 to 12, 9
    for 13);
    53 at Calibri 11 for the ``baseColWidth="8"`` openpyxl writes, which the dialog shows as 8
    (not 8.43). ``normal_font`` is the workbook's Normal font (CAPTURING.md section
    2 item 28); ``font`` is its deprecated name, removed in 1.0, and passing both raises
    ``ValueError``.
    """
    normal_font = _normal_font_argument(normal_font, font)
    if file_width is None:
        if default_col_width is None:
            mdw = _mdw(normal_font)
            return _metrics(base_col_width * mdw + column_padding(mdw), normal_font)
        file_width = default_col_width
    if file_width < 0:
        msg = f"file_width must not be negative, not {file_width!r}"
        raise ValueError(msg)
    return _metrics(file_points(file_width, normal_font), normal_font)


# --- the row model (CAPTURING.md section 2 items 44, 55 and 57) --------------------------------

# family (lower case) -> a: from 7.5 points a face's height (the row Excel fits to it alone, in
# any style) is floor(a * ceil(size) + 1.5); each slope is the middle of the window that every
# captured size from 7.5 to 72 allows (font_height.layout.tsv)
_SLOPE: dict[str, float] = {
    "calibri": 1.2689,
    "aptos narrow": 1.2815,
    "arial": 1.2122,
    "times new roman": 1.2162,
    "courier new": 1.3007,
    "georgia": 1.2190,
}
_DISPLAY = {
    "calibri": "Calibri",
    "aptos narrow": "Aptos Narrow",
    "arial": "Arial",
    "times new roman": "Times New Roman",
    "courier new": "Courier New",
    "georgia": "Georgia",
}
_SMALL = 7.5  # below this size the heights follow no slope: the captured ones, per half point
# 1 to 5.5 points alike in every face captured; 6, 6.5 and 7 per face
_SMALL_COMMON: dict[float, float] = {
    1: 4, 1.5: 7, 2: 7, 2.5: 7, 3: 7, 3.5: 8, 4: 8, 4.5: 9, 5: 9, 5.5: 10,
}  # fmt: skip
_SMALL_HEIGHT: dict[str, dict[float, float]] = {
    "calibri": {6: 9, 6.5: 10, 7: 10},
    "aptos narrow": {6: 9, 6.5: 10, 7: 10},
    "arial": {6: 10, 6.5: 11, 7: 11},
    "times new roman": {6: 10, 6.5: 10, 7: 10},
    "courier new": {6: 9, 6.5: 10, 7: 10},
    "georgia": {6: 10, 6.5: 10, 7: 10},
}
_LARGEST = 72.0  # the largest size captured


def _font_height(key: str, size: float) -> float:
    """The face's height: the row Excel fits to it alone."""
    if size < _SMALL:
        half = math.ceil(size * 2) / 2
        return float(_SMALL_HEIGHT[key].get(half, _SMALL_COMMON.get(half, 0.0)))
    return float(math.floor(_SLOPE[key] * math.ceil(size) + 1.5))


# family (lower case) -> post.underlineThickness / unitsPerEm: an accounting underline adds
# 3 * max(1, floor(size * this)) points to the row and to each wrapped line (CAPTURING.md section 2
# item 58); Courier New follows another rule, not known
_ACCOUNTING_THICKNESS: dict[str, float] = {
    "calibri": 134 / 2048,
    "aptos narrow": 50 / 2048,
    "arial": 150 / 2048,
    "times new roman": 100 / 2048,
    "georgia": 101 / 2048,
}
# Arial's double accounting adds 3 up to 35 points, then 6, then 9 from 55 (captured)
_ARIAL_DOUBLE_ACCOUNTING = ((55, 9.0), (36, 6.0), (0, 3.0))


def _accounting(key: str, size: float, underline: str | None) -> float:
    """Points an accounting underline adds to each line of the face; LookupError if not known."""
    if underline not in ("singleAccounting", "doubleAccounting"):
        return 0.0
    whole = math.ceil(size)
    known = (
        key in _ACCOUNTING_THICKNESS
        and 6 <= whole <= _LARGEST
        and not (key == "times new roman" and 37 <= whole <= 40)
    )
    if not known:
        msg = f"no accounting underline rows for {_DISPLAY[key]} {size:g}"
        raise LookupError(msg)
    if key == "arial" and underline == "doubleAccounting":
        return next(e for start, e in _ARIAL_DOUBLE_ACCOUNTING if whole >= start)
    return 3.0 * max(1, math.floor(whole * _ACCOUNTING_THICKNESS[key]))


def _bottom(key: str, size: float) -> float:
    """B, the face's descent: its height less its ascent, ceil(size) + 1 points."""
    return _font_height(key, size) - math.ceil(size) - 1.0


@dataclass(frozen=True, slots=True)
class RowMetrics:
    """A row as Excel for Mac fits it when it opens a file, in whole points.

    Attributes:
        height: The row's height: what Excel gives a row without a stored height (no ``ht``)
            holding the fonts and wrapped text given.
        font_height: The Normal font's height: a row in the Normal font alone, and the sheet's
            standard height.
        line_height: What each wrapped line in the Normal font takes: ``font_height`` plus 1, so
            one wrapped line is a point taller than the row without it.
    """

    height: float
    font_height: float
    line_height: float


def _row_face(font: FontSpec | None) -> tuple[str, float]:
    """(family key, size) of a face the row model has measured; LookupError otherwise."""
    if font is None:
        f = DEFAULT_FONT
    elif isinstance(font, str):
        f = Font(font, 11.0)
    elif isinstance(font, Font):
        f = font
    else:
        f = Font(font[0], float(font[1]))
    key = f.name.lower()
    if key not in _SLOPE:
        measured = ", ".join(_DISPLAY.values())
        msg = f"no row heights for font {f.name!r}; measured: {measured}"
        raise LookupError(msg)
    if not 1.0 <= f.size <= _LARGEST:
        msg = f"no row heights for {_DISPLAY[key]} {f.size:g}; measured sizes: 1 to 72"
        raise LookupError(msg)
    return key, float(f.size)


def row_metrics(
    normal_font: FontSpec | None = None,
    *,
    cell_fonts: Iterable[FontSpec] = (),
    wrapped: Iterable[tuple[FontSpec | None, int]] = (),
) -> RowMetrics:
    """The height Excel gives a row without a stored height, in the workbook's Normal font.

    A row holds its fonts' largest ascent and largest descent, rounded up: the Normal font and every
    cell's font in the row, an empty cell styled with a font included. Faces of one family split
    their heights alike, so the row is as tall as its tallest font; Arial's descent is smaller, so
    Arial 11 (14 points) beside Calibri 10 (14 points) makes a row of 15. Wrapped text of ``lines``
    lines in a face takes ``lines * (height + 1)`` points of that face, plus, when the Normal font
    or another cell's font in the row has a larger line inset, the difference of the insets
    (CAPTURING.md section 2 items 44, 55 and 57). A face's height is ``floor(a * ceil(size) +
    1.5)`` from 7.5 points, ``a`` its slope, and captured below that; its ascent is ``ceil(size) +
    1``, its descent and line inset follow from the two. Bold and italic change nothing.
    Measured: Calibri, Aptos Narrow, Arial, Times New Roman, Courier New and Georgia from 1 to 72
    points; any other face or size raises ``LookupError``.

    Several cells with wrapped text in one row are each computed this way and the row takes the
    tallest; that combination has not been captured.

    Args:
        normal_font: The workbook's Normal font.
        cell_fonts: The fonts of the row's cells that are not the Normal font and do not wrap.
        wrapped: One ``(font, lines)`` per cell with wrapped text: its font (``None`` for the
            Normal font) and its number of lines, as the caller broke them.

    Raises:
        LookupError: A font or size the row model has not measured.
        ValueError: A number of lines below 1.
    """
    normal = _row_face(normal_font)
    font_height = _font_height(*normal)

    def underline(font: FontSpec | None) -> str | None:
        return font.underline if isinstance(font, Font) else None

    # a row holds its fonts' largest ascent (size + 1) and largest descent B; the largest line
    # inset (B - 2) of its unwrapped fonts lifts wrapped text in a face with a smaller one
    ascent = math.ceil(normal[1]) + 1.0
    descent = _bottom(*normal) + _accounting(*normal, underline(normal_font))
    top_inset = _bottom(*normal) - 2.0
    for font in cell_fonts:
        key, size = _row_face(font)
        ascent = max(ascent, math.ceil(size) + 1.0)
        descent = max(descent, _bottom(key, size) + _accounting(key, size, underline(font)))
        top_inset = max(top_inset, _bottom(key, size) - 2.0)
    wraps = []
    for wrap_font, lines in wrapped:
        if lines < 1:
            msg = f"wrapped text has at least one line, not {lines!r}"
            raise ValueError(msg)
        key, size = normal if wrap_font is None else _row_face(wrap_font)
        extra = _accounting(key, size, underline(normal_font if wrap_font is None else wrap_font))
        wraps.append((key, size, lines, extra))
        # the wrapped cell's font is in the row too: one line of Aptos Narrow 9 under Arial 11
        # makes the row 15 points, as the two faces unwrapped do
        ascent = max(ascent, math.ceil(size) + 1.0)
        descent = max(descent, _bottom(key, size) + extra)
    height = ascent + descent
    for key, size, lines, extra in wraps:
        # an accounting underline adds its points to every wrapped line
        inset = max(0.0, top_inset - (_bottom(key, size) - 2.0))
        height = max(height, lines * (_font_height(key, size) + 1.0 + extra) + inset)
    return RowMetrics(height=height, font_height=font_height, line_height=font_height + 1.0)


# --- vertical placement (CAPTURING.md section 2 item 47) ----------------------------------------

# family (lower case) -> size -> (cap height, descender): the ink of "H" above the baseline and of
# "g" below it, as Excel draws them at zoom 100 on a 2x display (half points)
_INK: dict[str, dict[int, tuple[float, float]]] = {
    "calibri": {
        9: (6.0, 1.5), 10: (6.5, 2.0), 11: (7.0, 2.0), 12: (8.0, 2.0), 13: (8.5, 2.5),
        14: (9.0, 2.5), 15: (9.5, 2.5), 16: (10.5, 3.0), 17: (11.0, 3.0), 18: (11.5, 3.0),
        20: (13.0, 3.5), 22: (14.0, 4.0), 24: (15.5, 4.5),
    },
    "aptos narrow": {
        9: (6.0, 1.5), 10: (6.5, 2.0), 11: (7.5, 2.0), 12: (8.0, 2.0), 13: (8.5, 2.5),
        14: (9.5, 2.5), 15: (10.0, 3.0), 16: (10.5, 3.0), 17: (11.5, 3.0), 18: (12.0, 3.5),
        20: (13.5, 3.5), 22: (14.5, 4.0), 24: (16.0, 4.5),
    },
    "arial": {
        9: (6.5, 2.0), 10: (7.5, 2.0), 11: (8.0, 2.5), 12: (9.0, 2.5), 14: (10.0, 3.0),
        16: (11.5, 3.0), 18: (13.0, 4.0), 20: (14.5, 4.0), 24: (17.5, 5.0),
    },
    "times new roman": {
        9: (6.0, 2.0), 10: (7.0, 2.0), 11: (7.5, 2.5), 12: (8.0, 2.5), 14: (9.5, 3.0),
        16: (11.0, 3.0), 18: (12.0, 4.0), 20: (13.5, 4.0), 24: (16.0, 5.0),
    },
    "courier new": {
        9: (5.5, 1.5), 10: (6.0, 2.0), 11: (6.5, 2.0), 12: (7.0, 2.5), 14: (8.0, 2.5),
        16: (9.5, 3.0), 18: (10.5, 3.5), 20: (11.5, 4.0), 24: (14.0, 4.5),
    },
    "georgia": {
        9: (6.5, 2.0), 10: (7.0, 2.0), 11: (8.0, 2.5), 12: (8.5, 2.5), 14: (10.0, 3.0),
        16: (11.5, 3.5), 18: (12.5, 4.0), 20: (14.0, 4.0), 24: (17.0, 5.0),
    },
}  # fmt: skip


@dataclass(frozen=True, slots=True)
class VerticalMetrics:
    """Where Excel for Mac draws a face's text in a cell, in sheet points.

    Every field is a distance Excel's screen shows, not a split of one by convention: the top
    inset and the ascent are seen only as their sum, ``baseline_from_top``, and the descent and
    the bottom inset as ``baseline_from_bottom``.

    Attributes:
        baseline_from_top: The baseline's distance below the row's top for top-aligned text: the
            face's size in points plus 1. A row's top is the top of the gridline above it, which
            the row's own fill covers.
        baseline_from_bottom: The baseline's distance above the row's bottom for bottom-aligned
            text (Excel's default). It is at least the Normal font's own, and includes that
            when :func:`vertical_metrics` is given ``normal_font``: Calibri 9 and Arial 9 to 11
            have 2 points of their own and sit 3 up in a Calibri 11 or Aptos Narrow 11
            workbook, 2 in an Arial 10 one.
            Centered text sits midway between the two placements, to half a point.
        font_height: The height of a row in this face alone (no stored height).
        line_height: What each wrapped line in this face adds to a row: ``font_height`` plus 1.
        line_pitch: The distance between the baselines of wrapped lines: ``font_height``; the
            last line is placed as bottom-aligned text is.
        line_inset: The face's line inset, the ``p`` of the row model. Wrapped text in this face,
            in a row whose Normal font or other cell fonts have a larger ``line_inset``, takes
            that difference more height (``row_metrics``) and sits that much higher, the extra
            lying below its last line.

    The ink of the text, which is measured rather than given by a rule, is
    :func:`ink_metrics`'s; ``cap_height`` and ``descender`` here are deprecated and forward to it.
    """

    baseline_from_top: float
    baseline_from_bottom: float
    font_height: float
    line_height: float
    line_pitch: float
    line_inset: float
    _font: Font | None = field(default=None, repr=False, compare=False)

    @property
    def cap_height(self) -> float:
        """Deprecated (removed in 1.0); use ``ink_metrics(font).cap_height``."""
        return self._ink("cap_height").cap_height

    @property
    def descender(self) -> float:
        """Deprecated (removed in 1.0); use ``ink_metrics(font).descender``."""
        return self._ink("descender").descender

    def _ink(self, name: str) -> InkMetrics:
        warnings.warn(
            f"VerticalMetrics.{name} is deprecated and will be removed in 1.0: use "
            f"ink_metrics(font).{name}",
            DeprecationWarning,
            stacklevel=3,
        )
        return ink_metrics(self._font)


@dataclass(frozen=True, slots=True)
class InkMetrics:
    """The ink Excel for Mac draws for a face at zoom 100 on a 2x display, in sheet points.

    Attributes:
        cap_height: The ink of "H" above the baseline, to half a point.
        descender: The ink of "g" below the baseline, to half a point. Arial 11 and 16, Times
            New Roman 11, 16 and 20 and Georgia 11 and 20 draw it half a point longer or shorter
            by where the baseline falls on the display's pixels.
    """

    cap_height: float
    descender: float


def ink_metrics(font: FontSpec | None = None) -> InkMetrics:
    """The measured ink of ``font``'s "H" and "g" (CAPTURING.md section 2 items 47 and 55).

    Measured: Calibri, Aptos Narrow, Arial, Times New Roman, Courier New and Georgia at 9, 10, 11,
    12, 14, 16, 18, 20 and 24 points, and Calibri and Aptos Narrow at 13, 15, 17 and 22, in every
    style alike (styles captured for the first three).

    Raises:
        LookupError: A face or size not measured.
    """
    key, size = _row_face(font)
    ink = _INK.get(key, {}).get(int(size)) if size == int(size) else None
    if ink is None:
        msg = (
            f"no ink measured for {_DISPLAY[key]} {size:g}; measured: the six faces at 9, 10, "
            "11, 12, 14, 16, 18, 20 and 24, and Calibri and Aptos Narrow at 13, 15, 17 and 22"
        )
        raise LookupError(msg)
    return InkMetrics(cap_height=ink[0], descender=ink[1])


def vertical_metrics(
    font: FontSpec | None = None, *, normal_font: FontSpec | None = None
) -> VerticalMetrics:
    """Where Excel draws text in ``font`` within a cell (CAPTURING.md section 2 items 47 and 55).

    Bold and italic place as the regular face does. The workbook's Normal font moves text in two
    cases. Bottom-aligned text sits at least the Normal font's own ``baseline_from_bottom`` up:
    with ``normal_font`` given, ``baseline_from_bottom`` already includes this (Calibri 9 has 2
    points of its own and 3 under a Calibri 11 Normal font). Wrapped text in a face with a smaller
    ``line_inset`` than the row's other fonts sits that difference higher; that is left to the
    caller, because it depends on every unwrapped font in the row, not only the Normal font, and
    applies to wrapped text alone (see :class:`VerticalMetrics` and :func:`row_metrics`).
    The placement follows from the face's height (:func:`row_metrics`): top-aligned text has its
    baseline ``ceil(size) + 1`` below the row's top and bottom-aligned text ``B = height -
    ceil(size) - 1`` above its bottom, and the line inset is ``B - 2``. Supported: Calibri, Aptos
    Narrow, Arial, Times New Roman, Courier New and Georgia from 1 to 72 points; any other face
    or size raises ``LookupError``.

    Args:
        font: The cell's font; default ``("Calibri", 11)``.
        normal_font: The workbook's Normal font; default ``None``, the cell's own font, which
            gives the face's own placement (a cell in the Normal font).

    Raises:
        LookupError: A font or size not measured.
    """
    key, size = _row_face(font)
    height = _font_height(key, size)
    from_bottom = _bottom(key, size)
    if normal_font is not None:
        from_bottom = max(from_bottom, _bottom(*_row_face(normal_font)))
    return VerticalMetrics(
        baseline_from_top=math.ceil(size) + 1.0,
        baseline_from_bottom=from_bottom,
        font_height=height,
        line_height=height + 1.0,
        line_pitch=height,
        line_inset=_bottom(key, size) - 2.0,
        _font=Font(_DISPLAY[key], size),
    )


# --- rules: underlines and strikethrough (CAPTURING.md section 2 item 60) ------------------------

# family (lower case) -> (post.underlinePosition, post.underlineThickness, OS/2.yStrikeoutPosition,
# OS/2.yStrikeoutSize, OS/2.usWinDescent), in units of 2048 an em (capture/font_tables.py)
_RULE_TABLES: dict[str, tuple[int, int, int, int, int]] = {
    "calibri": (-232, 134, 512, 134, 550),
    "aptos narrow": (-100, 50, 731, 102, 563),
    "arial": (-217, 150, 530, 102, 434),
    "times new roman": (-223, 100, 530, 102, 443),
    "georgia": (-181, 101, 594, 101, 449),
}


@dataclass(frozen=True, slots=True)
class RuleMetrics:
    """Where Excel for Mac draws a face's rules at a zoom, in sheet points from the baseline.

    Each value is computed at the screen size ``size * zoom / 100`` in whole screen points and
    divided back, so it moves in steps of ``100 / zoom`` points.

    Attributes:
        underline_top: The top of a single underline, below the baseline.
        underline_thickness: The thickness of every underline line, accounting ones included.
        double_gap: How far a double underline's second line lies below its first (both kinds).
        accounting_top: The top of an accounting underline, below the baseline.
        strikethrough_top: The top of a strikethrough, above the baseline.
        strikethrough_thickness: The thickness of a strikethrough.
    """

    underline_top: float
    underline_thickness: float
    double_gap: float
    accounting_top: float
    strikethrough_top: float
    strikethrough_thickness: float


def rule_metrics(font: FontSpec | None = None, *, zoom: float = 100) -> RuleMetrics:
    """Where Excel draws underlines and a strikethrough in ``font`` at ``zoom`` percent.

    From the font's own tables at the screen size ``P = size * zoom / 100`` (CAPTURING.md
    section 2 item 60): an underline's top ``floor(P * -underlinePosition / em)`` below the
    baseline and its thickness ``max(1, floor(P * underlineThickness / em))``; an accounting
    underline's top ``round(P * usWinDescent / em)``; a strikethrough's top ``round(P *
    yStrikeoutPosition / em)`` above the baseline, ``max(1, round(P * yStrikeoutSize / em))``
    thick; a double's second line two thicknesses below its first. Captured at 9, 11, 14, 20 and
    36 points, zoom 100 and 200, in Calibri, Aptos Narrow, Arial, Times New Roman and Georgia.
    Known misses: a double underline's first line sometimes sits higher than a single one (24 of
    60 captured rows); Aptos Narrow's accounting underline at 9 and 20 points, zoom 100, sits a
    point lower; a few lines at zoom 200 are thinner. Courier New draws its underline elsewhere
    than its tables say, so it raises.

    Raises:
        LookupError: Courier New or another face without its rule, or a size outside 1 to 72.
    """
    key, size = _row_face(font)
    if key not in _RULE_TABLES:
        msg = (
            f"no rule positions for {_DISPLAY[key]}; measured: Calibri, Aptos Narrow, Arial, "
            "Times New Roman and Georgia"
        )
        raise LookupError(msg)
    position, thickness, strike_position, strike_size, descent = _RULE_TABLES[key]
    scale = zoom / 100
    p = size * scale

    def rounded(units: int) -> float:
        return float(math.floor(p * units / 2048 + 0.5))

    line = max(1.0, float(math.floor(p * thickness / 2048)))
    return RuleMetrics(
        underline_top=math.floor(p * -position / 2048) / scale,
        underline_thickness=line / scale,
        double_gap=2 * line / scale,
        accounting_top=rounded(descent) / scale,
        strikethrough_top=rounded(strike_position) / scale,
        strikethrough_thickness=max(1.0, rounded(strike_size)) / scale,
    )
