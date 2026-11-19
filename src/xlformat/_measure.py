"""Measure functions and the column model for the width decision (``ARCHITECTURE.md`` §8).

Character mode measures a text by its length and the budget is the width in characters. Excel mode
measures a text as the floored sum of glyph advances in points at the cell's font size and the
budget is the column's text area in points: Excel for Mac stores a column of ``w`` units as
``round(MDW * w + P)`` points, MDW being the maximum digit width of the Normal font (5 points at 9
and 10, 6 at 11 and 12, 7 at 14, 8 at 16, 9 at 18, 10 at 20 and 12 at 24, in Calibri and Aptos
Narrow alike; Arial 5, 6, 6, 7, 8, 9, 10, 11 and 13) and ``P`` the padding of :func:`column_padding`
(5 points up to a digit width of 8, 7 from 9 to 12, 9 at 13), and compares text against those points
less the padding. Both modes feed the same seam in ``_assemble``: fit when ``measure(text) <=
budget``, ``#`` and ``*`` counts by dividing the budget by the glyph's measure (CAPTURING.md section
2 items 24 and 25).
"""

from __future__ import annotations

import math
import warnings
from collections.abc import Callable, Iterable
from dataclasses import dataclass
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
    return Font(display, float(font.size), font.bold, font.italic)


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
            less the column padding (5 points up to a digit width of 8, 7 from 9 to 12), and 0
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
    points, ``P`` being ``2 * ceil(MDW / 4) + 1`` (5 up to a digit width of 8, 7 from 9 to 12);
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


# --- the row model (CAPTURING.md section 2 item 44) ----------------------------------------------

# family (lower case) -> size -> the face's font height in points: the height Excel gives a row in
# that font, the same for every style
_FONT_HEIGHT: dict[str, dict[int, float]] = {
    "calibri": {
        9: 12.0,
        10: 14.0,
        11: 15.0,
        12: 16.0,
        14: 19.0,
        16: 21.0,
        18: 24.0,
        20: 26.0,
        24: 31.0,
    },
    "aptos narrow": {
        9: 13.0,
        10: 14.0,
        11: 15.0,
        12: 16.0,
        14: 19.0,
        16: 22.0,
        18: 24.0,
        20: 27.0,
        24: 32.0,
    },
    "arial": {
        9: 12.0,
        10: 13.0,
        11: 14.0,
        12: 16.0,
        14: 18.0,
        16: 20.0,
        18: 23.0,
        20: 25.0,
        24: 30.0,
    },
}
# family (lower case) -> size -> the face's descent in points: a row of several faces is
# ceil(largest ascent + largest descent), where a face's ascent is its height less its descent
# (CAPTURING.md section 2 item 55: fitted to every pair of faces, each value the one nearest the
# font's own descent; Arial's smaller descent makes a row of Arial 11 and Calibri 10 15 points)
_DESCENT: dict[str, dict[int, float]] = {
    "calibri": {
        9: 2.0,
        10: 2.75,
        11: 3.0,
        12: 3.25,
        14: 3.75,
        16: 4.25,
        18: 4.75,
        20: 5.25,
        24: 6.5,
    },
    "aptos narrow": {
        9: 2.5,
        10: 2.75,
        11: 3.0,
        12: 3.25,
        14: 3.75,
        16: 4.5,
        18: 4.75,
        20: 5.75,
        24: 6.75,
    },
    "arial": {9: 2.0, 10: 2.0, 11: 2.25, 12: 3.25, 14: 3.25, 16: 3.5, 18: 4.25, 20: 4.25, 24: 5.5},
}
# family (lower case) -> size -> points of the Normal font's line that wrapped text in a smaller
# face does not fill: k wrapped lines in face f add max(0, inset(normal) - inset(f))
_LINE_INSET: dict[str, dict[int, float]] = {
    "calibri": {9: 0.0, 10: 1.0, 11: 1.0, 12: 1.0, 14: 2.0, 16: 2.0, 18: 3.0, 20: 3.0, 24: 4.0},
    "aptos narrow": {
        9: 1.0,
        10: 1.0,
        11: 1.0,
        12: 1.0,
        14: 2.0,
        16: 3.0,
        18: 3.0,
        20: 4.0,
        24: 5.0,
    },
    "arial": {9: 0.0, 10: 0.0, 11: 0.0, 12: 1.0, 14: 1.0, 16: 1.0, 18: 2.0, 20: 2.0, 24: 3.0},
}


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


def _row_face(font: FontSpec | None) -> tuple[str, int]:
    """(family key, size) of a face the row model has measured; LookupError otherwise."""
    f = resolve_font(font)
    key = f.name.lower()
    sizes = _FONT_HEIGHT.get(key, {})
    if f.size not in sizes:
        measured = ", ".join(f"{size:g}" for size in sizes)
        msg = f"no row heights for {f.name} {f.size:g}; measured sizes: {measured}"
        raise LookupError(msg)
    return key, int(f.size)


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
    (CAPTURING.md section 2 items 44 and 55). Bold and italic change nothing. Measured: Calibri,
    Aptos Narrow and Arial at 9, 10, 11, 12, 14, 16, 18, 20 and 24 points; the row model raises
    ``LookupError`` for any other face ("this font isn't measured") rather than scaling one.

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
    font_height = _FONT_HEIGHT[normal[0]][normal[1]]
    descent = _DESCENT[normal[0]][normal[1]]
    ascent = font_height - descent
    top_inset = _LINE_INSET[normal[0]][normal[1]]  # the largest inset of the row's unwrapped fonts
    for font in cell_fonts:
        key, size = _row_face(font)
        face_descent = _DESCENT[key][size]
        ascent = max(ascent, _FONT_HEIGHT[key][size] - face_descent)
        descent = max(descent, face_descent)
        top_inset = max(top_inset, _LINE_INSET[key][size])
    wraps = []
    for wrap_font, lines in wrapped:
        if lines < 1:
            msg = f"wrapped text has at least one line, not {lines!r}"
            raise ValueError(msg)
        key, size = normal if wrap_font is None else _row_face(wrap_font)
        wraps.append((key, size, lines))
        # the wrapped cell's font is in the row too: one line of Aptos Narrow 9 under Arial 11
        # makes the row 15 points, as the two faces unwrapped do
        face_descent = _DESCENT[key][size]
        ascent = max(ascent, _FONT_HEIGHT[key][size] - face_descent)
        descent = max(descent, face_descent)
    height = float(math.ceil(ascent + descent))
    for key, size, lines in wraps:
        inset = max(0.0, top_inset - _LINE_INSET[key][size])
        height = max(height, lines * (_FONT_HEIGHT[key][size] + 1.0) + inset)
    return RowMetrics(height=height, font_height=font_height, line_height=font_height + 1.0)


# --- vertical placement (CAPTURING.md section 2 item 47) ----------------------------------------

# family (lower case) -> size -> points between a bottom-aligned baseline and the row's bottom
_BASELINE_FROM_BOTTOM: dict[str, dict[int, float]] = {
    "calibri": {9: 3.0, 10: 3.0, 11: 3.0, 12: 3.0, 14: 4.0, 16: 4.0, 18: 5.0, 20: 5.0, 24: 6.0},
    "aptos narrow": {
        9: 3.0,
        10: 3.0,
        11: 3.0,
        12: 3.0,
        14: 4.0,
        16: 5.0,
        18: 5.0,
        20: 6.0,
        24: 7.0,
    },
}
# family (lower case) -> size -> (cap height, descender): the ink of "H" above the baseline and of
# "g" below it, as Excel draws them at zoom 100 on a 2x display (half points)
_INK: dict[str, dict[int, tuple[float, float]]] = {
    "calibri": {
        9: (6.0, 1.5), 10: (6.5, 2.0), 11: (7.0, 2.0), 12: (8.0, 2.0), 14: (9.0, 2.5),
        16: (10.5, 3.0), 18: (11.5, 3.0), 20: (13.0, 3.5), 24: (15.5, 4.5),
    },
    "aptos narrow": {
        9: (6.0, 1.5), 10: (6.5, 2.0), 11: (7.5, 2.0), 12: (8.0, 2.0), 14: (9.5, 2.5),
        16: (10.5, 3.0), 18: (12.0, 3.5), 20: (13.5, 3.5), 24: (16.0, 4.5),
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
            text (Excel's default). Centered text sits midway between the two placements, to half
            a point.
        font_height: The height of a row in this face alone (no stored height).
        line_height: What each wrapped line in this face adds to a row: ``font_height`` plus 1.
        line_pitch: The distance between the baselines of wrapped lines: ``font_height``; the
            last line is placed as bottom-aligned text is.
        line_inset: The face's line inset, the ``p`` of the row model. Wrapped text in this face,
            in a row whose Normal font or other cell fonts have a larger ``line_inset``, takes
            that difference more height (``row_metrics``) and sits that much higher, the extra
            lying below its last line.
        cap_height: The ink of "H" above the baseline, to half a point.
        descender: The ink of "g" below the baseline, to half a point.
    """

    baseline_from_top: float
    baseline_from_bottom: float
    font_height: float
    line_height: float
    line_pitch: float
    line_inset: float
    cap_height: float
    descender: float


def vertical_metrics(font: FontSpec | None = None) -> VerticalMetrics:
    """Where Excel draws text in ``font`` within a cell (CAPTURING.md section 2 item 47).

    Bold and italic place as the regular face does, and the workbook's Normal font moves nothing,
    except for wrapped text in a face with a smaller ``line_inset`` (see
    :class:`VerticalMetrics`). Measured: Calibri and Aptos Narrow at 9, 10, 11, 12, 14, 16, 18,
    20 and 24 points; any other face raises ``LookupError`` ("this font isn't measured"),
    Arial included: its widths and row heights are measured, its placement in the row not yet.

    Raises:
        LookupError: A font or size not measured.
    """
    key, size = _row_face(font)
    height = _FONT_HEIGHT[key][size]
    cap, descender = _INK[key][size]
    return VerticalMetrics(
        baseline_from_top=float(size) + 1.0,
        baseline_from_bottom=_BASELINE_FROM_BOTTOM[key][size],
        font_height=height,
        line_height=height + 1.0,
        line_pitch=height,
        line_inset=_LINE_INSET[key][size],
        cap_height=cap,
        descender=descender,
    )
