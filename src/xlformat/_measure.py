"""Measure functions and the column model for the width decision (``ARCHITECTURE.md`` §8).

Character mode measures a text by its length and the budget is the width in characters. Excel mode
measures a text as the sum of its glyphs' whole advances in the device units of the cell's font: on
the Mac each glyph's ``u * size / E`` rounded half up to whole points, for the 27 families of
``_faces_mac`` (CAPTURING.md section 2 item 82); on Windows the whole pixels of ``_faces_windows``
at the size's ppem (item 85). The budget is the column's text area: Excel for Mac stores a column
of ``w`` units as ``round(MDW * w + P)`` points, MDW being the maximum digit width of the Normal
font (Calibri's is 6 points at 11 and 12) and ``P`` the padding of :func:`column_padding` (5 points
for a digit width up to 8, 7 for 9 to 12, 9 for 13), and compares text against those points less
the padding. Both modes feed the same seam in ``_assemble``: fit when ``measure(text) <= budget``,
``#`` and ``*`` counts by dividing the budget by the glyph's measure (CAPTURING.md section 2 items
24 and 25).
"""

from __future__ import annotations

import math
import warnings
from collections.abc import Callable, Iterable
from dataclasses import dataclass, replace
from functools import cache, lru_cache

from xlformat import _faces_mac, _faces_windows
from xlformat._platform import _MAC, Platform, check_platform
from xlformat._rows_windows import WINDOWS_ROWS
from xlformat._rules_drawn import COURIER_GAP, DRAWN
from xlformat._screen_windows import WINDOWS_INK, WINDOWS_RULES

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

# Excel for Windows at 96 dpi: every face of _faces_windows.FACES in whole pixels per ppem
# (CAPTURING.md section 2 item 85)
_WINDOWS_FAMILIES = {family.lower(): family for family in _faces_windows.FACES}
# Excel for Mac: every face of _faces_mac.FACES at any size, by rule (CAPTURING.md section 2
# item 82)
_STYLES = ("regular", "bold", "italic", "bold italic")
_MAC_FAMILIES = {family.lower(): family for family in _faces_mac.FACES}
_MAC_SUBSTITUTES = {name.lower(): family for name, family in _faces_mac.SUBSTITUTES.items()}
MAC_VERIFIED_SIZE = 192.0
"""The largest size at which Excel for Mac's rule is captured: up to it each glyph's advance is
rounded to whole points (CAPTURING.md section 2 item 82). Larger sizes, to Excel's 409 points,
are answered by the same rule, unverified; there autofit stops rounding each glyph at a size
that depends on the face (Calibri from 286 points; Arial, Times New Roman, Verdana, Candara and
Georgia still round at 320), while fit and ``#`` decisions still round each glyph (digits and
commas in Calibri and Arial at 194 to 318 points, CAPTURING.md section 2 item 88)."""
_DIGITS = "0123456789"


def _mac_size(size: float) -> float:
    """The size Excel for Mac measures a face at: a fractional size rounds half up (10.5 is 11)."""
    return float(math.floor(size + 0.5))


@cache
def _mac_face(family: str, style: str) -> tuple[int, dict[str, float]]:
    """A Mac face's units per em E and each glyph's advance in 1/E em."""
    entry = _faces_mac.FACES[family][style]
    while isinstance(entry, str):
        entry = _faces_mac.FACES[family][entry]
    e, numbers = entry
    return e, dict(zip(_faces_mac.GLYPHS, map(float, numbers.split()), strict=True))


@cache
def _mac_overhang(family: str, style: str, size: float) -> float:
    """An italic Mac face's overhang at a whole size: ``floor(a * S + b)`` (item 87); else 0."""
    rule = _faces_mac.OVERHANG.get((family, style))
    if rule is None:
        return 0.0
    a_num, a_den, b_num, b_den = rule
    return float((a_num * b_den * int(size) + b_num * a_den) // (a_den * b_den))


def column_padding(mdw: float) -> float:
    """Points Excel adds to ``MDW * width`` for a column, not available to text.

    ``2 * ceil(MDW / 4) + 1``: 5 for a maximum digit width of 5 to 8 points, 7 for 9 to 12, 9
    for 13 (Arial 24; CAPTURING.md section 2 items 48 and 55).
    """
    return 2.0 * math.ceil(mdw / 4) + 1.0


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


def column_units(
    column_width: float, font: FontSpec | None = None, platform: Platform = _MAC
) -> float:
    """The column width Excel keeps after snapping to whole device units (8.43 becomes 8.5)."""
    return _units(column_points(column_width, font, platform), _mdw(font, platform))


def resolve_font(font: FontSpec | None, platform: Platform = _MAC) -> Font:
    """The canonical :class:`Font`; raises LookupError for a font, style or size not covered.

    On the Mac a name Excel for Mac has no face for and measures as another (``SUBSTITUTES``)
    resolves to that face; every size is measured by the rule (verified to
    :data:`MAC_VERIFIED_SIZE` points).
    """
    if font is None:
        return DEFAULT_FONT
    if isinstance(font, str):
        font = Font(font, 11.0)
    elif not isinstance(font, Font):
        font = Font(font[0], float(font[1]))
    if platform.name == "mac":
        key = font.name.lower()
        family = _MAC_FAMILIES.get(key) or _MAC_SUBSTITUTES.get(key)
        if family is None:
            msg = (
                f"no glyph advances for font {font.name!r}; measured: {', '.join(_faces_mac.FACES)}"
            )
            raise LookupError(msg)
        return Font(family, float(font.size), font.bold, font.italic, font.underline)
    display = _WINDOWS_FAMILIES.get(font.name.lower())
    if display is None:
        bundled = ", ".join(_faces_windows.FACES)
        msg = f"no glyph table for Excel for Windows for font {font.name!r}; measured: {bundled}"
        raise LookupError(msg)
    return Font(display, float(font.size), font.bold, font.italic, font.underline)


def _table(font: FontSpec | None, platform: Platform = _MAC) -> dict[str, float]:
    f = resolve_font(font, platform)
    return _whole_advances(f.name, f.size, f.style, platform.name)


@lru_cache(maxsize=128)
def _whole_advances(name: str, size: float, style: str, kind: str = "mac") -> dict[str, float]:
    """The advances every Excel-mode decision sums: whole points (pixels on Windows).

    On the Mac each glyph's advance ``u * size / E`` rounded half up (CAPTURING.md section 2
    item 82). On Windows the whole pixels measured at the size's ppem (item 85). A line feed and
    a tab take the space's advance and a carriage return none, so CR LF measures one space
    (item 51).
    """
    if kind == "mac":
        e, units = _mac_face(name, style)
        size = _mac_size(size)
        whole = {c: float(math.floor(u * size / e + 0.5)) for c, u in units.items()}
    else:
        whole = dict(_advances(name, size, style))
    whole["\n"] = whole["\t"] = whole[" "]
    whole["\r"] = 0.0
    # a glyph outside the table: the widest digit (at least a point: a column unit divides by it)
    whole[""] = max(1.0, *(whole[d] for d in _DIGITS))
    return whole


def _windows_ppem(size: float) -> int:
    """The ppem a size draws at in Excel for Windows at 96 dpi, a half pixel rounded up."""
    return math.floor(size * 4 / 3 + 0.5)


@lru_cache(maxsize=256)
def _windows_glyphs(family: str, style: str, ppem: int) -> dict[str, float]:
    """A Windows face's whole-pixel advances at one measured ppem, decoded on first use."""
    entry = _faces_windows.FACES[family][style]
    if isinstance(entry, str):  # a style with the same advances
        entry = _faces_windows.FACES[family][entry]
    assert not isinstance(entry, str)
    text = entry[_faces_windows.PPEMS.index(ppem)]
    return dict(zip(_faces_mac.GLYPHS, map(float, text.split()), strict=True))


@lru_cache(maxsize=128)
def _advances(name: str, size: float, style: str = "regular") -> dict[str, float]:
    """A Windows face at ``size``: its measured ppem, or the nearest measured one scaled."""
    ppem = _windows_ppem(size)
    if ppem in _faces_windows.PPEMS:
        return _windows_glyphs(name, style, ppem)
    # An unmeasured ppem: scale the nearest measured one (the larger on a tie) and round each
    # advance to whole pixels, the way hinting does. An approximation; the measured ppems are
    # exact.
    nearest = min(_faces_windows.PPEMS, key=lambda measured: (abs(measured - ppem), -measured))
    base = _windows_glyphs(name, style, nearest)
    return {ch: float(round(adv * ppem / nearest)) for ch, adv in base.items()}


def overhang(font: FontSpec | None = None, *, platform: str | Platform | None = None) -> float:
    """Points an italic face reaches past its advances on each side; 0 for upright faces.

    The text area of a cell in that font is this much smaller: an italic Calibri 11 cell holds
    one fill space fewer than a regular one. On the Mac one rule per face, ``floor(a * S + b)``
    points at the whole size S (a half size rounds half up), captured to hold at every whole size
    from 6 to 192 and unverified outside them. On Windows measured at every ppem from 11 to 32
    (sizes 8 to 24 and the half sizes between); another size scales the nearest measured
    overhang to whole pixels.
    """
    p = check_platform(platform, "overhang", windows=True)
    return font_overhang(font, p) / p.scale


def font_overhang(font: FontSpec | None, platform: Platform = _MAC) -> float:
    """:func:`overhang` in the platform's device units."""
    f = resolve_font(font, platform)
    return _overhang(f.name, f.size, f.style, platform.name)


@lru_cache(maxsize=128)
def _overhang(name: str, size: float, style: str, kind: str = "mac") -> float:
    if kind == "mac":
        return _mac_overhang(name, style, _mac_size(size))
    size = float(_windows_ppem(size))
    values = _faces_windows.OVERHANG.get((name, style))
    table = (
        {}
        if values is None
        else {float(p): float(v) for p, v in zip(_faces_windows.PPEMS, values.split(), strict=True)}
    )
    if not table:
        return 0.0
    if size in table:
        return float(table[size])
    nearest = min(table, key=lambda measured: (abs(measured - size), -measured))
    return float(round(table[nearest] * size / nearest))


def glyph_measure(font: FontSpec | None = None, platform: Platform = _MAC) -> Measure:
    """Excel mode: the sum of whole glyph advances in device units (points on the Mac).

    A line feed and a tab measure as a space, a carriage return as nothing, and a glyph
    outside the table (beyond the cp1252 characters) as the widest digit (the maximum digit
    width, which is what a column width unit is).
    """
    whole = _table(font, platform)
    digit = whole[""]

    def measure(text: str) -> float:
        return sum(whole.get(c, digit) for c in text)

    return measure


def text_width(
    text: str, font: FontSpec | None = None, *, platform: str | Platform | None = None
) -> float:
    """The width of ``text`` in points as Excel measures it in ``font`` on ``platform``.

    The sum of the whole glyph advances, the measure every Excel-mode decision uses (fit,
    ``#`` count, fill count); a line feed and a tab measure as a space, a carriage return as
    nothing, and a glyph beyond the cp1252 characters as the widest digit. Like reportlab's
    ``stringWidth``, but whole points as Excel rounds them. Excel's autofit alone measures
    differently: a text holding a non-ASCII glyph unrounded, and a long text a point or more
    wider.

    Raises:
        LookupError: A font, style or size xlformat has no advances for
            (:func:`measured_fonts`).
    """
    p = check_platform(platform, "text_width", windows=True)
    return glyph_measure(font, p)(text) / p.scale


def glyph_widths(
    text: str, font: FontSpec | None = None, *, platform: str | Platform | None = None
) -> tuple[float, ...]:
    """The advance of each character of ``text`` in whole points.

    For measuring runs of a text (words, spaces): ``text_width`` of a run is
    ``sum(glyph_widths(run))``.
    """
    p = check_platform(platform, "glyph_widths", windows=True)
    whole = _table(font, p)
    return tuple(whole.get(c, whole[""]) / p.scale for c in text)


def digit_width(font: FontSpec | None = None, *, platform: str | Platform | None = None) -> float:
    """The maximum digit width (MDW) of ``font`` in points.

    One column unit of a workbook whose Normal font is ``font``, and one indent level.
    """
    p = check_platform(platform, "digit_width", windows=True)
    return _mdw(font, p) / p.scale


def measured_fonts(
    *, platform: str | Platform | None = None
) -> dict[str, dict[str, tuple[int, ...]]]:
    """The fonts xlformat measures text in: family, then style, then the sizes verified.

    On the Mac every listed family and style is measured at any size by its glyph advances,
    recovered from Excel's autofit widths: captured to 192 points, and by the same rule, unverified,
    above (Excel allows 409). The sizes listed are those at which every glyph's width was captured.
    :func:`font_substitutes` lists names Excel for Mac measures as one of these. On Windows each
    listed family and style is measured in whole pixels at every ppem from 11 to 32 at 96 dpi:
    the sizes listed, 8 to 24, and the half sizes between; another size scales the nearest
    measured ppem and may differ from Excel by a pixel a glyph. Anything else raises
    ``LookupError``. A caller can test a workbook's fonts against this before choosing Excel
    mode. Text measurement and, on the Mac, :func:`text_insets` follow this list; row heights,
    vertical placement and rules cover fewer faces (each function says which).
    """
    p = check_platform(platform, "measured_fonts", windows=True)
    if p.name == "mac":
        return {
            family: dict.fromkeys(_STYLES, _faces_mac.VERIFIED_SIZES[family])
            for family in _faces_mac.FACES
        }
    sizes = tuple(s for s in range(1, 410) if _windows_ppem(s) in _faces_windows.PPEMS)
    return {family: dict.fromkeys(_STYLES, sizes) for family in _faces_windows.FACES}


def font_substitutes(*, platform: str | Platform | None = None) -> dict[str, str]:
    """Font names the platform has no face for, and the face Excel measures them in.

    Excel for Mac measures MS Sans Serif as Times New Roman (the Mac has no such face; every
    captured text autofits alike, CAPTURING.md section 2 item 82), and so does every function
    here that takes ``font=`` on the Mac. Segoe UI, also absent, autofits as none of the
    measured faces and raises ``LookupError``. Empty on Windows.
    """
    p = check_platform(platform, "font_substitutes", windows=True)
    return dict(_faces_mac.SUBSTITUTES) if p.name == "mac" else {}


def column_budget(
    column_width: float, font: FontSpec | None = None, platform: Platform = _MAC
) -> float:
    """Device units available to a number or date in a column of ``column_width`` units."""
    return column_points(column_width, font, platform) - column_padding(_mdw(font, platform))


def column_points(
    column_width: float, font: FontSpec | None = None, platform: Platform = _MAC
) -> float:
    """A column's width in whole device units: points on the Mac, pixels on Windows.

    ``round(MDW * w + P)``, and below one unit ``round(w * (MDW + P))``, a half rounding up,
    ``P`` being
    :func:`column_padding` (CAPTURING.md section 2 items 28 and 48).
    """
    mdw = _mdw(font, platform)
    if column_width < 1:
        exact = column_width * (mdw + column_padding(mdw))
    else:
        exact = mdw * column_width + column_padding(mdw)
    # a half rounds up on both platforms: 12.5 points is 13 at Calibri 9 and width 1.5, 36.5
    # pixels 37 at a digit of 7 pixels and width 4.5 (CAPTURING.md section 2 items 66 and 67)
    return float(math.floor(exact + 0.5))


def _mdw(font: FontSpec | None, platform: Platform = _MAC) -> float:
    return _table(font, platform)[""]


def _units(points: float, mdw: float) -> float:
    """Dialog units of a column of ``points``, exact (Excel shows them to two decimals)."""
    pad = column_padding(mdw)
    if points < mdw + pad:
        return points / (mdw + pad)
    return (points - pad) / mdw


def file_points(
    file_width: float, font: FontSpec | None = None, platform: Platform = _MAC
) -> float:
    """Device units of a column whose ``<col width>`` is ``file_width``.

    ECMA-376 Part 1 section 18.3.1.13's pixel formula with the maximum digit width in points,
    the width first truncated to 1/256 as Excel stores it:
    ``trunc((256 * w + trunc(128 / MDW)) / 256 * MDW)`` (CAPTURING.md section 2 item 28).
    """
    mdw = _mdw(font, platform)
    w = math.trunc(file_width * 256) / 256
    return float(math.trunc((256 * w + math.trunc(128 / mdw)) / 256 * mdw))


@dataclass(frozen=True, slots=True)
class ColumnMetrics:
    """A column as Excel lays it out, in whole points on the Mac and whole pixels on Windows.

    Attributes:
        width: The width in Column Width dialog units, exact (Excel shows it to two decimals:
            8.43 becomes 8.5 at Calibri 11); pass it as ``column_width``, unless it is 0 (a
            column of no width, which ``column_width`` refuses with ``ValueError``).
        points: The column's width in points (53 for an openpyxl column without a width).
        budget: Points a number or date may take before it becomes a ``#`` run: ``points``
            less the column padding (5 points for a digit width up to 8, 7 for 9 to 12, 9 for
            13), and 0
            for a column narrower than that.
        left_inset: Where left-aligned text in the Normal font starts at zoom 100, from the
            cell's left boundary: the gridline there (one screen unit: a point on the Mac, a pixel
            on Windows) and ``ceil(digit_width / 4)`` screen units of clear space, as
            :func:`text_insets` gives it for that font.
        right_inset: Where right-aligned text ends, before the cell's right boundary: the clear
            space alone, ``ceil(digit_width / 4)`` screen units (the gridline there is the next
            cell's).
        digit_width: The maximum digit width of the Normal font: one column unit, one indent
            level.

    The insets and the gridline between them make the column padding: a text in the Normal font
    fits the column when its width is at most ``points`` less the padding (CAPTURING.md section
    2 item 81). A cell in another font has its own insets (:func:`text_insets`).
    """

    width: float
    points: float
    budget: float
    left_inset: float
    right_inset: float
    digit_width: float


def _metrics(device: float, font: FontSpec | None, platform: Platform = _MAC) -> ColumnMetrics:
    """The column ``device`` units wide, reported in points."""
    mdw = _mdw(font, platform)
    scale = platform.scale
    return ColumnMetrics(
        width=_units(device, mdw),
        points=device / scale,
        budget=max(device - column_padding(mdw), 0.0) / scale,
        left_inset=(1 + math.ceil(mdw / 4)) / scale,
        right_inset=math.ceil(mdw / 4) / scale,
        digit_width=mdw / scale,
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
    platform: str | Platform | None = None,
) -> ColumnMetrics:
    """The column of ``column_width`` dialog units in the workbook's Normal font.

    Args:
        column_width: The width in Column Width dialog units.
        normal_font: The workbook's Normal font, whose digit width sets the column's units.
        font: Deprecated name of ``normal_font``, removed in 1.0.
        platform: Where the sheet is viewed: ``"mac"``, ``"windows"`` or a :class:`Platform`;
            ``None`` means the default, ``"mac"``. Excel for Windows is measured at 96 dpi;
            another resolution raises ``LookupError``. Results are in points on both.

    Raises:
        ValueError: ``column_width`` is not positive, or both font arguments were passed.
    """
    p = check_platform(platform, "column_metrics", windows=True)
    normal_font = _normal_font_argument(normal_font, font)
    if column_width <= 0:
        msg = f"column_width must be positive, not {column_width!r}"
        raise ValueError(msg)
    return _metrics(column_points(column_width, normal_font, p), normal_font, p)


def column_metrics_from_points(
    points: float, normal_font: FontSpec | None = None, *, platform: str | Platform | None = None
) -> ColumnMetrics:
    """The column ``points`` wide, in the workbook's Normal font.

    For a merged cell, pass the sum of its columns' ``points``: Excel formats it as one column of
    that width, with a single padding (CAPTURING.md section 2 item 40), and a hidden column counts
    as 0. The result's ``width`` is what ``column_width`` takes, unless it is 0.

    Raises:
        ValueError: ``points`` is negative.
    """
    p = check_platform(platform, "column_metrics_from_points", windows=True)
    if points < 0:
        msg = f"points must not be negative, not {points!r}"
        raise ValueError(msg)
    return _metrics(float(points) * p.scale, normal_font, p)


def file_column_metrics(
    file_width: float | None = None,
    *,
    base_col_width: float = 8,
    default_col_width: float | None = None,
    normal_font: FontSpec | None = None,
    font: FontSpec | None = None,
    platform: str | Platform | None = None,
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
    p = check_platform(platform, "file_column_metrics", windows=True)
    normal_font = _normal_font_argument(normal_font, font)
    if file_width is None:
        if default_col_width is None:
            mdw = _mdw(normal_font, p)
            # Windows rounds it up to a multiple of 8 pixels (default_column_step)
            device = base_col_width * mdw + column_padding(mdw)
            step = p.default_column_step
            if step > 1:
                device = float(math.ceil(device / step) * step)
            return _metrics(device, normal_font, p)
        file_width = default_col_width
    if file_width < 0:
        msg = f"file_width must not be negative, not {file_width!r}"
        raise ValueError(msg)
    return _metrics(file_points(file_width, normal_font, p), normal_font, p)


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


# family (lower case) -> post.underlineThickness / unitsPerEm. An accounting underline adds its
# line's thickness and its double's gap to the row and to each wrapped line (CAPTURING.md section
# 2 items 58 and 63): the thickness max(1, floor(size * this)), the gap two thicknesses, except
# where the drawn double is wider (Times New Roman from 37 to 40 points, Courier New always)
_ACCOUNTING_THICKNESS: dict[str, float] = {
    "calibri": 134 / 2048,
    "aptos narrow": 50 / 2048,
    "arial": 150 / 2048,
    "times new roman": 100 / 2048,
    "georgia": 101 / 2048,
}
# Arial's double accounting lines, as drawn, are 1 point thick up to 35 points, then 2, then 3
# from 55 (thinner than its single underline's; CAPTURING.md section 2 item 62)
_ARIAL_DOUBLE_LINE = ((55, 3), (36, 2), (0, 1))


def _accounting(key: str, size: float, underline: str | None) -> float:
    """Points an accounting underline adds to each line of the face; LookupError if not known."""
    if underline not in ("singleAccounting", "doubleAccounting"):
        return 0.0
    whole = math.ceil(size)
    known = (key in _ACCOUNTING_THICKNESS or key == "courier new") and 6 <= whole <= _LARGEST
    if not known:
        msg = f"no accounting underline rows for {_DISPLAY[key]} {size:g}"
        raise LookupError(msg)
    if key == "courier new":
        return max(1, math.floor(whole * 84 / 2048)) + COURIER_GAP[whole]
    line = max(1, math.floor(whole * _ACCOUNTING_THICKNESS[key]))
    if key == "arial" and underline == "doubleAccounting":
        line = next(t for start, t in _ARIAL_DOUBLE_LINE if whole >= start)
    wide = key == "times new roman" and _TIMES_WIDE_GAP[0] <= whole <= _TIMES_WIDE_GAP[1]
    return float(line + (3 if wide else 2) * line)


def _bottom(key: str, size: float) -> float:
    """B, the face's descent: its height less its ascent, ceil(size) + 1 points."""
    return _font_height(key, size) - math.ceil(size) - 1.0


@dataclass(frozen=True, slots=True)
class RowMetrics:
    """A row as Excel fits it when it opens a file, in points.

    Whole points on the Mac, whole pixels on Windows (0.75 points at 96 dpi).

    Attributes:
        height: The row's height: what Excel gives a row without a stored height (no ``ht``)
            holding the fonts and wrapped text given.
        font_height: The Normal font's height: a row in the Normal font alone, and the sheet's
            standard height.
        line_height: What each wrapped line in the Normal font takes: on the Mac
            ``font_height`` plus 1, so one wrapped line is a point taller than the row without
            it; on Windows ``font_height``.
        accounting_height: What the accounting underlines of the row's fonts add to
            ``height``: the difference from the same row without them (0 without one).
    """

    height: float
    font_height: float
    line_height: float
    accounting_height: float = 0.0


def _row_face(font: FontSpec | None) -> tuple[str, float]:
    """(family key, size) of a face the row model has measured; LookupError otherwise.

    A name Excel for Mac has no face for and measures as another (``font_substitutes``) is that
    face: MS Sans Serif is Times New Roman.
    """
    if font is None:
        f = DEFAULT_FONT
    elif isinstance(font, str):
        f = Font(font, 11.0)
    elif isinstance(font, Font):
        f = font
    else:
        f = Font(font[0], float(font[1]))
    key = f.name.lower()
    key = _MAC_SUBSTITUTES.get(key, key).lower()
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
    platform: str | Platform | None = None,
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

    An accounting underline (``Font.underline``) adds its line's thickness and its double's gap
    to the row and to each wrapped line, at the whole size ``ceil(size)`` from 6 to 72 points
    (CAPTURING.md section 2 items 58, 61 and 63): mostly three thicknesses, a thickness being
    ``max(1, floor(size * underlineThickness / em))`` and Arial's double accounting taking its
    thinner double lines; 4 points in Times New Roman from 36.5 to 40, where its double is three
    thicknesses wide; in Courier New its captured gap and its thickness.

    Several cells with wrapped text in one row are each computed this way and the row takes the
    tallest; that combination has not been captured.

    Args:
        normal_font: The workbook's Normal font.
        cell_fonts: The fonts of the row's cells that are not the Normal font and do not wrap.
        wrapped: One ``(font, lines)`` per cell with wrapped text: its font (``None`` for the
            Normal font) and its number of lines, as the caller broke them.
        platform: Where the sheet is viewed: ``"mac"``, ``"windows"`` or a :class:`Platform`;
            ``None`` means the default, ``"mac"``. Excel for Windows is measured at 96 dpi in
            Calibri at 6 to 40 points and the half sizes between and at 41 to 72, and in Aptos
            Narrow, Arial, Times New Roman and Courier New at 6 to 40, four styles each,
            without accounting underlines; anything else raises ``LookupError`` there.

    Raises:
        LookupError: A font or size the row model has not measured.
        ValueError: A number of lines below 1.
    """
    p = check_platform(platform, "row_metrics", windows=True)
    if p.name == "windows":
        return _windows_row(normal_font, cell_fonts, wrapped, p)
    cell_fonts, wrapped = list(cell_fonts), list(wrapped)
    row = _mac_row(normal_font, cell_fonts, wrapped, accounting=True)
    plain = _mac_row(normal_font, cell_fonts, wrapped, accounting=False)
    return replace(row, accounting_height=row.height - plain.height)


def _mac_row(
    normal_font: FontSpec | None,
    cell_fonts: list[FontSpec],
    wrapped: list[tuple[FontSpec | None, int]],
    *,
    accounting: bool,
) -> RowMetrics:
    """Excel for Mac's row; ``accounting`` False leaves accounting underlines out."""
    normal = _row_face(normal_font)
    font_height = _font_height(*normal)

    def underline(font: FontSpec | None) -> str | None:
        return font.underline if accounting and isinstance(font, Font) else None

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


def _windows_face(font: FontSpec | None) -> tuple[int, int]:
    """(H, D) in pixels of a face Excel for Windows' rows were captured in; LookupError else."""
    f = DEFAULT_FONT if font is None else resolve_font(font, Platform.WINDOWS)
    if isinstance(font, Font) and font.underline in ("singleAccounting", "doubleAccounting"):
        msg = f"no accounting underline rows for Excel for Windows ({f.name} {f.size:g})"
        raise LookupError(msg)
    sizes = WINDOWS_ROWS.get(f.name.lower(), {}).get(f.style, {})
    if f.size not in sizes:
        measured = ", ".join(f"{z:g}" for z in sizes)
        msg = f"no row heights for Excel for Windows for {f.name} {f.size:g}; measured: {measured}"
        raise LookupError(msg)
    return sizes[f.size]


def _windows_row(
    normal_font: FontSpec | None,
    cell_fonts: Iterable[FontSpec],
    wrapped: Iterable[tuple[FontSpec | None, int]],
    platform: Platform,
) -> RowMetrics:
    """Excel for Windows' row: ``max(lines * H - D) + max D`` pixels over its cells.

    The Normal font is a one-line cell among them; H and D per face (CAPTURING.md section 2
    item 69).
    """
    normal = _windows_face(normal_font)
    cells = [(normal, 1), *((_windows_face(f), 1) for f in cell_fonts)]
    for font, lines in wrapped:
        if lines < 1:
            msg = f"wrapped text has at least one line, not {lines!r}"
            raise ValueError(msg)
        cells.append((normal if font is None else _windows_face(font), lines))
    top = max(lines * h - d for (h, d), lines in cells)
    pixels = top + max(d for (_, d), _ in cells)
    scale = platform.scale
    font_height = normal[0] / scale
    return RowMetrics(height=pixels / scale, font_height=font_height, line_height=font_height)


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


@dataclass(frozen=True)
class VerticalMetrics:
    """Where Excel draws a face's text in a cell, in sheet points (whole pixels on Windows).

    Every field is a distance Excel's screen shows, not a split of one by convention: the top
    inset and the ascent are seen only as their sum, ``baseline_from_top``, and the descent and
    the bottom inset as ``baseline_from_bottom``.

    Attributes:
        baseline_from_top: The baseline's distance below the row's top for top-aligned text: on
            the Mac the face's size in points plus 1; on Windows ``H - D + 1`` pixels, the face's
            height and descent of :func:`row_metrics`. A row's top is the top of the gridline
            above it, which the row's own fill covers.
        baseline_from_bottom: The baseline's distance above the row's bottom for bottom-aligned
            text (Excel's default). It is at least the Normal font's own, and includes that
            when :func:`vertical_metrics` is given ``normal_font``: on the Mac Calibri 9 and Arial
            9 to 11 have 2 points of their own and sit 3 up in a Calibri 11 or Aptos Narrow 11
            workbook, 2 in an Arial 10 one; on Windows it is ``max(D, D') - 1`` pixels, ``D'``
            the Normal font's descent. Centered text sits midway between the two placements, to
            half a point on the Mac; on Windows ``floor((R - H) / 2)`` pixels below top-aligned
            text, in a row of ``R`` pixels.
        font_height: The height of a row in this face alone (no stored height).
        line_height: What each wrapped line in this face adds to a row: on the Mac
            ``font_height`` plus 1, on Windows ``font_height`` (Calibri 11: 15 and 15 points).
        line_pitch: The distance between the baselines of wrapped lines: ``font_height``; the
            last line is placed as bottom-aligned text is.
        line_inset: The face's line inset, the ``p`` of the row model; on Windows ``D - 1``
            pixels. Wrapped text in this face, in a row whose Normal font or other cell fonts
            have a larger ``line_inset``, takes that difference more height (``row_metrics``)
            and sits that much higher, the extra lying below its last line.

    The ink of the text, which is measured rather than given by a rule, is
    :func:`ink_metrics`'s; ``cap_height`` and ``descender`` here are deprecated and forward to it.
    """

    # Slots, with one that is not a field: the font the deprecated ink properties read, set by
    # vertical_metrics after construction (not in __init__, fields(), asdict(), repr or
    # equality). A VerticalMetrics made otherwise, directly or by dataclasses.replace, has none.
    __slots__ = (
        "_font",
        "baseline_from_bottom",
        "baseline_from_top",
        "font_height",
        "line_height",
        "line_inset",
        "line_pitch",
    )  # fmt: skip

    baseline_from_top: float
    baseline_from_bottom: float
    font_height: float
    line_height: float
    line_pitch: float
    line_inset: float

    def __getstate__(self) -> dict[str, object]:
        return {name: getattr(self, name) for name in self.__slots__ if hasattr(self, name)}

    def __setstate__(self, state: dict[str, object]) -> None:
        for name, value in state.items():
            object.__setattr__(self, name, value)

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
        font = getattr(self, "_font", None)
        if font is None:
            msg = (
                f"this VerticalMetrics does not know its font, so it has no {name}: use "
                f"ink_metrics(font).{name}"
            )
            raise LookupError(msg)
        return ink_metrics(font)


@dataclass(frozen=True, slots=True)
class InkMetrics:
    """The ink Excel draws for a face at zoom 100, in sheet points.

    To half a point on the Mac's 2x display, to a pixel (0.75 points) on Windows at 96 dpi.

    Attributes:
        cap_height: The ink of "H" above the baseline.
        descender: The ink of "g" below the baseline. On the Mac, Arial 11 and 16, Times New
            Roman 11, 16 and 20 and Georgia 11 and 20 draw it half a point longer or shorter by
            where the baseline falls on the display's pixels.
    """

    cap_height: float
    descender: float


def ink_metrics(
    font: FontSpec | None = None, *, platform: str | Platform | None = None
) -> InkMetrics:
    """The measured ink of ``font``'s "H" and "g" (CAPTURING.md section 2 items 47, 55 and 59).

    Measured: Calibri, Aptos Narrow, Arial, Times New Roman, Courier New and Georgia at 9, 10, 11,
    12, 14, 16, 18, 20 and 24 points, and Calibri and Aptos Narrow at 13, 15, 17 and 22, in every
    style alike (styles captured for the first three). Times New Roman, Courier New and Georgia,
    and the sizes 13, 15, 17 and 22, come from item 59; the rest from items 47 and 55.

    With ``platform="windows"`` (96 dpi, item 74): Calibri, Aptos Narrow and Arial at 9, 10, 11,
    12, 14, 16, 18, 20 and 24 points, and bold, italic and bold italic at 11, 12 and 20, each
    style as measured, in whole pixels.

    Raises:
        LookupError: A face or size not measured.
    """
    _platform = check_platform(platform, "ink_metrics", windows=True)
    if _platform.name == "windows":
        return _windows_ink(font)
    key, size = _row_face(font)
    ink = _INK.get(key, {}).get(int(size)) if size == int(size) else None
    if ink is None:
        msg = (
            f"no ink measured for {_DISPLAY[key]} {size:g}; measured: the six faces at 9, 10, "
            "11, 12, 14, 16, 18, 20 and 24, and Calibri and Aptos Narrow at 13, 15, 17 and 22"
        )
        raise LookupError(msg)
    return InkMetrics(cap_height=ink[0], descender=ink[1])


def _windows_ink(font: FontSpec | None) -> InkMetrics:
    f = DEFAULT_FONT if font is None else resolve_font(font, Platform.WINDOWS)
    ink = WINDOWS_INK.get((_DISPLAY.get(f.name.lower(), f.name), f.style, int(f.size)))
    if ink is None or f.size != int(f.size):
        msg = (
            f"no ink measured on Excel for Windows for {f.name} {f.size:g} {f.style}; measured: "
            "Calibri, Aptos Narrow and Arial at 9, 10, 11, 12, 14, 16, 18, 20 and 24, and bold, "
            "italic and bold italic at 11, 12 and 20"
        )
        raise LookupError(msg)
    return InkMetrics(cap_height=ink[0] * _PX, descender=ink[1] * _PX)


_PX = 0.75  # points a device pixel at 96 dpi and zoom 100


def _windows_vertical(font: FontSpec | None, normal_font: FontSpec | None) -> VerticalMetrics:
    """Excel for Windows' placement in whole pixels from the face's H and D.

    CAPTURING.md section 2 items 74 and 76: the H stems' last ink row is ``H - D`` below a
    top-aligned row's top and ``max(D, D')`` above a bottom-aligned row's bottom, the baseline
    the edge below that row.
    """
    h, d = _windows_face(font)
    below = d if normal_font is None else max(d, _windows_face(normal_font)[1])
    return VerticalMetrics(
        baseline_from_top=(h - d + 1) * _PX,
        baseline_from_bottom=(below - 1) * _PX,
        font_height=h * _PX,
        line_height=h * _PX,
        line_pitch=h * _PX,
        line_inset=(d - 1) * _PX,
    )


def vertical_metrics(
    font: FontSpec | None = None,
    *,
    normal_font: FontSpec | None = None,
    platform: str | Platform | None = None,
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
        platform: Where the sheet is viewed: ``"mac"``, ``"windows"`` or a :class:`Platform`;
            ``None`` means the default, ``"mac"``. On ``"windows"`` (96 dpi) the placement is
            in whole pixels from the face's H and D, the height and descent of
            :func:`row_metrics` (CAPTURING.md section 2 items 74 and 76): the baseline ``H - D
            + 1`` pixels below a top-aligned row's top, ``max(D, D') - 1`` above a
            bottom-aligned row's bottom (``D'`` the Normal font's), centered text ``floor((R -
            H) / 2)`` pixels lower than top-aligned text in a row of ``R`` pixels; wrapped
            lines ``H`` apart, the last raised by the difference of ``line_inset`` (``D - 1``)
            as on the Mac. The faces and sizes of the Windows row tables are supported.

    Raises:
        LookupError: A font or size not measured.
    """
    _platform = check_platform(platform, "vertical_metrics", windows=True)
    if _platform.name == "windows":
        return _windows_vertical(font, normal_font)
    key, size = _row_face(font)
    height = _font_height(key, size)
    from_bottom = _bottom(key, size)
    if normal_font is not None:
        from_bottom = max(from_bottom, _bottom(*_row_face(normal_font)))
    metrics = VerticalMetrics(
        baseline_from_top=math.ceil(size) + 1.0,
        baseline_from_bottom=from_bottom,
        font_height=height,
        line_height=height + 1.0,
        line_pitch=height,
        line_inset=_bottom(key, size) - 2.0,
    )
    object.__setattr__(metrics, "_font", Font(_DISPLAY[key], size))
    return metrics


# --- rules: underlines and strikethrough (CAPTURING.md section 2 item 60) ------------------------

# family (lower case) -> (post.underlinePosition, post.underlineThickness, OS/2.yStrikeoutPosition,
# OS/2.yStrikeoutSize, OS/2.usWinDescent), in units of 2048 an em (capture/font_tables.py)
_RULE_TABLES: dict[str, tuple[int, int, int, int, int]] = {
    "calibri": (-232, 134, 512, 134, 550),
    "aptos narrow": (-100, 50, 731, 102, 563),
    "arial": (-217, 150, 530, 102, 434),
    "times new roman": (-223, 100, 530, 102, 443),
    "georgia": (-181, 101, 594, 101, 449),
    "courier new": (-477, 84, 530, 102, 615),
}
# Times New Roman's double underline lies three line thicknesses apart at screen sizes 37 to 40
# (36.5 included, which rounds to 37), where its accounting rows add 4 points (CAPTURING.md
# section 2 item 63)
_TIMES_WIDE_GAP = (37, 40)


@dataclass(frozen=True, slots=True)
class RuleMetrics:
    """Where Excel draws a face's rules at a zoom, in sheet points from the baseline.

    On Windows each value is a whole number of device pixels at the zoom (``0.75 * 100 / zoom``
    points a pixel). A line may reach past the row's bottom: Excel draws it on into the row
    below, under that row's top gridline (CAPTURING.md section 2 item 80); a caller that draws
    within the row cuts it there.

    On the Mac each value is a whole number of screen points (half points above about 80 screen
    points, on the display of scale 2 it was captured on) at the screen size ``size * zoom /
    100``, divided back to sheet points.

    Attributes:
        underline_top: The top of a single underline, below the baseline.
        underline_thickness: The thickness of a single underline and of a single accounting one.
        double_gap: How far a double underline's second line lies below its first (both kinds).
        accounting_top: The top of an accounting underline (the first line of a double one),
            below the baseline.
        strikethrough_top: The top of a strikethrough, above the baseline.
        strikethrough_thickness: The thickness of a strikethrough.
        double_top: The top of a double underline's first line, below the baseline (negative:
            above it).
        double_thickness: The thickness of each line of a double underline, accounting or not.
    """

    underline_top: float
    underline_thickness: float
    double_gap: float
    accounting_top: float
    strikethrough_top: float
    strikethrough_thickness: float
    double_top: float
    double_thickness: float


_WINDOWS_RULE_FACES = {"calibri": "Calibri", "arial": "Arial", "aptos narrow": "Aptos Narrow"}
_CALIBRI_HIGHER = 171  # from this ppem Calibri's lines lie a pixel higher (item 79, item 80)


def _windows_rules(key: str, size: float, zoom: float) -> RuleMetrics:
    """Excel for Windows' lines at 96 dpi (CAPTURING.md section 2 items 75 and 77 to 80)."""
    face = _WINDOWS_RULE_FACES.get(key)
    if face is None:
        msg = (
            f"no rule positions on Excel for Windows for {_DISPLAY[key]}; measured: Calibri, "
            "Aptos Narrow and Arial"
        )
        raise LookupError(msg)
    if zoom != int(zoom) or not 10 <= zoom <= 400:
        msg = f"Excel for Windows draws at whole zooms from 10 to 400, not {zoom:g}"
        raise LookupError(msg)
    p = math.floor(size * 4 / 3 * zoom / 100 + 0.5)  # the ppem: pixels an em at 96 dpi
    if p < 1:
        msg = f"{_DISPLAY[key]} {size:g} at zoom {zoom:g} is under one pixel an em: not captured"
        raise LookupError(msg)

    def rounded(units: int) -> int:
        return math.floor(p * units / 2048 + 0.5)

    position, thickness, strike_position, strike_size, descent = _RULE_TABLES[key]
    line, top, accounting = max(1, rounded(thickness)), rounded(-position), rounded(descent)
    strike, strike_line = rounded(strike_position), max(1, rounded(strike_size))
    if key == "calibri" and p >= _CALIBRI_HIGHER:
        top, accounting, strike = top - 1, accounting - 1, strike + 1
    table = WINDOWS_RULES.get((face, p), (None,) * 8)
    if table[5] is not None:
        accounting = table[5]
    if key == "aptos narrow":  # the single's line stepped up until the double ends at accounting's
        first = top
        while first + 3 * line > accounting:
            first -= line
        double: tuple[int | None, int | None] = (first, line)
    else:
        double = (table[2], table[3])
    rule = (top, line, double[0], double[1], None, accounting, strike, strike_line)
    v = [r if t is None else t for r, t in zip(rule, table, strict=True)]
    if v[4] is None:
        v[4] = 2 * v[3]  # type: ignore[operator]
    # device pixels to sheet points: 0.75 at zoom 100
    u_top, u_thick, d_top, d_thick, gap, acc, s_top, s_thick = (x * 75 / zoom for x in v)  # type: ignore[operator]
    return RuleMetrics(
        underline_top=u_top,
        underline_thickness=u_thick,
        double_gap=gap,
        accounting_top=acc,
        strikethrough_top=s_top,
        strikethrough_thickness=s_thick,
        double_top=d_top,
        double_thickness=d_thick,
    )


def _courier_rule(q: int) -> tuple[int, ...]:
    """Courier New's lines at the whole screen size ``q``, in screen points."""
    gap = COURIER_GAP.get(q)
    if gap is None:
        lowest, highest = min(COURIER_GAP), max(COURIER_GAP)
        msg = (
            f"no rule positions for Courier New at screen size {q}: its double underline's "
            f"gap is measured at screen sizes {lowest} to {highest}"
        )
        raise LookupError(msg)
    g = int(gap)
    line = max(1, math.floor(q * 84 / 2048))
    top = math.floor(q * 477 / 2048)
    first = top - g + (1 if q == 12 else 0)  # the double alone is drawn a point lower at 12
    accounting = math.floor(q * 615 / 2048 + 0.5)
    single = top if g == 2 * line else first
    strike = (math.floor(q * 530 / 2048 + 0.5), max(1, math.floor(q * 102 / 2048 + 0.5)))
    return (single, line, first, line, g, accounting, *strike)


def rule_metrics(
    font: FontSpec | None = None, *, zoom: float = 100, platform: str | Platform | None = None
) -> RuleMetrics:
    """Where Excel draws underlines and a strikethrough in ``font`` at ``zoom`` percent.

    Excel draws the lines at the nearest whole screen size, ``P = round(size * zoom / 100)``
    (a half rounds up), and at a whole screen size they depend on ``P`` alone (CAPTURING.md
    section 2 items 62 and 63). Where the face and ``P`` were captured, each value is the line
    Excel drew: Calibri, Aptos Narrow, Arial, Times New Roman and Georgia at every ``P`` from 9
    to 36, every even one to 72 and 80 to 144 in steps of 8; Courier New at every ``P`` from 6
    to 144. Elsewhere, and for a line not read there, the value comes from the font's tables in
    whole screen points (item 60): an underline's top ``floor(P * -underlinePosition / em)``
    below the baseline and its thickness ``max(1, floor(P * underlineThickness / em))``; an
    accounting underline's top ``round(P * usWinDescent / em)``; a strikethrough's top
    ``round(P * yStrikeoutPosition / em)`` above the baseline, ``max(1, round(P *
    yStrikeoutSize / em))`` thick. A double underline's lines are ``max(1, floor(P *
    min(underlineThickness, yStrikeoutSize) / em))`` thick and two thicknesses apart; its first
    line starts at the single underline's top and steps up a thickness at a time until the
    double (three thicknesses) ends at or above the accounting underline's top.

    Where the double's gap is wider than two thicknesses, the double hangs up from the single's
    table position instead (its second line there, its first the gap above) and the single
    moves onto its first line: Times New Roman at ``P`` 37 to 40, three thicknesses apart, and
    Courier New, whose gap is a captured table with no closed form (2 up to ``P`` 17, then 3,
    rising to 19 at 144; at ``P`` 12 the double is drawn a point lower). A double accounting
    underline's second line lies the gap below its first.

    Known misses (QUIRKS.md): the rule misses some doubles at whole screen sizes not captured
    (Arial at most from 29, Times New Roman at some from 46, Calibri from 80, Georgia at 128
    and 144), and at a screen size that is not whole, Excel draws a line a screen point off
    in some rows (23 of 400 captured; Courier New's lines there also differ by zoom).

    With ``platform="windows"`` (96 dpi; Calibri, Aptos Narrow and Arial, at any whole zoom
    from 10 to 400) the lines are drawn at the ppem ``P = round(size * 4/3 * zoom / 100)`` (a
    half rounds up) and depend on ``P`` alone (CAPTURING.md section 2 items 77 to 80). In whole
    pixels, every rounding half up: an underline's top ``round(P * -underlinePosition / em)``
    below the baseline, its thickness ``max(1, round(P * underlineThickness / em))``; an
    accounting underline's top the face's hinted descent, ``round(P * usWinDescent / em)``
    (Arial's, captured at every ``P``); a strikethrough's top ``round(P * yStrikeoutPosition /
    em)`` above the baseline, ``max(1, round(P * yStrikeoutSize / em))`` thick; Calibri from
    ``P`` 171 draws these three a pixel higher. A double underline's lines are two thicknesses
    apart; Aptos Narrow's are the single's, stepped up as on the Mac, and Calibri's and Arial's
    are a captured table at every ``P`` from 1 to 384. Up to ``P`` 11 every value of the three
    faces is the captured line. A line may reach past the row (see :class:`RuleMetrics`).

    Raises:
        LookupError: A face without its rule, a size outside 1 to 72, or Courier New at a
            screen size outside 6 to 144, where its gap is not measured; on Windows, a face
            other than the three, a zoom that is not whole or outside 10 to 400, or a size and
            zoom under one pixel an em.
    """
    _platform = check_platform(platform, "rule_metrics", windows=True)
    key, size = _row_face(font)
    if _platform.name == "windows":
        return _windows_rules(key, size, zoom)
    if key not in _RULE_TABLES:
        msg = (
            f"no rule positions for {_DISPLAY[key]}; measured: Calibri, Aptos Narrow, Arial, "
            "Times New Roman and Georgia"
        )
        raise LookupError(msg)
    position, thickness, strike_position, strike_size, descent = _RULE_TABLES[key]
    exact = size * zoom / 100
    p = math.floor(exact + 0.5)  # Excel draws the lines at the nearest whole screen size

    def floored(units: int) -> int:
        return math.floor(p * units / 2048)

    def rounded(units: int) -> int:
        return math.floor(p * units / 2048 + 0.5)

    # the rule, in screen points
    top, line, accounting = floored(-position), max(1, floored(thickness)), rounded(descent)
    strike = (rounded(strike_position), max(1, rounded(strike_size)))
    if key == "courier new":
        rule = _courier_rule(p)
    elif key == "times new roman" and _TIMES_WIDE_GAP[0] <= p <= _TIMES_WIDE_GAP[1]:
        # a gap wider than two thicknesses: the double hangs up from the single's position, and
        # the single moves onto its first line
        wide = 3 * line
        rule = (top - wide, line, top - wide, line, wide, accounting, *strike)
    else:
        double = max(1, floored(min(thickness, strike_size)))
        first = top
        while first + 3 * double > accounting:
            first -= double
        rule = (top, line, first, double, 2 * double, accounting, *strike)
    # the captured lines, in half screen points, where read
    drawn = DRAWN.get((key, float(p)))
    half = [2 * v for v in rule] if drawn is None else [
        2 * v if d is None else d for d, v in zip(drawn, rule, strict=True)
    ]  # fmt: skip
    k = 2 * zoom / 100  # half screen points per sheet point
    u_top, u_thick, d_top, d_thick, gap, acc, s_top, s_thick = (v / k for v in half)
    return RuleMetrics(
        underline_top=u_top,
        underline_thickness=u_thick,
        double_gap=gap,
        accounting_top=acc,
        strikethrough_top=s_top,
        strikethrough_thickness=s_thick,
        double_top=d_top,
        double_thickness=d_thick,
    )


# --- text insets (CAPTURING.md section 2 item 81) -----------------------------------------------

# family (lower case) -> the advance of "0" in units of 2048 an em (capture/font_tables.py)
_DIGIT_ADVANCE = {"calibri": 1038, "aptos narrow": 1038, "arial": 1139}
# (family, ppem) -> the digit width Excel for Windows hints there, where it is not the rounded one
_WINDOWS_DIGIT = {("calibri", 72): 37, ("aptos narrow", 72): 37}


@dataclass(frozen=True, slots=True)
class TextInsets:
    """Where a cell's text starts and ends, in sheet points from the cell's boundaries.

    The left boundary's gridline counts on the left: one screen unit (a point on the Mac, two
    device pixels on its 2x display; a pixel on Windows). The right boundary's belongs to the
    next cell.

    Attributes:
        left: From the cell's left boundary to left-aligned text: the gridline and the clear
            space.
        right: From right-aligned text to the cell's right boundary: the clear space, and on
            the Mac for an italic face also its overhang at the screen size (the text ends
            that much earlier; an underline under it too).
        accounting_left: ``left`` for a cell with an accounting underline, whose line spans the
            cell between the two; the same as ``left`` on both platforms.
        accounting_right: ``right`` for such a cell: where its text ends. The accounting line
            itself ends at the clear space in an italic face too.
    """

    left: float
    right: float
    accounting_left: float
    accounting_right: float


def text_insets(
    font: FontSpec | None = None,
    *,
    normal_font: FontSpec | None = None,
    zoom: float = 100,
    platform: str | Platform | None = None,
) -> TextInsets:
    """Where Excel starts and ends the text of a cell in ``font`` (CAPTURING.md items 81-84).

    The text starts after the left gridline (one screen unit) and ``ceil(D / 4)`` screen units of
    clear space, and ends that clear space before the right boundary, ``D`` the width of the
    widest digit of the cell's font at its screen size: on the Mac, screen points at
    ``S = round(size * zoom / 100)``, ``D = round(S * u / E)`` from the face's recovered
    advances; on Windows (96 dpi), device pixels at the ppem ``round(size * 4/3 * zoom / 100)``,
    ``D = round(ppem * digit advance / em)`` as the font hints it (a half rounds up). On the Mac
    right-aligned italic text ends its overhang at the screen size further in. The same rule
    makes the Normal font's ``left + right`` the column padding, ``2 * ceil(mdw / 4) + 1``. It
    depends on the cell's font alone: the workbook's Normal font (``normal_font``, accepted for
    symmetry with :func:`vertical_metrics`) does not move it, and the zoom changes it only
    through the screen size, so in sheet points it varies with the zoom. A cell with an
    accounting underline places its text as a plain one.

    Coverage: on the Mac every face :func:`measured_fonts` lists, from 1 to 72 points;
    captured at zoom 100 and 150 for Calibri, Arial and Aptos Narrow from 9 to 72 points under
    two Normal fonts, and for Times New Roman, Verdana, Georgia and Courier New and the italics
    of Verdana, Calibri, Aptos Narrow and Arial at 11, 36 and 72 points, every row by this rule
    but two (an italic overhang at a screen size where it is not measured); carried over to the
    other faces and italics. On Windows Calibri, Aptos Narrow and Arial, every captured row by
    this rule, no italic among them.

    Raises:
        LookupError: A face not covered, a size outside 1 to 72, or a size and zoom under one
            screen unit an em.
    """
    p = check_platform(platform, "text_insets", windows=True)
    windows = p.name == "windows"
    if windows:
        key, size = _row_face(font)
        if normal_font is not None:
            _row_face(normal_font)
        advance = _DIGIT_ADVANCE.get(key)
        if advance is None:
            msg = f"no text insets for {_DISPLAY[key]}; measured: Calibri, Aptos Narrow and Arial"
            raise LookupError(msg)
        name, units, em = _DISPLAY[key], {"0": float(advance)}, 2048
    else:  # every face Excel for Mac measures, by its widest digit
        f = resolve_font(font, p)
        if normal_font is not None:
            resolve_font(normal_font, p)
        name, size = f.name, f.size
        if not 1.0 <= size <= _LARGEST:
            msg = f"no text insets for {name} {size:g}; covered sizes: 1 to 72"
            raise LookupError(msg)
        em, units = _mac_face(f.name, f.style)
    per_point = 4 / 3 if windows else 1.0  # screen units a point at zoom 100
    screen = math.floor(size * per_point * zoom / 100 + 0.5)
    if screen < 1:
        msg = f"{name} {size:g} at zoom {zoom:g} is under one screen unit an em"
        raise LookupError(msg)
    digit = max(math.floor(screen * units[d] / em + 0.5) for d in _DIGITS if d in units)
    if windows:
        digit = _WINDOWS_DIGIT.get((key, screen), digit)
    clear = math.ceil(digit / 4)
    step = 75 if windows else 100  # sheet points a screen unit, times the zoom
    left, right = (1 + clear) * step / zoom, clear * step / zoom
    if not windows and f.italic:
        # right-aligned italic text ends its overhang at the screen size further in (the
        # accounting line still runs to the clear space; insets_faces.layout.tsv)
        right += _overhang(f.name, float(screen), f.style) * step / zoom
    return TextInsets(left=left, right=right, accounting_left=left, accounting_right=right)
