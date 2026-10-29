"""Capture suites: which (value, format, width, date system) combinations Excel is asked about.

Each suite is a function returning ``Case`` objects. A case names the fact it exercises in
``tag`` so a failure reads as "rounding tie" rather than a row number. Suites are small cross
products chosen for one purpose each.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Literal

from capture.schema import read_rows

ROOT = Path(__file__).resolve().parents[1]

# Excel column widths (dialog units) used by width suites. 255 is Excel's maximum and stands
# in for "wide enough for anything" when string value is read instead of the clipboard.
GENERAL_WIDTHS: tuple[float, ...] = (1, 2, 3, 4, 5, 6, 7, 8, 8.43, 9, 10, 11, 12, 13, 14, 16, 20)

# Fast-tier cap per (format, width) group; see capture/tiers.py. Chosen so the planned suites
# (about 27,000 rows) give roughly 5,000 fast rows.
DEFAULT_FAST_CAP = 15
# Suites that exist only as full-tier regression data.
FULL_ONLY: dict[str, int] = {"capture2023": 0, "capture2023_sample": 0}


@dataclass(frozen=True, slots=True)
class CellFont:
    """A cell's own font, set on the cell instead of inheriting the workbook's Normal style."""

    name: str
    size: float
    bold: bool = False
    italic: bool = False

    @property
    def label(self) -> str:
        """``Calibri 14``, ``Calibri 11 bold italic``: used in tags and measure files."""
        style = " ".join(s for s, on in (("bold", self.bold), ("italic", self.italic)) if on)
        return f"{self.name} {self.size:g}" + (f" {style}" if style else "")


@dataclass(frozen=True, slots=True)
class Case:
    """One case before capture."""

    suite: str
    tag: str
    value: object
    format: str
    date1904: bool = False
    formula: str | None = None
    """When set (starting with ``=``), the cell holds this formula instead of ``value``; Excel
    evaluates it in IEEE arithmetic on open. ``value`` is then the double the formula denotes,
    and ``tag`` carries ``formula:`` plus the text, so the TSV schema is unchanged."""
    written: str | None = None
    """When set, the exact decimal text written into the file for a numeric cell. openpyxl's
    default is ``"%.16g"``, which denotes a different double for about half of all floats;
    ``repr(value)`` is the shortest text that denotes ``value`` exactly."""
    font: CellFont | None = None
    """When set, the cell's own font; otherwise the cell inherits the Normal style."""

    @property
    def value_type(self) -> str:
        """Type name as stored in the TSV."""
        return type(self.value).__name__

    @property
    def value_repr(self) -> str:
        """``repr`` of the value as stored in the TSV."""
        return repr(self.value)


@dataclass(frozen=True, slots=True)
class Suite:
    """A named list of cases plus the widths they are captured at.

    ``widths`` holds ``None`` for the width-less clipboard capture and numbers for
    ``string value`` reads at that Excel column width. ``measure`` asks the capture to autofit
    each distinct (value, format, cell font) in its own column and record Excel's own width
    for it.

    With ``width_source="file"`` the widths are written into the file as ``<col width>`` by
    openpyxl instead of being set in dialog units by the script, ``None`` meaning a column
    without a ``<col>`` entry, and the script only reads back what Excel made of them.
    ``sheet_format`` sets attributes of ``<sheetFormatPr>`` (``baseColWidth``,
    ``defaultColWidth``). A ``probe`` suite writes under ``capture/build/probes/`` instead of
    ``tests/captures/``: its rows are evidence for a model not yet built, not test data.
    """

    name: str
    cases: tuple[Case, ...]
    widths: tuple[float | None, ...] = (None,)
    measure: bool = False
    colors: bool = False  # read each width-less cell's display-format font color (slow loop)
    font_size: int | None = None  # Normal style font size for the workbook (openpyxl default 11)
    font_name: str = "Calibri"  # Normal style font family
    width_source: Literal["dialog", "file"] = "dialog"
    sheet_format: tuple[tuple[str, float], ...] = ()  # <sheetFormatPr> attributes to set
    probe: bool = False
    date1904: bool = False
    fast_cap: int = DEFAULT_FAST_CAP  # rows per (format, width) in the fast tier; 0 = full only


def capture2023_sample(step: int = 45) -> Suite:
    """Every ``step``-th row of the 2023 capture, for validating the pipeline against it."""
    _, rows = read_rows(ROOT / "tests" / "captures" / "capture2023.tsv.gz")
    cases = tuple(
        Case("capture2023_sample", "legacy", row.value, row.format_string) for row in rows[::step]
    )
    return Suite("capture2023_sample", cases, fast_cap=0)


def capture2026() -> Suite:
    """All 89,680 cases of the 2023 capture, rebuilt with exact number texts.

    Kept permanently next to ``capture2023``: the diff between the two measures Excel version
    drift and the effect of the 16-digit texts the 2023 workbook was written with.
    """
    _, rows = read_rows(ROOT / "tests" / "captures" / "capture2023.tsv.gz")
    cases = tuple(Case("capture2026", "legacy", row.value, row.format_string) for row in rows)
    return Suite("capture2026", cases, fast_cap=0)


_GENERAL_VALUES: tuple[tuple[str, object], ...] = (
    ("int-1digit", 1),
    ("int-2digits", 12),
    ("int-3digits", 123),
    ("int-5digits", 12345),
    ("int-8digits", 12345678),
    ("int-10digits", 1234567890),
    ("int-11digits", 12345678901),
    ("int-12digits", 123456789012),
    ("int-15digits", 123456789012345),
    ("int-16digits", 1234567890123456),
    ("int-1e15", 1000000000000000),
    ("int-1e11", 100000000000),
    ("int-neg-8digits", -12345678),
    ("int-neg-12digits", -123456789012),
    ("int-zero", 0),
    ("dec-short", 1.5),
    ("dec-2places", 12.34),
    ("dec-3places", 123.456),
    ("dec-4places", 1234.5678),
    ("dec-many", 1234567.891),
    ("dec-11chars", 123456.7891),
    ("dec-long-mantissa", 12345.6789012345),
    ("dec-pi", 3.14159265358979),
    ("dec-third", 1 / 3),
    ("dec-two-thirds", 2 / 3),
    ("dec-neg", -1234567.891),
    ("dec-neg-short", -1.5),
    ("tie-2.5", 2.5),
    ("tie-9.5", 9.5),
    ("tie-99.5", 99.5),
    ("tie-0.5", 0.5),
    ("tie-0.25", 0.25),
    ("tie-0.125", 0.125),
    ("frac-0.1", 0.1),
    ("frac-9places", 0.123456789),
    ("frac-10places", 0.1234567891),
    ("frac-0.05", 0.05),
    ("frac-neg-0.05", -0.05),
    ("frac-0.000123", 0.000123),
    ("frac-0.0001234567", 0.0001234567),
    ("frac-1e-5", 0.00001),
    ("frac-5e-5", 0.00005),
    ("frac-1.5e-7", 1.5e-7),
    ("frac-0.9999999999", 0.9999999999),
    ("dec-999.9999999", 999.9999999),
    ("general-sci-threshold", 99999999999.5),
    ("general-sci-just-below", 99999999999.4),
    ("general-sci-1.23457e11", 123457000000.0),
    ("huge-1e21", 1e21),
    ("huge-1e100", 1e100),
    ("huge-1.5e300", 1.5e300),
    ("tiny-1e-100", 1e-100),
    ("tiny-1e-300", 1e-300),
    ("neg-tiny", -1.5e-7),
    ("neg-huge", -1e100),
    ("text-long", "hello world text"),
    ("text-short", "abc"),
    ("bool-true", True),
)


def width_general() -> Suite:
    """General format at many column widths: digit reduction, scientific fallback, overflow."""
    cases = tuple(Case("width_general", tag, value, "General") for tag, value in _GENERAL_VALUES)
    return Suite("width_general", cases, widths=(None, *GENERAL_WIDTHS), measure=True)


def width_general14() -> Suite:
    """``width_general`` in a workbook whose Normal style is Calibri 14: Excel mode at a size
    that is neither openpyxl's 11 nor Excel for Mac's 12 (CAPTURING.md section 2 item 26)."""
    cases = tuple(Case("width_general14", tag, value, "General") for tag, value in _GENERAL_VALUES)
    return Suite(
        "width_general14", cases, widths=(None, *GENERAL_WIDTHS), measure=True, font_size=14
    )


_NUMERIC_VALUES: tuple[tuple[str, object], ...] = (
    ("zero", 0),
    ("one", 1),
    ("neg-one", -1),
    ("five", 5),
    ("nine", 9),
    ("ten", 10),
    ("ninety-nine", 99),
    ("hundred", 100),
    ("123", 123),
    ("neg-123", -123),
    ("999", 999),
    ("thousand", 1000),
    ("1234", 1234),
    ("neg-1234", -1234),
    ("12345", 12345),
    ("99999", 99999),
    ("100000", 100000),
    ("123456", 123456),
    ("neg-123456", -123456),
    ("1234567", 1234567),
    ("12345678", 12345678),
    ("123456789", 123456789),
    ("1234567890", 1234567890),
    ("12345678901", 12345678901),
    ("123456789012", 123456789012),
    ("16-digits", 1234567890123456),
    ("1e15", 10**15),
    ("1e16", 10**16),
    ("half", 0.5),
    ("neg-half", -0.5),
    ("1.5", 1.5),
    ("tie-2.5", 2.5),
    ("tie-0.125", 0.125),
    ("tie-0.375", 0.375),
    ("tie-1.0625", 1.0625),
    ("tie-0.25", 0.25),
    ("tie-0.75", 0.75),
    ("1.005", 1.005),
    ("1.015", 1.015),
    ("2.675", 2.675),
    ("9.995", 9.995),
    ("0.045", 0.045),
    ("0.055", 0.055),
    ("1234.5", 1234.5),
    ("1234.56", 1234.56),
    ("1234.567", 1234.567),
    ("1234.5678", 1234.5678),
    ("neg-1234.5678", -1234.5678),
    ("0.1", 0.1),
    ("0.01", 0.01),
    ("0.001", 0.001),
    ("neg-0.001", -0.001),
    ("neg-0.0001", -0.0001),
    ("neg-0.005", -0.005),
    ("0.005", 0.005),
    ("0.0049", 0.0049),
    ("0.0051", 0.0051),
    ("123.456", 123.456),
    ("9.915", 9.915),
    ("9.925", 9.925),
    ("0.123456789", 0.123456789),
    ("1234567.891", 1234567.891),
    ("99999.995", 99999.995),
    ("999.995", 999.995),
    ("0.9999", 0.9999),
    ("0.99999", 0.99999),
    ("1e-7", 1e-7),
    ("1e-10", 1e-10),
    ("123456789.123", 123456789.123),
    ("neg-999999.5", -999999.5),
    ("33.333333", 33.333333),
    ("66.666666", 66.666666),
    ("third", 1 / 3),
    ("two-thirds", 2 / 3),
    ("1e21", 1e21),
    ("1e100", 1e100),
    ("neg-1e-7", -1e-7),
    ("0.5e-3", 0.0005),
    ("12.5", 12.5),
    ("125", 125),
    ("1250", 1250),
    ("1500", 1500),
    ("1499.5", 1499.5),
    ("999999", 999999),
    ("1234567.5", 1234567.5),
    ("0.015", 0.015),
    ("0.025", 0.025),
    ("0.0", 0.0),
    ("text-abc", "abc"),
    ("text-empty", ""),
    ("text-digits", "123"),
    ("text-spaces", " x "),
    ("text-words", "Hello World"),
    ("bool-true", True),
    ("bool-false", False),
    ("none", None),
)

_NUMERIC_FORMATS: tuple[str, ...] = (
    "General",
    "0",
    "0.0",
    "0.00",
    "0.000",
    "#",
    "#.#",
    "#.##",
    "0.#",
    "0.0#",
    "0.##",
    "#,##0",
    "#,##0.0",
    "#,##0.00",
    "#,###",
    "#,###.##",
    "?.??",
    "???.???",
    "0,000",
    "#,##0,",
    "#,##0,,",
    "#,##0.0,",
    "0%",
    "0.0%",
    "0.00%",
    "0%%",
    "#,##0%",
    "00000",
    "000-00-0000",
    "###-####",
    "00000-0000",
    "(000) 000-0000",
    "$#,##0",
    "$#,##0.00",
    "$#,##0.00_);($#,##0.00)",
    "$#,##0_);[Red]($#,##0)",
    '"$"#,##0.00',
    '#,##0.00 "USD"',
    '"text"0',
    '0"text"',
    '0 "items"',
    '"Total: "0.00',
    'General"x"',
    "0 x",
    "[$€-2] #,##0.00",
    "[$€-x-euro2] #,##0.00_);[Red]([$€-x-euro2] #,##0.00)",
    "#,##0.00 [$€-x-euro1]",
    "0;-0",
    "0;(0)",
    "0;-0;0",
    "0;-0;;@",
    "0;;",
    "0;;;",
    '0;-0;"zero"',
    "0.00;[Red]-0.00",
    "[Red]0.00;[Blue]-0.00;[Green]0.00;[Magenta]@",
    ";;;",
    '"Yes";"Yes";"No"',
    '"pos";"neg";"zero";"text"',
    ";",
    "[>100]0.00;0",
    '[>=100]"big";"small"',
    '[<0]"neg";[>0]"pos";"zero"',
    '[=0]"zero";0.0',
    '[<>0]0.0;"nil"',
    '[>1000]#,##0,"K";#,##0',
    "[<=9999999]###-####;(###) ###-####",
    "[>100]0.00;0;@",
    "[Red][>5]0;[Blue]0",
    "0;[Red]0;@",
    "@",
    '@" units"',
    '"Name: "@',
    "@@",
    "(@)",
    "_(@_)",
    "*-@",
    "@*-",
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    "_(0_)",
    "* 0",
    "0* ",
    "*x0.00",
    "0.00*x",
    "_-0_-",
    ".",
    ".0",
    "0.",
    ",0",
    "0,",
    "0.0,",
    "0,0",
    "0 0",
    "0-0",
    "0.0.0",
    '0 "and" 0',
    "\\0",
    "0.00000000000000000000",
    "0.000000000000000000000000000000",
    "[>100]0.00",
    "[>1]",
    "[<0]0.0",
    "[Red]General",
    "[Red]",
    "[Blue]0.00",
    "[Red]( General )",
    "[<5]0",
    "[<-5]0",
    "[<=-1]0",
    "[<=0]0",
    "[>-5]0",
    '[<0]"neg"',
    "[<0]0;0",
)


def numeric_core() -> Suite:
    """Number, General and text sections: placeholders, grouping, scaling, percent, literals,
    sections, conditions, sign handling, rounding ties."""
    cases = tuple(
        Case("numeric_core", tag, value, fmt)
        for fmt in _NUMERIC_FORMATS
        for tag, value in _NUMERIC_VALUES
    )
    return Suite("numeric_core", cases)


def _t(h: int, m: int, s: float) -> float:
    """Fraction of a day for a time of day."""
    return (h * 3600 + m * 60 + s) / 86400


_DATETIME_VALUES: tuple[tuple[str, object], ...] = (
    ("serial-0", 0),
    ("serial-1", 1),
    ("serial-2", 2),
    ("serial-59", 59),
    ("serial-60", 60),
    ("serial-61", 61),
    ("serial-62", 62),
    ("serial-100", 100),
    ("serial-365", 365),
    ("serial-366", 366),
    ("serial-1000", 1000),
    ("y2000", 36526),
    ("y2000-12-31", 36891),
    ("y2023", 44927),
    ("y2023-noon", 44927.5),
    ("y2023-12-31", 45291),
    ("max", 2958465),
    ("max-noon", 2958465.5),
    ("max-23:59:59.5", 2958465 + _t(23, 59, 59.5)),
    ("max-23:59:59.999", 2958465 + _t(23, 59, 59.999)),
    ("max+1", 2958466),
    ("max+35", 2958500),
    ("neg-1", -1),
    ("neg-half", -0.5),
    ("neg-tiny", -1e-9),
    ("noon", 0.5),
    ("just-before-noon", _t(11, 59, 59.5)),
    ("11:59:59.4", _t(11, 59, 59.4)),
    ("23:59:59.5", _t(23, 59, 59.5)),
    ("23:59:59.4", _t(23, 59, 59.4)),
    ("23:59:59.999", _t(23, 59, 59.999)),
    ("23:59:59.9995", _t(23, 59, 59.9995)),
    ("23:59:30", _t(23, 59, 30)),
    ("23:59:29.5", _t(23, 59, 29.5)),
    ("00:00:30", _t(0, 0, 30)),
    ("00:00:29.5", _t(0, 0, 29.5)),
    ("12:00:00.5", _t(12, 0, 0.5)),
    ("1h", _t(1, 0, 0)),
    ("1h30", _t(1, 30, 0)),
    ("1min", _t(0, 1, 0)),
    ("1s", _t(0, 0, 1)),
    ("59.5s", _t(0, 0, 59.5)),
    ("36h", 1.5),
    ("42h", 1.75),
    ("60h", 2.5),
    ("100d-12h", 100.5),
    ("0.1", 0.1),
    ("0.123456789", 0.123456789),
    ("45000.123456789", 45000.123456789),
    ("third", 1 / 3),
    ("y1900-02-28-noon", 59.5),
    ("y1900-02-29-noon", 60.5),
    ("y1900-03-01-noon", 61.5),
    ("text", "abc"),
    ("bool", True),
)

_DATETIME_FORMATS: tuple[str, ...] = (
    "m/d/yy",
    "m/d/yyyy",
    "d-mmm-yy",
    "dddd, mmmm d, yyyy",
    "yyyy-mm-dd",
    "yy",
    "yyyy",
    "e",
    "mmm",
    "mmmm",
    "mmmmm",
    "ddd",
    "dddd",
    "d",
    "dd",
    "h",
    "hh",
    "h:m",
    "m:s",
    "mm",
    "d-m-yy",
    "h m s",
    "h:mm",
    "h:mm:ss",
    "hh:mm:ss",
    "h:mm AM/PM",
    "h:mm:ss AM/PM",
    "h AM/PM",
    "h A/P",
    "h am/pm",
    "h a/p",
    "hh:mm am/pm",
    "ss",
    "ss.0",
    "ss.00",
    "ss.000",
    "s",
    "s.0",
    "mm:ss.000",
    "[h]:mm:ss",
    "[h]",
    "[hh]",
    "[m]",
    "[mm]:ss",
    "[s]",
    "[ss]",
    "[h]:mm",
    "[m]:ss.00",
    "yyyy-mm-dd hh:mm:ss.000",
    "m/d/yy h:mm",
    "[$-F800]",
    "[$-F400]",
    "[$-F800]dddd, mmmm dd, yyyy",
    "[$-x-sysdate]",
    "[$-x-systime]",
    "dddd dd mmmm yyyy",
    "mmmm yyyy",
    "yyyymmdd",
    "hhmmss",
    'h"h" mm"m"',
    'd"th" mmm',
    "yyyy-mm-dd;@",
    'm/d/yy;"neg"',
    "[Red]m/d/yy;[Blue]General",
    "m/d/yy General",
    "General h:mm",
)


def datetime_suite() -> Suite:
    """Dates and times in the 1900 system: leap quirk, limits, rounding, elapsed, AM/PM."""
    cases = tuple(
        Case("datetime", tag, value, fmt)
        for fmt in _DATETIME_FORMATS
        for tag, value in _DATETIME_VALUES
    )
    return Suite("datetime", cases)


_DATE1904_VALUES: tuple[tuple[str, object], ...] = (
    ("serial-0", 0),
    ("serial-1", 1),
    ("serial-59", 59),
    ("serial-60", 60),
    ("serial-61", 61),
    ("serial-365", 365),
    ("serial-366", 366),
    ("serial-1461", 1461),
    ("serial-1462", 1462),
    ("neg-1", -1),
    ("neg-2", -2),
    ("neg-365", -365),
    ("neg-1462", -1462),
    ("neg-1463", -1463),
    ("neg-half", -0.5),
    ("neg-1.5", -1.5),
    ("neg-quarter", -0.25),
    ("neg-tiny", -0.0001),
    ("y2023", 44927 - 1462),
    ("max", 2958465 - 1462),
    ("max+1", 2958466 - 1462),
    ("noon", 0.5),
    ("36h", 1.5),
    ("leap-period-1900", -1462 + 59),
)

_DATE1904_FORMATS: tuple[str, ...] = (
    "m/d/yyyy",
    "dddd, mmmm d, yyyy",
    "yyyy-mm-dd",
    "h:mm",
    "h:mm:ss",
    "[h]:mm:ss",
    "hh:mm",
    "mm:ss",
    "[m]",
    "General",
    "0.00",
    "yyyy-mm-dd hh:mm:ss",
    "h:mm AM/PM",
    "-h:mm",
)


def date1904() -> Suite:
    """The 1904 date system, in its own workbook with ``date 1904`` switched on."""
    cases = tuple(
        Case("date1904", tag, value, fmt, date1904=True)
        for fmt in _DATE1904_FORMATS
        for tag, value in _DATE1904_VALUES
    )
    return Suite("date1904", cases, date1904=True)


_COND_FORMATS: tuple[str, ...] = (
    '[=0]"zero";0',
    '[=0]"zero";0.0',
    "[<=0]0",
    "[<=0]0.0",
    "[<0]0",
    "[<1]0",
    "[>=-1]0",
    "[<=0]0;0",
    "0;[<=0]0",
    '[>0]0;"neg"',
    '[<>0]0;"zero"',
)
_COND_VALUES: tuple[tuple[str, object], ...] = tuple(
    (f"{'neg-' if v < 0 else ''}{abs(v)}", v)
    for v in (
        0,
        0.001,
        -0.001,
        0.05,
        -0.05,
        0.4,
        -0.4,
        0.5,
        -0.5,
        0.6,
        -0.6,
        1e-7,
        -1e-7,
        1,
        -1,
        0.95,
        -0.95,
    )
)


def conditions_rounding() -> Suite:
    """Does a condition see the value before or after rounding to the section's precision?"""
    cases = tuple(
        Case("conditions_rounding", tag, value, fmt)
        for fmt in _COND_FORMATS
        for tag, value in _COND_VALUES
    )
    return Suite("conditions_rounding", cases)


_CLOCK_TIMES: tuple[tuple[int, int, float], ...] = (
    (0, 0, 0.5),
    (0, 0, 29.5),
    (0, 0, 59.5),
    (0, 59, 59.5),
    (1, 0, 0.5),
    (6, 30, 30.5),
    (11, 59, 59.5),
    (12, 0, 0.5),
    (23, 59, 59.5),
    (23, 59, 29.5),
)


def _clock_values() -> list[tuple[str, object]]:
    """Times at .5-second boundaries built four ways, so the doubles differ in the last bits."""
    out: list[tuple[str, object]] = []
    for h, m, sec in _CLOCK_TIMES:
        label = f"{h:02d}:{m:02d}:{sec:04.1f}"
        out.append((f"{label}-total", (h * 3600 + m * 60 + sec) / 86400))
        out.append((f"{label}-sum", h / 24 + m / 1440 + sec / 86400))
        out.append((f"{label}-nested", ((h * 60 + m) * 60 + sec) / 86400))
        out.append((f"{label}-day44927", 44927 + (h * 3600 + m * 60 + sec) / 86400))
    return out


def _clock_formulas() -> list[str]:
    """Excel expressions (without ``=``) for ``_clock_values`` in the same order.

    Each literal has at most 15 significant digits and each expression performs the same IEEE
    operations in the same order as the Python construction, so it denotes the same double.
    """
    out: list[str] = []
    for h, m, sec in _CLOCK_TIMES:
        out.append(f"({h}*3600+{m}*60+{sec})/86400")
        out.append(f"{h}/24+{m}/1440+{sec}/86400")
        out.append(f"(({h}*60+{m})*60+{sec})/86400")
        out.append(f"44927+({h}*3600+{m}*60+{sec})/86400")
    return out


_LEAK_SERIALS: tuple[int, ...] = (
    0,
    1,
    2,
    10,
    50,
    100,
    200,
    365,
    366,
    500,
    1000,
    2000,
    3000,
    5000,
    7000,
    10000,
    15000,
    20000,
    25000,
    30000,
    36526,
    40000,
    44927,
    45000,
    50000,
    60000,
    80000,
    100000,
    200000,
    500000,
    1000000,
    2000000,
    2958465,
)


def clock_boundaries() -> Suite:
    """Half-second boundaries under second formats, and the General leak across serials."""
    cases = [
        Case("clock_boundaries", tag, value, fmt)
        for fmt in ("ss", "[mm]:ss", "h:mm:ss")
        for tag, value in _clock_values()
    ]
    for serial in _LEAK_SERIALS:
        cases.append(Case("clock_boundaries", f"leak-{serial}", serial, "m/d/yy General"))
        cases.append(Case("clock_boundaries", f"leak-{serial}.5", serial + 0.5, "m/d/yy General"))
    return Suite("clock_boundaries", tuple(cases))


_SCI_VALUES: tuple[tuple[str, object], ...] = (
    ("zero", 0),
    ("one", 1),
    ("neg-one", -1),
    ("ten", 10),
    ("hundred", 100),
    ("thousand", 1000),
    ("123", 123),
    ("1234.5", 1234.5),
    ("123456.5", 123456.5),
    ("neg-123456.5", -123456.5),
    ("0.5", 0.5),
    ("0.0123", 0.0123),
    ("0.000123", 0.000123),
    ("1e-5", 1e-5),
    ("1e-7", 1e-7),
    ("1e-10", 1e-10),
    ("1e-30", 1e-30),
    ("1.5e-30", 1.5e-30),
    ("1e-100", 1e-100),
    ("1e-300", 1e-300),
    ("1e15", 1e15),
    ("1e21", 1e21),
    ("1e100", 1e100),
    ("1.5e300", 1.5e300),
    ("neg-1e100", -1e100),
    ("9.95", 9.95),
    ("9.995", 9.995),
    ("99.95", 99.95),
    ("999.95", 999.95),
    ("9999.5", 9999.5),
    ("99999.5", 99999.5),
    ("999999.5", 999999.5),
    ("0.995", 0.995),
    ("0.0995", 0.0995),
    ("9.5", 9.5),
    ("95", 95),
    ("950", 950),
    ("0.95", 0.95),
    ("0.095", 0.095),
    ("1234567.891", 1234567.891),
    ("123451234512345.5", 123451234512345.5),
    ("0.1234512345123456", 0.1234512345123456),
    ("12", 12),
    ("120", 120),
    ("1200", 1200),
    ("12000", 12000),
    ("120000", 120000),
    ("1.2", 1.2),
    ("0.12", 0.12),
    ("0.012", 0.012),
    ("neg-0.0123", -0.0123),
    ("neg-0.5", -0.5),
    ("tie-2.5", 2.5),
    ("tie-0.125", 0.125),
    ("third", 1 / 3),
    ("two-thirds", 2 / 3),
    ("pi", 3.14159265358979),
    ("e", 2.718281828459045),
    ("text", "abc"),
    ("bool", True),
)
_SCI_FORMATS: tuple[str, ...] = (
    "0.00E+00",
    "0.00E-00",
    "0.E+00",
    "0E+0",
    "0E+00",
    "#.#E+#",
    "#.##E+##",
    "##0.0E+0",
    "##0.00E+00",
    "000.0E+0",
    "#,##0.0E+0",
    "0.00E+00;[Red]-0.00E+00",
    "0.000E+000",
    "\\E0.00E+00",
    '0.00E+00" units"',
    "0.0%E+0",
)


def scientific() -> Suite:
    """Scientific sections: mantissa digit rule, exponent padding and sign, zero, carries."""
    cases = tuple(
        Case("scientific", tag, value, fmt) for fmt in _SCI_FORMATS for tag, value in _SCI_VALUES
    )
    return Suite("scientific", cases)


_FRACTION_VALUES: tuple[tuple[str, object], ...] = tuple(
    [(f"n32-{k}", k / 32) for k in range(1, 32)]
    + [(f"n64-{k}", k / 64) for k in (1, 3, 5, 31, 33, 63)]
    + [
        ("zero", 0),
        ("one", 1),
        ("two", 2),
        ("neg-one", -1),
        ("half", 0.5),
        ("neg-half", -0.5),
        ("third", 1 / 3),
        ("two-thirds", 2 / 3),
        ("seventh", 1 / 7),
        ("tenth", 0.1),
        ("0.15", 0.15),
        ("1.15", 1.15),
        ("4.15", 4.15),
        ("0.123", 0.123),
        ("1234.51", 1234.51),
        ("0.499", 0.499),
        ("0.4999", 0.4999),
        ("0.501", 0.501),
        ("0.5001", 0.5001),
        ("0.9999", 0.9999),
        ("0.99999", 0.99999),
        ("1.0001", 1.0001),
        ("0.0001", 0.0001),
        ("0.001", 0.001),
        ("0.01", 0.01),
        ("0.05", 0.05),
        ("0.95", 0.95),
        ("0.995", 0.995),
        ("0.9995", 0.9995),
        ("2.5", 2.5),
        ("12.75", 12.75),
        ("123.456", 123.456),
        ("1234.5678", 1234.5678),
        ("neg-1234.5678", -1234.5678),
        ("pi", 3.14159265358979),
        ("e", 2.718281828459045),
        ("sqrt2", 1.4142135623730951),
        ("99999.5", 99999.5),
        ("999999.5", 999999.5),
        ("1234567.891", 1234567.891),
        ("1e9+0.5", 1e9 + 0.5),
        ("12345678901.25", 12345678901.25),
        ("1e14+0.5", 1e14 + 0.5),
        ("123451234512345.5", 123451234512345.5),
        ("1e15", 1e15),
        ("1e16", 1e16),
        ("0.03125", 0.03125),
        ("0.015625", 0.015625),
        ("0.0078125", 0.0078125),
        ("1/1000", 0.001),
        ("999/1000", 0.999),
        ("1/999", 1 / 999),
        ("1/99", 1 / 99),
        ("1/9", 1 / 9),
        ("5/9", 5 / 9),
        ("8/9", 8 / 9),
        ("0.11", 0.11),
        ("0.51", 0.51),
        ("0.91", 0.91),
        ("0.19", 0.19),
        ("0.59", 0.59),
        ("0.99", 0.99),
        ("1.5", 1.5),
        ("neg-1.5", -1.5),
        ("neg-0.03125", -0.03125),
        ("neg-0.0001", -0.0001),
        ("text", "abc"),
        ("bool", True),
    ]
)
_FRACTION_FORMATS: tuple[str, ...] = (
    "# ?/?",
    "# ??/??",
    "# ???/???",
    "?/?",
    "??/??",
    "0/0",
    "#/#",
    "# 0/0",
    "0 0/0",
    "# ?/2",
    "# ?/4",
    "# ?/8",
    "# ??/16",
    "# ?/10",
    "# ??/100",
    "# ???/1000",
    "#,##0 ?/?",
    '# ?/?" units"',
    "[Red]# ?/?;[Blue]-# ?/?",
    "# ?/?;(# ?/?)",
    "?/?%",
)


def fraction() -> Suite:
    """Fraction sections: free and fixed denominators, zero fraction display, whole parts."""
    cases = tuple(
        Case("fraction", tag, value, fmt)
        for fmt in _FRACTION_FORMATS
        for tag, value in _FRACTION_VALUES
    )
    return Suite("fraction", cases)


def _around(base: float, fractions: tuple[str, ...]) -> list[tuple[str, float]]:
    return [(f"{int(base)}+{f}", float(f"{int(base)}{f}")) for f in fractions]


_ROUNDING_FRACTIONS: tuple[str, ...] = (
    ".25",
    ".4999",
    ".5",
    ".5001",
    ".51",
    ".52",
    ".53",
    ".54",
    ".55",
    ".56",
    ".57",
    ".58",
    ".59",
    ".75",
)


def _rounding_values() -> list[tuple[str, float]]:
    fr = _ROUNDING_FRACTIONS
    return (
        _around(12345123451234, fr)  # 14 integer digits + decimals
        + _around(123451234512345, fr)  # 15 integer digits
        + _around(1234512345123455, fr)  # 16 integer digits
        + [
            ("99999999999999+.5", 99999999999999.5),
            ("999999999999999+.5", 999999999999999.5),
            ("9999999999999999+.5", 9999999999999999.5),
        ]
    )


def _rounding_formulas() -> list[str]:
    """Excel expressions for ``_rounding_values`` in the same order, literals of 15 digits.

    A 16-digit integer part is written as a 15-digit literal times ten plus a one-digit
    whole and the fraction; the sum is exact or rounds to the same double as the constant
    (checked by ``clock_formulas``).
    """
    fr = _ROUNDING_FRACTIONS
    return (
        [f"12345123451234+0{f}" for f in fr]
        + [f"123451234512345+0{f}" for f in fr]
        + [f"123451234512345*10+5{f}" for f in fr]
        + ["99999999999999+0.5", "999999999999999+0.5", "999999999999999*10+9.5"]
    )


def _python_value(expr: str) -> float:
    """Evaluate a formula expression with Python floats (only ``+ * /`` and literals occur)."""
    expr = expr.replace(" ", "")
    for ch in expr:
        if ch not in "0123456789.+*/()":
            msg = f"unexpected character {ch!r} in {expr}"
            raise ValueError(msg)
    return float(eval(expr))  # a vetted arithmetic expression: digits and + * / ( ) only


def rounding_15digit() -> Suite:
    """The discriminating grid for Excel's 15-significant-digit tie rule.

    Evidence for CAPTURING.md section 2 item 8.
    """
    cases = tuple(
        Case("rounding_15digit", tag, value, fmt)
        for fmt in ("0", "0.0", "0.00")
        for tag, value in _rounding_values()
    )
    return Suite("rounding_15digit", cases)


def clock_formulas() -> Suite:
    """Boundary values as constants and as formulas, to test Excel's xlsx number parsing.

    Hypothesis, since disproved (CAPTURING.md section 2 items 15 and 16): Excel's parser for
    numbers in the file is not correctly rounded, so a 16- or 17-digit decimal written by
    openpyxl can become a double one ulp from Python's, which decides the displayed digit at .5
    boundaries. A formula
    is evaluated by Excel in IEEE arithmetic, so its double is known exactly. Every value here
    appears as the constant and as a formula with the ``formula:`` tag.
    """
    cases: list[Case] = []

    def pair(tag: str, value: float, fmt: str, expr: str) -> None:
        if _python_value(expr) != value:
            msg = f"{expr} evaluates to {_python_value(expr)!r}, not {value!r}"
            raise ValueError(msg)
        cases.append(Case("clock_formulas", tag, value, fmt))
        cases.append(Case("clock_formulas", f"formula:={expr}", value, fmt, formula=f"={expr}"))

    for fmt in ("ss", "[mm]:ss", "h:mm:ss"):
        for (tag, value), expr in zip(_clock_values(), _clock_formulas(), strict=True):
            assert isinstance(value, float)
            pair(tag, value, fmt, expr)
    for serial in _LEAK_SERIALS:
        pair(f"leak-{serial}", float(serial), "m/d/yy General", str(serial))
        pair(f"leak-{serial}.5", serial + 0.5, "m/d/yy General", f"{serial}+0.5")
    for fmt in ("0", "0.0", "0.00"):
        for (tag, value), expr in zip(_rounding_values(), _rounding_formulas(), strict=True):
            pair(tag, value, fmt, expr)
    return Suite("clock_formulas", tuple(cases))


_PARSER_MISC: tuple[float, ...] = (
    0.1234512345123456,
    0.27118634259259256,
    1 / 3,
    2 / 3,
    0.1,
    0.3,
    1e-7,
    123456.7891,
    12345.6789012345,
    3.14159265358979,
    0.000123,
    44927.5,
    2958465.9999942128,
    0.9999999999,
    1.5e-7,
)


def _bits(x: float) -> tuple[int, int, int]:
    """(exp, high, low): ``x * 2**exp`` is an integer in [2**52, 2**53); its top and low 30 bits."""
    if x <= 0:
        msg = "positive values only"
        raise ValueError(msg)
    exp = 52 - math.floor(math.log2(x))
    m = int(x * 2.0**exp)
    assert 2**52 <= m < 2**53
    assert m == x * 2.0**exp
    return exp, m >> 30, m & (2**30 - 1)


def parser_check() -> Suite:
    """Read back the double Excel loaded for a decimal text, bit-exactly.

    Each value is written twice: as openpyxl's default ``%.16g`` text (case value: the double
    that text denotes) and as ``repr`` (17 digits, denotes the Python double). Two formula rows
    follow each value cell: ``INT(B*2^P/2^30)`` and ``B*2^P - INT(B*2^P/2^30)*2^30``, the high
    and low bits of the significand, exact in IEEE arithmetic and shown exactly by ``General``.
    A row fails only if Excel's parser gave a different double from a correctly rounded one.
    """
    values: list[tuple[str, float]] = []
    for (tag, v), _expr in zip(_clock_values(), _clock_formulas(), strict=True):
        assert isinstance(v, float)
        values.append((tag, v))
    values += _rounding_values()
    values += [(f"misc-{v!r}", v) for v in _PARSER_MISC]
    cases: list[Case] = []
    for tag, v in values:
        for form, text in (("g16", f"{v:.16g}"), ("repr", repr(v))):
            x = float(text)
            row = len(cases) + 1
            exp, high, low = _bits(x)
            cases.append(Case("parser_check", f"{form}:{tag}", x, "General", written=text))
            f_high = f"=INT(B{row}*2^{exp}/2^30)"
            f_low = f"=B{row}*2^{exp}-INT(B{row}*2^{exp}/2^30)*2^30"
            cases.append(Case("parser_check", f"formula:{f_high}", high, "General", formula=f_high))
            cases.append(Case("parser_check", f"formula:{f_low}", low, "General", formula=f_low))
    return Suite("parser_check", tuple(cases), fast_cap=0)


_LONE_ZERO_FORMATS: tuple[str, ...] = (
    "[<=0]0",
    "[<=0]0.0",
    "[<=0]0;[>1]0",
    "[<=0.5]0",
    "[<=0.1]0",
    "[<>0]0",
    '[=0]"zero";0',
    "[<0]0",
    "[<=0]0;0",
    "[<=-0.2]0",
    "[>=-1]0",
    "[<=0]0.00;[>1]0.00",
)
_LONE_ZERO_VALUES: tuple[tuple[str, object], ...] = (
    ("neg-0.3", -0.3),
    ("neg-0.04", -0.04),
    ("neg-0.001", -0.001),
    ("neg-0.5", -0.5),
    ("neg-0.6", -0.6),
    ("neg-1", -1),
    ("zero", 0),
    ("pos-0.3", 0.3),
)


def lone_zero() -> Suite:
    """The re-render model for negatives that round to zero (CAPTURING.md section 2 item 17).

    Hypothesis: Excel renders a negative as its magnitude plus a minus sign; when the magnitude
    renders as zero it renders the positive magnitude again from section selection onward.
    Predictions: ``[<=0]0;[>1]0`` with -0.3 shows ``0.3``, ``[<=0.5]0`` shows ``0``,
    ``[<=0.1]0`` shows ``0.3``, ``[<>0]0`` shows ``0``, ``[=0]"zero";0`` shows ``0``,
    ``[<=0]0.0`` with -0.04 shows ``0.04``.
    """
    cases = tuple(
        Case("lone_zero", tag, value, fmt)
        for fmt in _LONE_ZERO_FORMATS
        for tag, value in _LONE_ZERO_VALUES
    )
    return Suite("lone_zero", cases)


_SELECTION_FORMATS: tuple[str, ...] = (
    "[<=0]0;[>1]0;0.0",  # two conditions that fail, then an unconditioned third section
    "0;[>1]0",  # unconditioned first section, conditioned second
    "[>1]0;0;0.0",  # failing condition, "everything else" second, a third section
    "0;[>1]0;0.0",
    '[<=0]0;[>1]0;"else"',
)
_SELECTION_VALUES: tuple[tuple[str, object], ...] = (
    ("neg-0.3", -0.3),
    ("neg-0.5", -0.5),
    ("neg-2", -2),
    ("zero", 0),
    ("pos-0.3", 0.3),
    ("pos-0.5", 0.5),
    ("pos-1.5", 1.5),
    ("pos-2", 2),
)


def selection_fallback() -> Suite:
    """Section selection when conditions fail and more sections follow.

    Evidence for CAPTURING.md section 2 item 19.
    """
    cases = tuple(
        Case("selection_fallback", tag, value, fmt)
        for fmt in _SELECTION_FORMATS
        for tag, value in _SELECTION_VALUES
    )
    return Suite("selection_fallback", cases)


_COLOR_NAMES: tuple[str, ...] = (
    "Black",
    "White",
    "Red",
    "Green",
    "Blue",
    "Yellow",
    "Magenta",
    "Cyan",
)


def colors() -> Suite:
    """Section colors: the eight names, Color1..56, color in later sections, with conditions."""
    cases: list[Case] = []

    def add(tag: str, value: object, fmt: str) -> None:
        cases.append(Case("colors", tag, value, fmt))

    signs: tuple[tuple[str, object], ...] = (("pos", 1), ("neg", -1), ("zero", 0))
    for name in _COLOR_NAMES:
        for tag, v in signs:
            add(f"name-{name}-{tag}", v, f"[{name}]0")
    for i in range(1, 57):
        add(f"index-{i}", 1, f"[Color{i}]0")
    for i in (1, 8, 9, 10, 17, 25, 33, 41, 49, 56):
        add(f"index-{i}-neg", -1, f"[Color{i}]0")
    add("spelling-lower-space", 1, "[color 3]0")
    add("spelling-upper", 1, "[COLOR3]0")
    for fmt in (
        "0;[Red]-0",
        "0;[Red]-0;[Blue]0",
        "0;0;[Green]0",
        "[Blue]0;[Red]-0;[Green]0;[Magenta]@",
    ):
        for tag, v in (*signs, ("text", "abc")):
            add(f"sections-{tag}", v, fmt)
    for fmt in (
        "[Red][<0]0;[Blue]0",
        "[<0][Red]0;0",
        "[Blue][>100]0;[Green][>10]0;[Red]0",
        '[Green][=0]"zero";[Red]0',
    ):
        for v in (-5, 0, 5, 50, 500):
            add(f"cond-{v}", v, fmt)
    for v in ("abc", 1, True):
        add("text-section", v, "0;-0;0;[Magenta]@")
    for v in ("abc", 5):
        add("text-only", v, "[Cyan]@")
    for v in (1, "abc", True):
        add("general-colored", v, "[Red]General")
    for v in (1, -1, "abc"):
        add("color-alone", v, "[Red]")
    for v in (1, -1):
        add("empty-colored", v, "[Red];[Blue]")
    for fmt in ("[Red]0.00E+00", "[Red]# ?/?", "[Red]m/d/yy", "[Red][h]:mm", "[Blue]0%"):
        for v in (1234.5, -1234.5):
            add("kinds", v, fmt)
    for fmt in ("0", "General", "0.00;-0.00", "@"):
        for tag, v in (*signs, ("text", "abc")):
            add(f"control-{tag}", v, fmt)
    for v in (-0.3, -1, 0.3):
        add("rerender", v, "[Red][<=0]0;[Blue]0")  # does the re-rendered 0.3 take the blue section?
    return Suite("colors", tuple(cases), colors=True)


_TIE_FORMATS: dict[int, str] = {
    10: "# ?/10",
    100: "# ??/100",
    1000: "# ???/1000",
    16: "# ??/16",
    7: "# ?/7",
    3: "# ?/3",
    2: "# ?/2",
    4: "# ?/4",
    8: "# ?/8",
    32: "# ??/32",
}
_TIE_WHOLES: tuple[int, ...] = (0, 1, 12, 1000)
_TIE_TARGETS: tuple[float, ...] = (0.25, 0.5, 1.0, 1.5, 2.0, 4.0, 8.0)


def _ulp_distance(x: float, whole: int, k: int, d: int) -> float:
    """Signed distance of ``(x - whole) * d`` below ``k + 0.5``, in ulps of ``k + 0.5``."""
    frac = Decimal(x) - whole
    return float((Decimal(k) + Decimal("0.5") - frac * d) / Decimal(math.ulp(k + 0.5)))


def _near_half_steps(d: int, whole: int) -> list[tuple[str, float]]:
    """Doubles whose fraction part times ``d`` sits close to a half step, by ulp distance.

    For every numerator ``k`` the doubles around ``whole + (k + 0.5) / d`` are measured; for
    each target distance (both signs) the closest achieved one is kept, so the tag records
    where the value really sits. Whole parts above zero reach only coarse distances.
    """
    found: list[tuple[float, float, int]] = []
    ks = range(d) if d <= 100 else range(0, d, 37)
    for k in ks:
        x = whole + (k + 0.5) / d
        for _ in range(6):
            x = math.nextafter(x, -math.inf)
        for _ in range(13):
            if math.floor(x) == whole:
                found.append((_ulp_distance(x, whole, k, d), x, k))
            x = math.nextafter(x, math.inf)
    out: dict[float, tuple[str, float]] = {}
    for target in _TIE_TARGETS:
        for sign in (1, -1):
            t = sign * target
            dist, x, k = min(found, key=lambda f: abs(f[0] - t))
            if abs(dist - t) <= max(0.15, 0.5 * target) and x not in out:
                out[x] = (f"D{d}-W{whole}-k{k}-{dist:+.3f}ulp", x)
    return list(out.values())


def fraction_ties() -> Suite:
    """Where Excel rounds the fixed-denominator numerator (CAPTURING.md section 2 item 22).

    Values whose fraction part times the denominator sits a measured number of ulps below or
    above a half step, plus exact binary ties, written with exact texts.
    """
    cases: list[Case] = []
    for d in (10, 100, 1000, 16, 7, 3):
        for whole in _TIE_WHOLES:
            for tag, x in _near_half_steps(d, whole):
                cases.append(Case("fraction_ties", tag, x, _TIE_FORMATS[d]))
    for d in (2, 4, 8, 16, 32):
        for whole in _TIE_WHOLES:
            for k in sorted({0, 1, d - 1}):
                x = whole + (k + 0.5) / d
                assert Decimal(x) == Decimal(whole) + (Decimal(2 * k + 1) / (2 * d))
                cases.append(Case("fraction_ties", f"tie-D{d}-W{whole}-k{k}", x, _TIE_FORMATS[d]))
    return Suite("fraction_ties", tuple(cases))


_FILL_WIDTHS: tuple[float, ...] = (3, 6, 9, 12)
_FILL_FORMATS: tuple[str, ...] = (
    "0",
    "00",
    "000000",
    "0000000000",
    "0.0",
    "0.00",
    "0.000000",
    "0*0",
    "*00",
    "0.00*0",
    "*00.00",
    "0*00*0",
    "0*0;0*0",
    "0*0;-0*0;0*0;@*0",
    "00*0",
    "0000*0",
    "0.0*0",
    "000.000",
    "0_)",
    "0_)0",
    "_(0",
    "0*0_)",
    "yyyymmdd",
    "yyyy",
    "hhmmss",
    "[hh]",
    "General",
    "@",
    "@*0",
    "0;0;0;@*0",
)
_FILL_VALUES: tuple[tuple[str, object], ...] = (
    ("zero", 0),
    ("one-digit", 7),
    ("two-digits", 12),
    ("four-digits", 1234),
    ("six-digits", 123456),
    ("seven-and-half", 1234567.5),
    ("negative", -1234),
    ("half", 0.5),
    ("eleven-digits", 12345678901),
    ("text", "ab"),
    ("bool", True),
)


def width_fill() -> Suite:
    """``*`` fill placement, overflow of fixed formats, text never overflowing, at four widths.

    Output is mostly digits so that character and pixel widths agree (ARCHITECTURE.md section
    8); the few formats with letters or slashes test the overflow kind, not the fit threshold.
    """
    cases = tuple(
        Case("width_fill", tag, value, fmt) for fmt in _FILL_FORMATS for tag, value in _FILL_VALUES
    )
    return Suite("width_fill", cases, widths=(None, *_FILL_WIDTHS))


def numeric_core_width() -> Suite:
    """Every numeric_core case at three column widths, read with ``string value``."""
    base = numeric_core()
    cases = tuple(Case("numeric_core_width", c.tag, c.value, c.format) for c in base.cases)
    return Suite("numeric_core_width", cases, widths=(6, 10, 14))


_TEXT_OVERFLOW_FORMATS: tuple[str, ...] = ('@" units"', "_(@_)", '"["@"]"', "@@", "@", '"<"@">"')


def text_overflow() -> Suite:
    """Does text through a section with literals or pads (no fill) ever overflow to '#'?"""
    cases = tuple(
        Case("text_overflow", tag, value, fmt)
        for fmt in _TEXT_OVERFLOW_FORMATS
        for tag, value in (("short", "ab"), ("long", "Hello World"))
    )
    return Suite("text_overflow", cases, widths=(None, 3, 8, 15))


_MONTHS = (
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December",
)  # fmt: skip
_DAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
_MIXED = (
    "Hello World",
    "1,234.50",
    "-12345678",
    "1.23457E+11",
    "12/31/2023",
    "$1,234.00",
    "(1,234)",
    "TRUE",
    "hello world text",
    "Tuesday, January 4, 2000",
    "Sunday, December 3, 2000",
    "1234.50---",
)


def _measure_strings() -> list[str]:
    out: list[str] = []
    for code in range(32, 127):
        out.extend(chr(code) * n for n in (1, 2, 5, 10))
    out += [f"{a}{b}" for a in "0123456789" for b in "0123456789"]
    out += [*_MONTHS, *(m[:3] for m in _MONTHS), *_DAYS, *(d[:3] for d in _DAYS)]
    out += ["AM", "PM", "E+07", "####", *_MIXED]
    return list(dict.fromkeys(out))


def _measure(name: str, font_size: int | None) -> Suite:
    cases = tuple(Case(name, f"s{i}", text, "@") for i, text in enumerate(_measure_strings()))
    return Suite(name, cases, widths=(), measure=True, font_size=font_size)


_SPACE_FILL_FORMATS: tuple[str, ...] = (
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    "0* ",
    "* 0",
)
_SPACE_FILL_VALUES: tuple[tuple[str, object], ...] = (
    ("zero", 0),
    ("one", 1),
    ("neg-one", -1),
    ("twelve", 12),
    ("1234.5", 1234.5),
    ("neg-1234.5", -1234.5),
    ("half", 0.5),
    ("million", 1234567.891),
    ("text", "abc"),
    ("bool", True),
)


def _space_fill(name: str, font_size: int) -> Suite:
    """Fills of spaces at several widths in a workbook of the given Normal font size.

    Autofit cannot resolve the space's advance (CAPTURING.md section 2 item 24); the count of fill
    spaces around a known text in a known budget can, so each measured size gets one of these.
    """
    cases = tuple(
        Case(name, tag, value, fmt)
        for fmt in _SPACE_FILL_FORMATS
        for tag, value in _SPACE_FILL_VALUES
    )
    return Suite(name, cases, widths=(None, 6, 10, 14, 20), font_size=font_size)


def space_fill9() -> Suite:
    """Space fills at Calibri 9."""
    return _space_fill("space_fill9", 9)


def space_fill10() -> Suite:
    """Space fills at Calibri 10."""
    return _space_fill("space_fill10", 10)


def space_fill11() -> Suite:
    """Space fills at Calibri 11 (openpyxl's default workbook font)."""
    return _space_fill("space_fill11", 11)


def space_fill12() -> Suite:
    """Space fills at Calibri 12 (Excel for Mac's default)."""
    return _space_fill("space_fill12", 12)


def space_fill14() -> Suite:
    """Space fills at Calibri 14."""
    return _space_fill("space_fill14", 14)


def space_fill16() -> Suite:
    """Space fills at Calibri 16."""
    return _space_fill("space_fill16", 16)


def measure9() -> Suite:
    """Autofit widths of the measure texts in a Calibri 9 workbook."""
    return _measure("measure9", 9)


def measure10() -> Suite:
    """Autofit widths of the measure texts in a Calibri 10 workbook."""
    return _measure("measure10", 10)


def measure14() -> Suite:
    """Autofit widths of the measure texts in a Calibri 14 workbook."""
    return _measure("measure14", 14)


def measure16() -> Suite:
    """Autofit widths of the measure texts in a Calibri 16 workbook."""
    return _measure("measure16", 16)


def measure11() -> Suite:
    """Autofit widths (units and points) of about 540 texts in a Calibri 11 workbook."""
    return _measure("measure11", None)


def measure12() -> Suite:
    """The same texts in a workbook whose Normal style is Calibri 12 (Excel for Mac's default)."""
    return _measure("measure12", 12)


# Column model (0.2). Widths written into the file, columns without <col>, columns narrower than
# one unit, and the 255-character limit (CAPTURING.md section 2 item 28).

_FILE_WIDTHS: tuple[float | None, ...] = (
    None,
    *(k / 8 for k in range(4, 97)),  # 0.5 to 12 in eighths
    3.7109375, 8.43, 8.7109375, 9.140625, 10.7109375, 13, 15, 20, 30, 60, 100, 255,
    # MDW * width just past a half, for each MDW from 5 to 8: separates the ECMA-376 pixel
    # formula (rounds up from 1 - trunc(128 / MDW) / 256 * MDW, about .508 at 6) from half-down
    *(round((k + f) / mdw, 7) for mdw in (5, 6, 7, 8) for k in (20, 50)
      for f in (0.502, 0.504, 0.506, 0.51)),
)  # fmt: skip


def _column_cases(name: str) -> tuple[Case, ...]:
    """Digit runs, two fills and two General values: enough to read a column's text budget."""
    cases = [Case(name, f"digits{n}", int("8" * n), "0") for n in range(1, 16)]
    cases += [
        Case(name, "fill-dash", 1, "0*-"),
        Case(name, "fill-space", 1, "* 0"),
        Case(name, "general-7.3", 1234567.891, "General"),
        Case(name, "general-0.9", 0.123456789, "General"),
    ]
    return tuple(cases)


def _column_file(
    name: str,
    font_size: int | None,
    font_name: str = "Calibri",
    sheet_format: tuple[tuple[str, float], ...] = (),
    widths: tuple[float | None, ...] = _FILE_WIDTHS,
) -> Suite:
    return Suite(
        name,
        _column_cases(name),
        widths=widths,
        font_size=font_size,
        font_name=font_name,
        width_source="file",
        sheet_format=sheet_format,
    )


def column_file9() -> Suite:
    """``<col width>`` as openpyxl writes it, read back as dialog units and points, Calibri 9."""
    return _column_file("column_file9", 9)


def column_file10() -> Suite:
    """The same at Calibri 10."""
    return _column_file("column_file10", 10)


def column_file11() -> Suite:
    """The same in openpyxl's default workbook (Calibri 11)."""
    return _column_file("column_file11", None)


def column_file12() -> Suite:
    """The same at Calibri 12."""
    return _column_file("column_file12", 12)


def column_file14() -> Suite:
    """The same at Calibri 14."""
    return _column_file("column_file14", 14)


def column_file16() -> Suite:
    """The same at Calibri 16."""
    return _column_file("column_file16", 16)


_DEFAULT_WIDTHS: tuple[float | None, ...] = (None, 8, 8.43, 10)


def column_default_base10_11() -> Suite:
    """A column without ``<col>`` when ``baseColWidth`` is 10 (Calibri 11)."""
    return _column_file("column_default_base10_11", None, sheet_format=(("baseColWidth", 10),),
                        widths=_DEFAULT_WIDTHS)  # fmt: skip


def column_default_base10_14() -> Suite:
    """A column without ``<col>`` when ``baseColWidth`` is 10 (Calibri 14)."""
    return _column_file("column_default_base10_14", 14, sheet_format=(("baseColWidth", 10),),
                        widths=_DEFAULT_WIDTHS)  # fmt: skip


def column_default_dcw12_11() -> Suite:
    """A column without ``<col>`` when ``defaultColWidth`` is 12 (Calibri 11)."""
    return _column_file("column_default_dcw12_11", None, sheet_format=(("defaultColWidth", 12),),
                        widths=_DEFAULT_WIDTHS)  # fmt: skip


def column_default_dcw843_14() -> Suite:
    """A column without ``<col>`` when ``defaultColWidth`` is 8.43 (Calibri 14)."""
    return _column_file("column_default_dcw843_14", 14,
                        sheet_format=(("defaultColWidth", 8.43),),
                        widths=_DEFAULT_WIDTHS)  # fmt: skip


_NARROW_CASES: tuple[tuple[str, object, str], ...] = (
    ("digit", 8, "0"),
    ("digits2", 88, "0"),
    ("general-1", 1, "General"),
    ("general-12", 12, "General"),
    ("general-half", 0.5, "General"),
    ("general-neg", -1, "General"),
    ("general-big", 123456, "General"),
    ("fixed", 1.5, "0.0"),
    ("fill", 1, "0*-"),
    ("date", 36529, "m/d/yy"),
    ("bool", True, "General"),
    ("text", "abc", "General"),
    ("text-at", "abc", "@"),
    ("text-fill", "abc", "@*-"),
    ("literal", 1, '"x"'),
    ("empty-section", 1, ";;;"),
)


def _column_narrow(name: str, font_size: int | None, mdw: int) -> Suite:
    """Dialog widths below one unit and just above, set by the script: every point to MDW + 7."""
    widths = tuple(round(p / (mdw + 5), 2) for p in range(1, mdw + 6)) + tuple(
        round((p - 5) / mdw, 2) for p in range(mdw + 6, mdw + 13)
    )
    cases = tuple(Case(name, tag, value, fmt) for tag, value, fmt in _NARROW_CASES)
    return Suite(name, cases, widths=(None, *widths), font_size=font_size)


def column_narrow11() -> Suite:
    """Columns narrower than one unit, Calibri 11."""
    return _column_narrow("column_narrow11", None, 6)


def column_narrow14() -> Suite:
    """Columns narrower than one unit, Calibri 14."""
    return _column_narrow("column_narrow14", 14, 7)


_LIMIT_FORMATS: tuple[str, ...] = ("0*-", "* 0", "0.00*-", "@*-", "*-@", '"x"*.0', "0")
_LIMIT_VALUES: tuple[tuple[str, object], ...] = (
    ("one", 1),
    ("long", 123456789012345),
    ("huge", 1e300),
    ("text", "abc"),
)


def limits() -> Suite:
    """Fills and ``#`` runs near 255 characters, in wide columns (Calibri 11)."""
    widths: tuple[float | None, ...] = (None, 40, 60, 80, 82, 84, 86, 88, 90, 100, 120, 125,
                                        127, 128, 130, 150, 200, 254, 255)  # fmt: skip
    cases = tuple(
        Case("limits", tag, value, fmt) for fmt in _LIMIT_FORMATS for tag, value in _LIMIT_VALUES
    )
    return Suite("limits", cases, widths=widths)


# General at whole-point widths (0.2). The width_general suites use whole dialog widths (and
# 8.43, which is 8.5); columns read from files are mostly fractional (openpyxl's 8.43 is 7.67
# units at Calibri 11), so these read the General budget at every point.


def _general_points(name: str, font_size: int | None, mdw: int) -> Suite:
    widths = tuple(round((p - 5) / mdw, 2) for p in range(mdw + 5, 101))
    cases = tuple(Case(name, c.tag, c.value, c.format) for c in width_general().cases)
    return Suite(name, cases, widths=(None, *widths), font_size=font_size)


def general_points9() -> Suite:
    """The ``width_general`` values at every whole-point column up to 100 points, Calibri 9."""
    return _general_points("general_points9", 9, 5)


def general_points11() -> Suite:
    """The same at Calibri 11."""
    return _general_points("general_points11", None, 6)


def general_points12() -> Suite:
    """The same at Calibri 12."""
    return _general_points("general_points12", 12, 6)


def general_points14() -> Suite:
    """The same at Calibri 14."""
    return _general_points("general_points14", 14, 7)


def general_points16() -> Suite:
    """The same at Calibri 16."""
    return _general_points("general_points16", 16, 8)


_GENERAL_LITERAL_FORMATS: tuple[str, ...] = ('General"x"', "( General )", '"$"General',
                                             'General" units"', "General;-General")  # fmt: skip


def general_literals_points() -> Suite:
    """General among literals at every whole-point column up to 100 points (Calibri 11)."""
    name = "general_literals_points"
    widths = tuple(round((p - 5) / 6, 2) for p in range(11, 101))
    cases = tuple(
        Case(name, c.tag, c.value, fmt)
        for fmt in _GENERAL_LITERAL_FORMATS
        for c in width_general().cases
    )
    return Suite(name, cases, widths=(None, *widths))


# Cells whose font differs from the workbook's Normal style (0.2; CAPTURING.md section 2 item
# 29). The tag ends in the cell font's label in brackets, "[normal]" for the Normal style; the
# capture test reads the cell font from it.

_CELL_FONT_VALUES: tuple[tuple[str, object, str], ...] = (
    *((f"digits{n}", int("8" * n), "0") for n in range(3, 15)),
    ("fill-dash", 1, "0*-"),
    ("fill-space", 1, "* 0"),
    ("text-fill", "ab", "@*-"),
    ("general-7.3", 1234567.891, "General"),
    ("general-neg8", -12345678, "General"),
    ("general-small", 0.000123, "General"),
    ("general-12", 123456789012, "General"),
    ("comma", 1234567.891, "#,##0.00"),
    ("date-long", 36529, "dddd, mmmm d, yyyy"),
    ("letters", 1, '0" Wmwmwm"'),
)


def _cell_font_cases(
    name: str, fonts: tuple[CellFont | None, ...], values: tuple[tuple[str, object, str], ...]
) -> tuple[Case, ...]:
    return tuple(
        Case(name, f"{tag} [{font.label if font else 'normal'}]", value, fmt, font=font)
        for font in fonts
        for tag, value, fmt in values
    )


def _points(mdw: int, low: int = 11, high: int = 100) -> tuple[float, ...]:
    """Dialog widths at every whole point from ``low`` to ``high`` for a Normal font of MDW."""
    return tuple(round((p - 5) / mdw, 2) for p in range(max(low, mdw + 5), high + 1))


_REGULAR_CELL_FONTS: tuple[CellFont | None, ...] = (
    None,
    *(CellFont("Calibri", s) for s in (9, 10, 12, 14, 16)),
)


def cell_font() -> Suite:
    """Calibri 9 to 16 cells in a Calibri 11 workbook: fit, fills, General, dates, literals."""
    cases = _cell_font_cases("cell_font", _REGULAR_CELL_FONTS, _CELL_FONT_VALUES)
    return Suite("cell_font", cases, widths=(None, *_points(6)))


def _general_values() -> tuple[tuple[str, object, str], ...]:
    return tuple((c.tag, c.value, c.format) for c in width_general().cases)


def general_cell_font() -> Suite:
    """The ``width_general`` values in Calibri 9 to 16 cells of a Calibri 11 workbook."""
    fonts = tuple(CellFont("Calibri", s) for s in (9, 10, 12, 14, 16))
    cases = _cell_font_cases("general_cell_font", fonts, _general_values())
    return Suite("general_cell_font", cases, widths=(None, *_points(6)))


def cell_font_numeric() -> Suite:
    """Every ``numeric_core`` case in Calibri 9 and 14 cells of a Calibri 11 workbook."""
    values = tuple((c.tag, c.value, c.format) for c in numeric_core().cases)
    fonts = (CellFont("Calibri", 9), CellFont("Calibri", 14))
    cases = _cell_font_cases("cell_font_numeric", fonts, values)
    return Suite("cell_font_numeric", cases, widths=(6, 10, 14))


_ZERO_SIGN_FORMATS: tuple[str, ...] = ("General", "[Red]General", "@", "[>100]0.00",
                                       "General;General", "[Blue]General;[Red]General",
                                       '@" x"', "")  # fmt: skip
_ZERO_SIGN_VALUES: tuple[tuple[str, object], ...] = (
    ("neg-0.05", -0.05),
    ("neg-0.001", -0.001),
    ("neg-1.5e-07", -1.5e-07),
    ("neg-0.4", -0.4),
    ("neg-1.5", -1.5),
)


def general_zero_sign() -> Suite:
    """Negatives that General reduces to zero, under the plain format and General sections."""
    name = "general_zero_sign"
    cases = tuple(
        Case(name, tag, value, fmt)
        for fmt in _ZERO_SIGN_FORMATS
        for tag, value in _ZERO_SIGN_VALUES
    )
    return Suite(name, cases, widths=(None, *_points(6, 11, 50)))


def aptos_narrow_measure11() -> Suite:
    """Autofit widths of the measure texts in a workbook whose Normal style is Aptos Narrow 11."""
    s = _measure("aptos_narrow_measure11", 11)
    return Suite(s.name, s.cases, widths=(), measure=True, font_size=11,
                 font_name="Aptos Narrow", probe=True)  # fmt: skip


def aptos_narrow_space_fill11() -> Suite:
    """Space fills at Aptos Narrow 11."""
    s = _space_fill("aptos_narrow_space_fill11", 11)
    return Suite(s.name, s.cases, widths=s.widths, font_size=11, font_name="Aptos Narrow",
                 probe=True)  # fmt: skip


SUITES: dict[str, Callable[[], Suite]] = {
    "capture2023_sample": capture2023_sample,
    "capture2026": capture2026,
    "width_general": width_general,
    "rounding_15digit": rounding_15digit,
    "numeric_core": numeric_core,
    "datetime": datetime_suite,
    "date1904": date1904,
    "conditions_rounding": conditions_rounding,
    "clock_boundaries": clock_boundaries,
    "scientific": scientific,
    "fraction": fraction,
    "clock_formulas": clock_formulas,
    "parser_check": parser_check,
    "lone_zero": lone_zero,
    "selection_fallback": selection_fallback,
    "colors": colors,
    "fraction_ties": fraction_ties,
    "width_fill": width_fill,
    "numeric_core_width": numeric_core_width,
    "text_overflow": text_overflow,
    "measure11": measure11,
    "measure12": measure12,
    "measure9": measure9,
    "measure10": measure10,
    "measure14": measure14,
    "measure16": measure16,
    "space_fill9": space_fill9,
    "space_fill10": space_fill10,
    "space_fill11": space_fill11,
    "space_fill12": space_fill12,
    "space_fill14": space_fill14,
    "space_fill16": space_fill16,
    "width_general14": width_general14,
    "column_file9": column_file9,
    "column_file10": column_file10,
    "column_file11": column_file11,
    "column_file12": column_file12,
    "column_file14": column_file14,
    "column_file16": column_file16,
    "column_default_base10_11": column_default_base10_11,
    "column_default_base10_14": column_default_base10_14,
    "column_default_dcw12_11": column_default_dcw12_11,
    "column_default_dcw843_14": column_default_dcw843_14,
    "column_narrow11": column_narrow11,
    "column_narrow14": column_narrow14,
    "limits": limits,
    "general_points9": general_points9,
    "general_points11": general_points11,
    "general_points12": general_points12,
    "general_points14": general_points14,
    "general_points16": general_points16,
    "general_literals_points": general_literals_points,
    "cell_font": cell_font,
    "general_cell_font": general_cell_font,
    "cell_font_numeric": cell_font_numeric,
    "general_zero_sign": general_zero_sign,
    "aptos_narrow_measure11": aptos_narrow_measure11,
    "aptos_narrow_space_fill11": aptos_narrow_space_fill11,
}


def distinct_measure_cases(cases: Iterable[Case]) -> list[Case]:
    """The first case of each distinct (value, format, cell font), for the measure sheet."""
    seen: dict[tuple[str, str, str, CellFont | None], Case] = {}
    for case in cases:
        seen.setdefault((case.value_type, case.value_repr, case.format, case.font), case)
    return list(seen.values())
