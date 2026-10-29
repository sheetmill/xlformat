"""Apply Excel number-format strings to values and get the text Excel displays.

Pure Python, no runtime dependencies. ``format_value`` returns the plain string;
``format_cell`` returns a ``FormattedCell`` with color, alignment, fill and overflow
information. ``NumberFormat.parse`` gives the compiled format. See ``ARCHITECTURE.md``.
"""

from __future__ import annotations

from xlformat._assemble import FormattedCell
from xlformat._errors import FormatError
from xlformat._measure import (
    ColumnMetrics,
    Font,
    FontSpec,
    Measure,
    column_metrics,
    file_column_metrics,
)
from xlformat._palette import DEFAULT_PALETTE
from xlformat._parser import NumberFormat, Section, SectionKind
from xlformat._quirks import Quirks

__version__ = "0.1.0"

__all__ = [
    "DEFAULT_PALETTE",
    "ColumnMetrics",
    "Font",
    "FormatError",
    "FormattedCell",
    "NumberFormat",
    "Quirks",
    "Section",
    "SectionKind",
    "__version__",
    "column_metrics",
    "file_column_metrics",
    "format_cell",
    "format_value",
]


def format_value(
    value: object,
    format_string: str = "General",
    *,
    width: int | None = None,
    column_width: float | None = None,
    font: FontSpec | None = None,
    normal_font: FontSpec | None = None,
    measure: Measure | None = None,
    date1904: bool = False,
    quirks: Quirks | None = None,
) -> str:
    """Return the text Excel displays for ``value`` under ``format_string``.

    Args:
        value: ``int``, ``float``, ``Decimal``, ``bool``, ``str``, ``None``, ``date``,
            ``datetime``, ``time`` or ``timedelta``.
        format_string: An Excel number format; ``""`` means ``General``.
        width: Character mode: the width in characters, or ``None`` for the unconstrained
            rendering (``*`` fills emit nothing, ``_x`` pads one space, overflow is 255 ``#``).
        column_width: Excel mode: the column width in Column Width dialog units; the text is
            measured with the glyph table of ``font`` as Excel for Mac does, so overflow, ``#``
            and fill counts and General digit reduction match the sheet. Not with ``width``.
        font: The cell's font for Excel mode, a name, ``(name, size)`` or :class:`Font`;
            default ``("Calibri", 11)``, openpyxl's workbook default. The text is measured in
            it. Calibri is measured at 9, 10, 11, 12, 14 and 16 points and scaled from the
            nearest measured size elsewhere; other fonts and styles without a table raise
            ``LookupError``. Glyphs outside the table measure as a digit.
        normal_font: The workbook's Normal font for Excel mode, whose digit width sets the
            column (one unit of ``column_width``); default ``font``.
        measure: A custom text measure for ``width``, whose units then also apply to ``width``.
        date1904: Use the 1904 date system for dates and serials.
        quirks: Which Excel quirks to reproduce; ``None`` means ``Quirks.EXCEL``.

    Raises:
        FormatError: The format string is one Excel refuses to open from a file.
        TypeError: ``value`` has an unsupported type.
        ValueError: ``value`` is NaN or infinite.
    """
    return NumberFormat.parse(format_string).format(
        value,
        width=width,
        column_width=column_width,
        font=font,
        normal_font=normal_font,
        measure=measure,
        date1904=date1904,
        quirks=quirks,
    )


def format_cell(
    value: object,
    format_string: str = "General",
    *,
    width: int | None = None,
    column_width: float | None = None,
    font: FontSpec | None = None,
    normal_font: FontSpec | None = None,
    measure: Measure | None = None,
    date1904: bool = False,
    quirks: Quirks | None = None,
) -> FormattedCell:
    """Like :func:`format_value` but return a :class:`FormattedCell` with display details."""
    return NumberFormat.parse(format_string).format_cell(
        value,
        width=width,
        column_width=column_width,
        font=font,
        normal_font=normal_font,
        measure=measure,
        date1904=date1904,
        quirks=quirks,
    )
