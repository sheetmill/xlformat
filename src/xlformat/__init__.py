"""Apply Excel number-format strings to values and get the text Excel displays.

Pure Python, no runtime dependencies. ``format_value`` returns the plain string;
``format_cell`` returns a ``FormattedCell`` with color, alignment, fill and overflow
information. ``NumberFormat.parse`` gives the compiled format. See ``ARCHITECTURE.md``.
"""

from __future__ import annotations

from xlformat import _render
from xlformat._assemble import FormattedCell
from xlformat._builtin import BUILTIN_FORMATS, builtin_format
from xlformat._error_values import ExcelError
from xlformat._errors import FormatError
from xlformat._measure import (
    ColumnMetrics,
    Font,
    FontSpec,
    InkMetrics,
    Measure,
    RowMetrics,
    RuleMetrics,
    TextInsets,
    VerticalMetrics,
    column_metrics,
    column_metrics_from_points,
    digit_width,
    file_column_metrics,
    font_substitutes,
    glyph_widths,
    ink_metrics,
    measured_fonts,
    overhang,
    row_metrics,
    rule_metrics,
    text_insets,
    text_width,
    vertical_metrics,
)
from xlformat._palette import DEFAULT_PALETTE, section_rgb
from xlformat._parser import NumberFormat, Section, SectionKind
from xlformat._platform import (
    DEFAULT_PLATFORM,
    Platform,
    RegionalSettings,
    check_platform,
    check_regional,
)
from xlformat._quirks import Quirks

__version__ = "0.9.0"

__all__ = [
    "BUILTIN_FORMATS",
    "DEFAULT_PALETTE",
    "DEFAULT_PLATFORM",
    "ColumnMetrics",
    "ExcelError",
    "Font",
    "FormatError",
    "FormattedCell",
    "InkMetrics",
    "NumberFormat",
    "Platform",
    "Quirks",
    "RegionalSettings",
    "RowMetrics",
    "RuleMetrics",
    "Section",
    "SectionKind",
    "TextInsets",
    "VerticalMetrics",
    "__version__",
    "builtin_format",
    "column_metrics",
    "column_metrics_from_points",
    "digit_width",
    "file_column_metrics",
    "font_substitutes",
    "format_cell",
    "format_value",
    "glyph_widths",
    "ink_metrics",
    "measured_fonts",
    "overhang",
    "row_metrics",
    "rule_metrics",
    "section_rgb",
    "text_function",
    "text_insets",
    "text_width",
    "vertical_metrics",
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
    show_zeros: bool = True,
    quirks: Quirks | None = None,
    platform: str | Platform | None = None,
    regional_settings: RegionalSettings | None = None,
) -> str:
    """Return the text Excel displays for ``value`` under ``format_string``.

    Args:
        value: ``int``, ``float``, ``Decimal``, ``bool``, ``str``, ``None``, ``date``,
            ``datetime``, ``time``, ``timedelta`` or :class:`ExcelError`. An error shows its
            own text under any format, centered, and overflows to ``#`` like a number.
        format_string: An Excel number format; ``""`` means ``General``.
        width: Character mode: the width in characters, or ``None`` for the unconstrained
            rendering (``*`` fills emit nothing, ``_x`` pads one space, overflow is 255 ``#``).
        column_width: Excel mode: the column width in Column Width dialog units; the text is
            measured with the glyph table of ``font`` as Excel on ``platform`` does, so overflow,
            ``#`` and fill counts and General digit reduction match the sheet. Not with ``width``.
        font: The cell's font for Excel mode, a name, ``(name, size)`` or :class:`Font`; default
            ``("Calibri", 11)``, openpyxl's workbook default. The text is measured in it. On the
            Mac the families of :func:`measured_fonts` are measured in four styles at any size
            (captured to 192 points, the same rule unverified above); on Windows Calibri,
            Aptos Narrow and Arial at 9 to 24 points, scaled from the nearest measured size
            elsewhere; other fonts raise ``LookupError``.
            A line feed and a tab measure as a space and a carriage return as nothing; glyphs
            beyond the cp1252 characters measure as the widest digit.
        normal_font: The workbook's Normal font for Excel mode, whose digit width sets the
            column (one unit of ``column_width``); default ``font``.
        measure: A custom text measure for ``width``, whose units then also apply to ``width``.
        date1904: Use the 1904 date system for dates and serials.
        show_zeros: False for a sheet that hides zeros (``showZeros="0"``, Excel's "Show a zero
            in cells that have zero value" turned off): a number equal to 0, ``-0.0``
            included, shows nothing under a format of fewer than three sections. A format
            with a zero section shows it as usual, and so does a value that only rounds to 0.
            Pass the value as stored.
        quirks: Which Excel quirks to reproduce; ``None`` means ``Quirks.EXCEL``.
        platform: Where the sheet is viewed: ``"mac"``, ``"windows"`` or a :class:`Platform`;
            ``None`` means the default, ``"mac"``. Excel for Windows is measured at 96 dpi;
            another resolution raises ``LookupError``.
        regional_settings: The viewer's :class:`RegionalSettings`; ``None`` means the
            platform's en-US preset.

    Raises:
        FormatError: The format string is one Excel refuses to open from a file.
        TypeError: ``value`` has an unsupported type.
        ValueError: ``value`` is NaN or infinite.
    """
    check_regional(regional_settings, check_platform(platform, "format_value", windows=True))
    return NumberFormat.parse(format_string).format(
        value,
        width=width,
        column_width=column_width,
        font=font,
        normal_font=normal_font,
        measure=measure,
        date1904=date1904,
        show_zeros=show_zeros,
        quirks=quirks,
        platform=platform,
        regional_settings=regional_settings,
    )


def text_function(
    value: object,
    format_string: str,
    *,
    date1904: bool = False,
    quirks: Quirks | None = None,
    platform: str | Platform | None = None,
    regional_settings: RegionalSettings | None = None,
) -> str | ExcelError:
    """Return what Excel's ``TEXT`` function returns for ``value`` under ``format_string``.

    The text a cell of no particular width shows, but for what ``TEXT`` does differently: a
    ``None`` value (a blank cell) is 0, an :class:`ExcelError` value is returned as it is, the
    empty format gives the empty text for a number (``-`` for a negative one) and leaves text and
    booleans as they are, and a format Excel refuses, or a value the format cannot show (a
    negative date), gives ``ExcelError.VALUE`` rather than raising or showing ``#``.

    Text that reads as a number stays text here: a caller evaluating ``TEXT("12", "0.00")``
    converts it first.

    Args:
        value: What :func:`format_value` accepts.
        format_string: An Excel number format.
        date1904: The workbook uses the 1904 date system.
        quirks: The quirk settings; ``None`` means ``Quirks.EXCEL``.
        platform: Where the sheet is viewed: ``"mac"``, ``"windows"`` or a :class:`Platform`;
            ``None`` means the default, ``"mac"``. ``TEXT`` gives the same text on both
            (CAPTURING.md section 2 item 72); it chooses the regional settings' default.
        regional_settings: The viewer's :class:`RegionalSettings`; ``None`` means the
            platform's en-US preset.

    Returns:
        The text as a ``str``, or an :class:`ExcelError` (itself a ``str``).
    """
    check_regional(regional_settings, check_platform(platform, "text_function", windows=True))
    if isinstance(value, ExcelError):
        return value
    if format_string == "":
        return _render.text_function_empty_format(value, date1904=date1904, quirks=quirks)
    try:
        fmt = NumberFormat.parse(format_string)
    except FormatError:
        return ExcelError.VALUE
    return fmt.text_function(
        value,
        date1904=date1904,
        quirks=quirks,
        platform=platform,
        regional_settings=regional_settings,
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
    show_zeros: bool = True,
    quirks: Quirks | None = None,
    platform: str | Platform | None = None,
    regional_settings: RegionalSettings | None = None,
) -> FormattedCell:
    """Like :func:`format_value` but return a :class:`FormattedCell` with display details."""
    check_regional(regional_settings, check_platform(platform, "format_cell", windows=True))
    return NumberFormat.parse(format_string).format_cell(
        value,
        width=width,
        column_width=column_width,
        font=font,
        normal_font=normal_font,
        measure=measure,
        date1904=date1904,
        show_zeros=show_zeros,
        quirks=quirks,
        platform=platform,
        regional_settings=regional_settings,
    )
