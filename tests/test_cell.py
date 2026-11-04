"""FormattedCell fields and the public API surface."""

import xlformat
from xlformat import (
    DEFAULT_PALETTE,
    ExcelError,
    FormattedCell,
    NumberFormat,
    Quirks,
    format_cell,
    format_value,
    text_function,
)


def test_public_names() -> None:
    assert set(xlformat.__all__) >= {
        "format_value",
        "format_cell",
        "NumberFormat",
        "FormattedCell",
        "Quirks",
        "FormatError",
    }


def test_format_cell_fields_for_a_colored_negative() -> None:
    cell = format_cell(-1234.5, "[Red]0.00;[Blue]-0.00")
    assert cell == FormattedCell("-1234.50", "-1234.50", "Blue", "right", None, False)


def test_text_function_is_the_width_less_text_with_excels_differences() -> None:
    assert text_function(1234.5, "*-0") == "1235"
    assert text_function(1234.5, "0_)") == "1235 "
    assert text_function(None, "0.00") == "0.00"
    assert text_function(12.5, "abc") == "a43c"
    assert text_function("abc", "0.00") == "abc"


def test_text_function_returns_errors_rather_than_raising_or_showing_hashes() -> None:
    assert text_function(-1, "m/d/yy") is ExcelError.VALUE
    assert text_function(1, "0.00e+00") is ExcelError.VALUE
    assert text_function(ExcelError.NA, "0.00e+00") is ExcelError.NA
    assert text_function(1, '"#"') == "#"
    assert type(text_function(1, '"#VALUE!"')) is str


def test_text_function_under_the_empty_format() -> None:
    assert text_function(5, "") == ""
    assert text_function(-5, "") == "-"
    assert text_function("abc", "") == "abc"
    assert text_function(True, "") == "TRUE"
    assert text_function(None, "") == ""
    assert NumberFormat.parse("").text_function(5) == "5"  # parse("") is General


def test_format_cell_fill_position_and_alignments() -> None:
    assert format_cell(5, "0*-").fill == (1, "-")
    assert format_cell("abc", "@").align == "left"
    assert format_cell(True, "0").align == "center"
    assert format_cell(None, "0").text == ""
    assert format_cell(1, "0").overflow is False


def test_full_text_is_always_set() -> None:
    cell = format_cell(1234.5, "#,##0.00")
    assert cell.full_text == cell.text == "1,234.50"
    assert format_value(1234.5, "#,##0.00") == "1,234.50"


def test_default_palette() -> None:
    assert len(DEFAULT_PALETTE) == 64
    assert DEFAULT_PALETTE["Red"] == DEFAULT_PALETTE["Color3"] == "FF0000"
    assert DEFAULT_PALETTE["Color25"] == DEFAULT_PALETTE["Color11"] == "000080"  # duplicates
    assert DEFAULT_PALETTE[format_cell(-1, "[Red]0;[Blue]-0").color or ""] == "0000FF"


def test_strict_fixed_denominator_is_the_closest_fraction() -> None:
    # Excel's nudge carries 1.515 over the half step (QUIRKS.md, fraction_denominators)
    assert format_value(1.515, "# ??/100") == "1 52/100"
    assert format_value(1.515, "# ??/100", quirks=Quirks.STRICT) == "1 51/100"
    assert format_value(0.91, "# ?/?") == "1    "
    assert format_value(0.91, "# ?/?", quirks=Quirks.STRICT) == " 8/9"
