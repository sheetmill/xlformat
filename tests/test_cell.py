"""FormattedCell fields and the public API surface."""

import xlformat
from xlformat import DEFAULT_PALETTE, FormattedCell, Quirks, format_cell, format_value


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
