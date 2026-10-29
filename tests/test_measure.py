"""Excel-mode measure: measured sizes are exact tables, other sizes scale the nearest one."""

import pytest

from xlformat import Font, format_value
from xlformat._glyphs_calibri import CALIBRI
from xlformat._measure import _advances, column_budget, glyph_measure, resolve_font

MEASURED_SIZES = (9, 10, 11, 12, 14, 16)


def test_measured_sizes_are_bundled() -> None:
    assert tuple(sorted(CALIBRI)) == MEASURED_SIZES
    # the maximum digit width per size (CAPTURING.md section 2 item 26)
    assert [CALIBRI[s]["0"] for s in MEASURED_SIZES] == [5, 5, 6, 6, 7, 8]
    assert [CALIBRI[s][" "] for s in MEASURED_SIZES] == [2, 2, 2, 3, 3, 4]


@pytest.mark.parametrize("size", MEASURED_SIZES)
def test_measured_size_uses_its_own_table(size: int) -> None:
    assert _advances("Calibri", float(size)) is CALIBRI[size]


def test_unmeasured_size_scales_the_nearest_table() -> None:
    # 13 lies between 12 and 14; the larger wins the tie and every advance is a whole point
    scaled = _advances("Calibri", 13.0)
    assert scaled["0"] == round(CALIBRI[14]["0"] * 13 / 14)
    assert all(adv == int(adv) for adv in scaled.values())
    assert _advances("Calibri", 18.0)["0"] == round(CALIBRI[16]["0"] * 18 / 16)


def test_column_budget_uses_the_size_digit_width() -> None:
    # points = MDW * w + 5 at every size (space_fill and width_general14 headers)
    assert column_budget(6, ("Calibri", 9)) == 30
    assert column_budget(6, ("Calibri", 14)) == 42
    assert column_budget(20, ("Calibri", 16)) == 160
    assert column_budget(8.43, ("Calibri", 14)) == 59  # round(7 * 8.43 + 5) = 64, less 5


def test_resolve_font_defaults_and_errors() -> None:
    assert resolve_font(None) == Font("Calibri", 11)
    assert resolve_font("calibri") == Font("Calibri", 11.0)
    assert resolve_font(("CALIBRI", 14)) == Font("Calibri", 14.0)
    with pytest.raises(LookupError, match="Calibri"):
        resolve_font("Comic Sans")


def test_cell_font_against_the_normal_font() -> None:
    # cell_font and general_cell_font rows: the column from the Normal font (Calibri 11), the
    # text measured in the cell's (CAPTURING.md section 2 item 29)
    normal = Font("Calibri", 11)
    assert format_value(88888888, "0", column_width=8, normal_font=normal) == "88888888"
    # six '#' of 7 points in the 48-point budget (cell_font, digits8 [Calibri 14] at width 8)
    cell = Font("Calibri", 14)
    assert format_value(88888888, "0", column_width=8, font=cell, normal_font=normal) == "######"
    # General counts characters in the cell font's digits: 36 points / 5 at Calibri 9
    assert format_value(1234567.891, "General", column_width=6, font=("Calibri", 9),
                        normal_font=normal) == "1234568"  # fmt: skip
    assert format_value(-0.001, "[Red]General", column_width=6, font=("Calibri", 14),
                        normal_font=normal) == "0"  # fmt: skip
    assert format_value(-0.001, "General", column_width=6, font=("Calibri", 14),
                        normal_font=normal) == "-0"  # fmt: skip


def test_glyph_measure_floors_the_sum() -> None:
    measure = glyph_measure(("Calibri", 11))
    assert measure("1,000.0") == 36  # fits width 6 (36 points) while -1234.6 is reduced
    assert measure("10000%") == 38


def test_space_fill_at_a_new_size() -> None:
    # space_fill14 row: the accounting zero section at width 6 in Calibri 14
    fmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    # 42 points: pads and ? blanks as 3-point spaces, the dash 4, eight fill spaces
    assert format_value(0, fmt, column_width=6, font=("Calibri", 14)) == "         -   "


@pytest.mark.parametrize(
    ("value", "fmt", "column_width", "font", "expected"),
    [
        # general_points11 and general_literals_points rows (CAPTURING.md section 2 item 27)
        (1234567.891, "General", 4.67, ("Calibri", 11), "1E+06"),
        (12345, "General", 4.67, ("Calibri", 11), "#####"),
        (12345, "( General )", 6.33, ("Calibri", 11), "( 1E+04 )"),
        (3.14159265358979, 'General" units"', 8.17, ("Calibri", 11), "3.14 units"),
        (-1.5, "( General )", 2.67, ("Calibri", 11), "( 0 )"),
        (-1.5e-07, "( General )", 2.67, ("Calibri", 11), "###"),
        (-0.05, "General;-General", 1, ("Calibri", 11), "0"),
        (-12345678, "General", 7.57, ("Calibri", 14), "-1.2E+07"),
    ],
)
def test_general_at_fractional_widths(
    value: float, fmt: str, column_width: float, font: tuple[str, int], expected: str
) -> None:
    assert format_value(value, fmt, column_width=column_width, font=font) == expected


def test_file_widths_and_the_default_column() -> None:
    # column_file11 and column_default_* headers (CAPTURING.md section 2 item 28)
    from xlformat import file_column_metrics

    assert file_column_metrics().points == 53  # baseColWidth 8, no defaultColWidth
    assert file_column_metrics().width == 8
    assert file_column_metrics(base_col_width=10).points == 65
    assert file_column_metrics(default_col_width=12).points == 72
    assert [file_column_metrics(w).points for w in (8.43, 60, 3.7109375, 0.75)] == [51, 360, 22, 4]
    assert file_column_metrics(1.75, font=("Calibri", 16)).width == 1.125  # the dialog: 1.13
    assert file_column_metrics(8.43).budget == 46


def test_column_metrics_round_trip_and_narrow_columns() -> None:
    from xlformat import column_metrics, file_column_metrics

    assert column_metrics(8.43).points == 56
    assert column_metrics(0.09).points == 1  # below one unit: round(w * (MDW + 5))
    for w in (0.5, 1.75, 8.43, 13, 255):
        m = file_column_metrics(w, font=("Calibri", 14))
        assert column_metrics(m.width, font=("Calibri", 14)).points == m.points
    assert format_value(1, "0", column_width=0.5) == ""  # no '#' fits
    assert format_value(1, "General", column_width=0.5) == "1"
    assert format_value(12, "General", column_width=0.5) == "#"


def test_the_255_character_limit() -> None:
    # limits rows: 256 characters with the fill, and 1E+300 under '0' (301 digits)
    assert format_value(1, "* 0", column_width=86) == "#" * 103
    assert format_value(1, "* 0", column_width=84) == " " * 249 + "1"
    assert format_value(1e300, "0") == "#" * 255
    assert format_value(1e300, "0", column_width=255) == "#" * 255
