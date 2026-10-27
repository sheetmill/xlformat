"""Excel-mode measure: measured sizes are exact tables, other sizes scale the nearest one."""

import pytest

from xlformat import format_value
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
    assert resolve_font(None) == ("Calibri", 11)
    assert resolve_font("calibri") == ("Calibri", 11.0)
    with pytest.raises(LookupError, match="Calibri"):
        resolve_font("Aptos Narrow")


def test_glyph_measure_floors_the_sum() -> None:
    measure = glyph_measure(("Calibri", 11))
    assert measure("1,000.0") == 36  # fits width 6 (36 points) while -1234.6 is reduced
    assert measure("10000%") == 38


def test_space_fill_at_a_new_size() -> None:
    # space_fill14 row: the accounting zero section at width 6 in Calibri 14
    fmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    # 42 points: pads and ? blanks as 3-point spaces, the dash 4, eight fill spaces
    assert format_value(0, fmt, column_width=6, font=("Calibri", 14)) == "         -   "
