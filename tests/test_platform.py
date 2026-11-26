"""``platform=`` and ``regional_settings=``: the defaults change nothing; Windows works where its
evidence is captured (widths, rows, built-in formats and ``TEXT`` at 96 dpi) and raises
elsewhere."""

from __future__ import annotations

import dataclasses
import inspect
from collections.abc import Callable

import pytest

import xlformat
from xlformat import DEFAULT_PLATFORM, NumberFormat, Platform, RegionalSettings

FONT = ("Calibri", 11)
# every function whose result can differ by platform, with one call of it
CALLS: dict[str, Callable[..., object]] = {
    "format_value": lambda **k: xlformat.format_value(1234.5, "#,##0.00", column_width=8, **k),
    "format_cell": lambda **k: xlformat.format_cell(-1, "0;[Red]0", column_width=8, **k),
    "text_function": lambda **k: xlformat.text_function(45000, "m/d/yy", **k),
    "builtin_format": lambda **k: xlformat.builtin_format(14, **k),
    "text_width": lambda **k: xlformat.text_width("Total", FONT, **k),
    "glyph_widths": lambda **k: xlformat.glyph_widths("Total", FONT, **k),
    "digit_width": lambda **k: xlformat.digit_width(FONT, **k),
    "overhang": lambda **k: xlformat.overhang(("Calibri", 11, False, True), **k),
    "measured_fonts": lambda **k: xlformat.measured_fonts(**k),
    "column_metrics": lambda **k: xlformat.column_metrics(8.43, FONT, **k),
    "column_metrics_from_points": lambda **k: xlformat.column_metrics_from_points(64, FONT, **k),
    "file_column_metrics": lambda **k: xlformat.file_column_metrics(10, normal_font=FONT, **k),
    "row_metrics": lambda **k: xlformat.row_metrics(FONT, **k),
    "vertical_metrics": lambda **k: xlformat.vertical_metrics(FONT, **k),
    "ink_metrics": lambda **k: xlformat.ink_metrics(FONT, **k),
    "rule_metrics": lambda **k: xlformat.rule_metrics(FONT, zoom=150, **k),
    "text_insets": lambda **k: xlformat.text_insets(FONT, **k),
    "autofit_width": lambda **k: xlformat.autofit_width("Total", FONT, **k),
    "NumberFormat.format": lambda **k: NumberFormat.parse("0.0").format(2.25, **k),
    "NumberFormat.format_cell": lambda **k: NumberFormat.parse("0.0").format_cell(2.25, **k),
    "NumberFormat.text_function": lambda **k: NumberFormat.parse("0.0").text_function(2.25, **k),
}
REGIONAL = {"format_value", "format_cell", "text_function", "builtin_format"} | {
    n for n in CALLS if n.startswith("NumberFormat.")
}
# the functions with Windows evidence at 96 dpi: widths, rows, built-in formats and TEXT
# (CAPTURING.md section 2 items 65 to 72)
WINDOWS = {
    "format_value", "format_cell", "text_width", "glyph_widths", "digit_width", "overhang",
    "measured_fonts", "column_metrics", "column_metrics_from_points", "file_column_metrics",
    "NumberFormat.format", "NumberFormat.format_cell", "row_metrics", "builtin_format",
    "text_function", "NumberFormat.text_function", "vertical_metrics", "ink_metrics",
    "rule_metrics", "text_insets", "autofit_width",
}  # fmt: skip


def _function(name: str) -> Callable[..., object]:
    owner, _, attr = name.rpartition(".")
    return getattr(NumberFormat, attr) if owner else getattr(xlformat, name)


@pytest.mark.parametrize("name", sorted(CALLS))
def test_every_function_takes_the_settings(name: str) -> None:
    params = inspect.signature(_function(name)).parameters
    assert params["platform"].kind is inspect.Parameter.KEYWORD_ONLY
    assert params["platform"].default is None
    assert ("regional_settings" in params) == (name in REGIONAL)


@pytest.mark.parametrize("name", sorted(CALLS))
def test_the_default_is_windows(name: str) -> None:
    call = CALLS[name]
    nothing = call()
    assert call(platform=None) == nothing
    assert call(platform="windows") == nothing
    assert call(platform=Platform.WINDOWS) == nothing
    if name in REGIONAL:
        assert call(regional_settings=RegionalSettings.EN_US_WINDOWS) == nothing
        assert call(platform="windows", regional_settings=None) == nothing
    assert DEFAULT_PLATFORM == "windows"


@pytest.mark.parametrize("name", sorted(CALLS))
def test_the_mac_is_one_argument_away(name: str) -> None:
    call = CALLS[name]
    mac = call(platform="mac")
    assert call(platform=Platform.MAC) == mac
    if name in REGIONAL:
        assert call(platform="mac", regional_settings=RegionalSettings.EN_US_MAC) == mac


def test_what_the_default_changed() -> None:
    # Calibri 11: a 7-pixel digit (5.25 points) on Windows, 6 points on the Mac
    assert xlformat.digit_width() == 5.25
    assert xlformat.digit_width(platform="mac") == 6.0
    assert xlformat.builtin_format(14) == "m/d/yyyy"
    assert xlformat.builtin_format(14, platform="mac") == "m/d/yy"
    assert xlformat.row_metrics().height == 15.0


@pytest.mark.parametrize("name", sorted(CALLS))
def test_windows_raises_until_its_evidence_is_captured(name: str) -> None:
    other_dpi = dataclasses.replace(Platform.WINDOWS, dpi=120)
    unsupported = [other_dpi, dataclasses.replace(Platform.MAC, long_fill="cap")]
    if name in WINDOWS:
        CALLS[name](platform="windows")
        CALLS[name](platform=Platform.WINDOWS)
    else:
        unsupported += ["windows", Platform.WINDOWS]
    for platform in unsupported:
        with pytest.raises(LookupError, match=f"^{name} has no captured evidence"):
            CALLS[name](platform=platform)


def test_windows_measures_in_pixels_reported_in_points() -> None:
    # Calibri 11 at 96 dpi: a digit of 7 pixels (measure11 on Windows), 5.25 points
    assert xlformat.digit_width(FONT, platform="windows") == 7 * 72 / 96
    windows = xlformat.measured_fonts(platform="windows")
    assert len(windows) == 25
    assert set(windows) | {"Courier", "Helvetica"} == set(xlformat.measured_fonts(platform="mac"))
    assert windows["Times New Roman"]["bold italic"] == tuple(range(8, 25))
    with pytest.raises(LookupError, match="for Excel for Windows for font 'Courier'"):
        xlformat.text_width("x", ("Courier", 11), platform="windows")


def test_platform_names_and_fields() -> None:
    with pytest.raises(ValueError, match="'mac', 'windows' or a Platform"):
        xlformat.text_width("x", platform="linux")
    with pytest.raises(ValueError, match="dpi is a Windows setting"):
        Platform("mac", dpi=96)
    with pytest.raises(ValueError, match="positive dpi"):
        Platform("windows")
    assert dataclasses.replace(Platform.WINDOWS, dpi=120).dpi == 120
    assert Platform.WINDOWS.default_column_step == 8
    assert Platform.WINDOWS.long_fill == "cap"


def test_regional_settings() -> None:
    assert xlformat.builtin_format(14, platform="mac") == "m/d/yy"
    assert xlformat.builtin_format(22, platform="mac") == "m/d/yy h:mm"
    windows = RegionalSettings.EN_US_WINDOWS
    assert xlformat.builtin_format(14, regional_settings=windows) == "m/d/yyyy"
    assert xlformat.builtin_format(22, regional_settings=windows) == "m/d/yyyy h:mm"
    assert xlformat.builtin_format(15, regional_settings=windows) == "d-mmm-yy"
    custom = RegionalSettings(short_date="yyyy-mm-dd")
    assert xlformat.builtin_format(14, regional_settings=custom) == "yyyy-mm-dd"
    mac = RegionalSettings.EN_US_MAC
    assert xlformat.builtin_format(14, platform="mac", regional_settings=mac) == "m/d/yy"
    with pytest.raises(ValueError, match="only en-US regional settings"):
        RegionalSettings(decimal_separator=",")
