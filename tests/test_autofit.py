"""``autofit_width`` against every autofit capture of both platforms.

The Mac: the faces captures (CAPTURING.md section 2 item 82): the check captures, the padding
captures at every whole size from 6 to 192 and the half sizes, and the switch captures from 193
to 320 points. Windows: the long-text captures (item 90), every pixel from 10 to 45 widths of the
Normal font's digit in 13 Normal fonts, and the faces captures (item 85, in
``tests/test_faces_windows.py``).
"""

from __future__ import annotations

import ast
import gzip
from pathlib import Path

import pytest

import xlformat
from capture import autofit_windows, faces

FACES_DIR = Path(__file__).parent / "captures" / "faces"
# sizes 1 to 5 (faces_p1to5) and the recovery sizes 383 to 409 (faces_r*) are not answered by the
# rule; Segoe UI raises LookupError
MAC_FILES = [
    p
    for p in sorted(FACES_DIR.glob("faces_[cps]*.measure.tsv*"))
    if "segoe_ui" not in p.name and "p1to5" not in p.name
]
WINDOWS_FILES = sorted(autofit_windows.AUTOFIT_DIR.glob("autofitw_*.measure.tsv.gz"))
BLANK = frozenset(" \xa0\xad")  # runs autofit does not resolve (tests/test_faces.py)


def _font(label: str) -> xlformat.Font:
    """A measure file's font label: ``Calibri 11``, ``Verdana 12 bold italic``."""
    words = label.split()
    style = [w for w in words if w in ("bold", "italic")]
    rest = words[: len(words) - len(style)]
    return xlformat.Font(" ".join(rest[:-1]), float(rest[-1]), "bold" in style, "italic" in style)


@pytest.mark.parametrize("path", MAC_FILES, ids=lambda p: p.name.split(".")[0])
def test_mac_autofit_captures(path: Path) -> None:
    header, widths = faces.read_measure(path)
    normal = _font(header)
    misses = []
    for (label, text), points in widths.items():
        if set(text) <= BLANK or label.startswith("Segoe UI"):
            continue
        got = xlformat.autofit_width(text, _font(label), normal_font=normal, platform="mac")
        if got != points:
            misses.append((label, text[:24], got, points))
    assert not misses, misses[:10]


def _windows_rows(path: Path) -> tuple[xlformat.Font, list[tuple[str, str, float]]]:
    lines = gzip.decompress(path.read_bytes()).decode("utf-8").split("\n")
    normal = _font(lines[0].split("font: ", 1)[1].split(";")[0])
    assert lines[2].startswith("value_type\t")
    rows = []
    for line in lines[3:]:
        if line:
            fields = line.split("\t")
            rows.append((fields[5], ast.literal_eval(fields[1]), float(fields[4])))
    return normal, rows


def test_every_windows_suite_is_captured() -> None:
    assert [p.name.split(".")[0] for p in WINDOWS_FILES] == sorted(autofit_windows.SUITE_NAMES)


@pytest.mark.parametrize("path", WINDOWS_FILES, ids=lambda p: p.name.split(".")[0])
def test_windows_autofit_captures(path: Path) -> None:
    normal, rows = _windows_rows(path)
    digit = xlformat.digit_width(normal, platform="windows")
    reach = 0.0
    for label, text, points in rows:
        font = _font(label)
        got = xlformat.autofit_width(text, font, normal_font=normal, platform="windows")
        assert got == points, (label, text, got, points)
        over = xlformat.overhang(font, platform="windows")
        reach = max(reach, (xlformat.text_width(text, font, platform="windows") + over) / digit)
    assert reach >= 45


def test_the_bonus_ladders() -> None:
    # Calibri 11: a 7-pixel digit on Windows, 6 points on the Mac; the padding is 7 on both
    zero = xlformat.Font("Calibri", 11)
    win = [xlformat.autofit_width("0" * k, zero, platform="windows") * 4 / 3 for k in (11, 12, 16)]
    assert win == [77 + 7, 84 + 7 + 1, 112 + 7 + 2]  # one pixel from 12 digits, one more every 4
    mac = [xlformat.autofit_width("0" * k, zero) for k in (14, 15, 20)]
    assert mac == [84 + 7, 90 + 7 + 1, 120 + 7 + 2]  # one point from 15 digits, one more every 5


def test_autofit_differs_from_text_width_where_autofit_measures_differently() -> None:
    # the Mac's autofit sums a non-ASCII text unrounded; text_width rounds each glyph
    text = "é" * 10  # 5 whole points each in Calibri 11, 5.47 unrounded: 54.7 for ten
    assert xlformat.text_width(text) == 50.0
    assert xlformat.autofit_width(text) == 55 + 7


def test_one_line_only() -> None:
    for text in ("a\nb", "a\rb", "a\tb", "\r\n"):
        with pytest.raises(ValueError, match="one line"):
            xlformat.autofit_width(text)


def test_empty_text() -> None:
    assert xlformat.autofit_width("") == 0.0
    assert xlformat.autofit_width("", platform="windows") == 0.0
    with pytest.raises(LookupError):
        xlformat.autofit_width("", ("Geneva", 11))


def test_unmeasured_fonts_raise() -> None:
    with pytest.raises(LookupError):
        xlformat.autofit_width("x", ("Geneva", 11))
    with pytest.raises(LookupError):
        xlformat.autofit_width("x", normal_font=("Geneva", 11))
    with pytest.raises(LookupError):
        xlformat.autofit_width("x", ("Courier", 11), platform="windows")
    with pytest.raises(ValueError, match="platform"):
        xlformat.autofit_width("x", platform="linux")
