"""Excel for Mac's faces: the recovered advances and the width rule against the faces captures.

``capture/face_fit.py`` must give ``src/xlformat/_faces_mac.py`` from ``tests/captures/faces/``.
Every row of the check captures (``faces_c<size>_<family>``: each glyph ten times, texts mixing
ASCII and other glyphs, long texts) must be the rule's autofit width: the text's width
(:func:`xlformat.text_width`; for a text holding a non-ASCII glyph the unrounded sum, rounded
half up), the padding ``2 * ceil(d / 4) + 3`` (d the cell face's widest
digit), twice the italic overhang, and the autofit bonus ``max(0, floor(W / D / 5) - 2)`` (W the
text's width plus one overhang, D the Normal font's digit).
"""

from __future__ import annotations

import math
from pathlib import Path

import pytest

import xlformat
from capture import face_fit, faces
from xlformat import _measure

FACES_DIR = Path(__file__).parent / "captures" / "faces"
# Segoe UI's widths follow no per-face advance (capture/face_fit.py); it raises LookupError
CHECKS = [p for p in sorted(FACES_DIR.glob("faces_c*.measure.tsv*")) if "segoe_ui" not in p.name]


def test_the_module_is_what_the_captures_give() -> None:
    face_fit.main(["--check"])


def autofit(text: str, font: xlformat.Font, normal: xlformat.Font) -> float:
    width = xlformat.text_width(text, font, platform="mac")
    if not text.isascii():  # autofit sums such a text unrounded
        f = _measure.resolve_font(font)
        exact = _measure._linear_advances(f.name, f.size, f.style)
        width = math.floor(sum(exact.get(c, exact[""]) for c in text) + 0.5)
    digit = xlformat.digit_width(font, platform="mac")
    over = xlformat.overhang(font, platform="mac")
    # the bonus counts the text and one overhang
    steps = (width + over) / xlformat.digit_width(normal, platform="mac") / 5
    return width + 2 * math.ceil(digit / 4) + 3 + 2 * over + max(0, math.floor(steps) - 2)


def _font(label: str) -> xlformat.Font:
    family, size, style = face_fit.parse_label(label)
    return xlformat.Font(family, size, "bold" in style, "italic" in style)


@pytest.mark.parametrize("path", CHECKS, ids=lambda p: p.name.split(".")[0])
def test_check_capture(path: Path) -> None:
    header, widths = faces.read_measure(path)
    normal = _font(header)
    misses = []
    for (label, text), points in widths.items():
        font = _font(label)
        if set(text) <= KNOWN_BLANK:
            continue
        if autofit(text, font, normal) != points:
            misses.append((label, text[:24], autofit(text, font, normal), points))
    assert not misses, misses[:10]


# Runs autofit does not resolve: a run of spaces, no-break spaces or soft hyphens.
KNOWN_BLANK = frozenset(" \xa0\xad")


def test_substitutes_measure_as_their_face() -> None:
    assert xlformat.font_substitutes(platform="mac") == {"MS Sans Serif": "Times New Roman"}
    assert xlformat.font_substitutes(platform="windows") == {}
    for name, face in xlformat.font_substitutes(platform="mac").items():
        text = "Total: 1,234.56 €"
        assert xlformat.text_width(text, (name, 13)) == xlformat.text_width(text, (face, 13))


def test_sizes_run_from_1_to_192_points() -> None:
    assert xlformat.text_width("0", ("Verdana", 192)) > 0
    for size in (0.5, 192.5, 409):
        with pytest.raises(LookupError):
            xlformat.text_width("0", ("Verdana", size))
    with pytest.raises(LookupError):
        xlformat.text_width("0", ("Geneva", 11))
    with pytest.raises(LookupError):
        xlformat.text_width("0", ("Segoe UI", 11))


def test_measured_fonts_lists_every_face_in_four_styles() -> None:
    fonts = xlformat.measured_fonts(platform="mac")
    assert list(fonts) == list(faces.FAMILIES)
    assert fonts["Calibri"]["bold italic"] == (9, 10, 11, 12, 13, 14, 16, 18, 20, 24)
    assert fonts["Verdana"]["regular"] == (10, 13)
    assert set(xlformat.measured_fonts(platform="windows")) == {"Calibri", "Aptos Narrow", "Arial"}


def test_fit_measures_every_glyph_whole() -> None:
    # autofit sums a text holding a non-ASCII glyph unrounded (the check captures); fit, # and
    # fill decisions round every glyph (numeric_core_width_arial10: '€ 99.00 ' needs 7 units)
    font = xlformat.Font("Arial", 10)
    whole = _measure._whole_advances("Arial", 10.0, "regular")
    assert xlformat.text_width("€ 99.00 ", font) == sum(whole[c] for c in "€ 99.00 ")
    assert xlformat.glyph_widths("é", font) == (whole["é"],)
