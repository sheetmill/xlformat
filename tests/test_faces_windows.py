"""Excel for Windows' whole-pixel advances per ppem (CAPTURING.md section 2 item 85).

``capture/face_fit_windows.py`` must give ``src/xlformat/_faces_windows.py`` from
``tests/captures/windows/faces/``, and every captured autofit width must be the package's
measure of its text plus the autofit padding and bonus, through the public functions.
"""

from __future__ import annotations

import math
import shutil

import pytest

import xlformat
from capture import face_fit_windows, faces_windows

WINDOWS = xlformat.Platform.WINDOWS


@pytest.mark.skipif(
    shutil.which("ruff") is None,
    reason="ruff (development group) formats the regenerated module before comparing it",
)
def test_the_module_is_what_the_captures_give() -> None:
    face_fit_windows.main(["--check"])


def _px(points: float) -> int:
    px = points * 96 / 72
    assert px == round(px)
    return round(px)


@pytest.mark.parametrize("name", faces_windows.captured()[1])
def test_every_autofit_width_is_advances_padding_and_bonus(name: str) -> None:
    normal_label, by_label = face_fit_windows.read(
        faces_windows.FACES_DIR / f"{name}.measure.tsv.gz"
    )
    family, size, _ = face_fit_windows.parse_label(normal_label)
    normal_digit = _px(xlformat.digit_width((family, size), platform=WINDOWS))
    checked = 0
    for label, widths in by_label.items():
        family, size, style = face_fit_windows.parse_label(label)
        font = xlformat.Font(family, size, "bold" in style, "italic" in style)
        digit = _px(xlformat.digit_width(font, platform=WINDOWS))
        over = _px(xlformat.overhang(font, platform=WINDOWS))
        pad = 2 * math.ceil(digit / 4) + 3 + 2 * over
        for text, width in widths.items():
            px = _px(xlformat.text_width(text, font, platform=WINDOWS))
            assert width == px + pad + face_fit_windows.bonus(px, over, normal_digit), (label, text)
            checked += 1
    assert checked == len(by_label) * (
        len(faces_windows.faces.GLYPHS) + 1 + len(faces_windows.CHECKS)
    )


def test_a_size_draws_at_its_ppem() -> None:
    # 8 and 8.5 points both draw at ppem 11, 9.5 and 10 at 13 (10.67 and 11.33, 12.67 and 13.33
    # rounded): each pair autofits alike in every glyph
    for a, b in ((8, 8.5), (9.5, 10), (23, 23.5)):
        for style in ("regular", "bold italic"):
            fa = xlformat.Font("Times New Roman", a, "bold" in style, "italic" in style)
            fb = xlformat.Font("Times New Roman", b, "bold" in style, "italic" in style)
            text = "".join(faces_windows.faces.GLYPHS)
            assert xlformat.glyph_widths(text, fa, platform=WINDOWS) == xlformat.glyph_widths(
                text, fb, platform=WINDOWS
            )
    # every whole size whose ppem was captured
    assert xlformat.measured_fonts(platform=WINDOWS)["Calibri"]["regular"] == tuple(range(8, 25))
