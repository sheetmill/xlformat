"""Autofit widths of long texts on Excel for Windows: the long-text bonus at every pixel.

    uv run python -m capture.autofit_windows count
    uv run python -m capture.job write autofit_windows JOB_ID OUT_DIR
    uv run python -m capture.job read autofit_windows JOB_DIR RESULT_DIR

One workbook per Normal font of :data:`NORMALS` (``autofitw_<face>_<size>``), chosen so that
the Normal font's widest digit D runs from 6 to 17 pixels at 96 dpi. Each workbook autofits,
one text a column, ASCII texts whose advances reach every whole pixel from 10 to 45 widths of
D: a run of ``0`` and up to three letters, the advances taken from xlformat's measured Windows
tables (sizes 8 to 24, CAPTURING.md item 85). A few workbooks also hold the same pixel range
in the Normal face's italic (whose overhang the bonus may count) and in a cell font of another
size than the Normal font (whose digit the bonus may follow). The autofit width less the
advances and the padding is the bonus; the texts reach the bonus's steps at several digit
widths, which separates its rule from the Mac's ``max(0, floor(W / D / 5) - 2)`` shifted to
start at 12 widths.

``capture.job read`` writes ``tests/captures/windows/autofit/<suite>.measure.tsv.gz``.
"""

from __future__ import annotations

import itertools
import string
import sys
from functools import cache
from pathlib import Path

from capture import cases as xc
from capture import faces

ROOT = Path(__file__).resolve().parents[1]
AUTOFIT_DIR = ROOT / "tests" / "captures" / "windows" / "autofit"
GROUP = "autofit_windows"
# Normal fonts (face, size): Calibri's digit is 6, 7, 8, 9, 10, 11, 12, 14 and 16 pixels at these
# sizes; Arial 10, Times New Roman 12 and Verdana 11 and 20 repeat or extend those digits in
# other faces
NORMALS = (
    ("Calibri", 9),
    ("Calibri", 11),
    ("Calibri", 12),
    ("Calibri", 13),
    ("Calibri", 14),
    ("Calibri", 16),
    ("Calibri", 18),
    ("Calibri", 20),
    ("Calibri", 24),
    ("Arial", 10),
    ("Times New Roman", 12),
    ("Verdana", 11),
    ("Verdana", 20),
)
# workbooks that also hold the Normal face's italic at the Normal size
ITALIC = frozenset({("Calibri", 11), ("Calibri", 16), ("Calibri", 24), ("Arial", 10)})
# workbooks that also hold a cell font of another size (face, size) than the Normal font
CROSS = {("Calibri", 11): ("Calibri", 20), ("Calibri", 20): ("Calibri", 11)}
LOW, HIGH = 10, 45  # the pixel range, in widths of the Normal font's digit
LETTERS = string.ascii_letters


def _windows_font(family: str, size: float, italic: bool = False) -> object:
    import xlformat

    return xlformat.Font(family, size, italic=italic)


@cache
def advances(family: str, size: float, italic: bool = False) -> dict[str, int]:
    """Whole-pixel advances of ``0`` and the letters in the face, from xlformat's tables."""
    import xlformat

    font = _windows_font(family, size, italic)
    widths = xlformat.glyph_widths("0" + LETTERS, font, platform="windows")  # type: ignore[arg-type]
    return {c: round(w * 4 / 3) for c, w in zip("0" + LETTERS, widths, strict=True)}


def digit(family: str, size: float) -> int:
    """The face's widest digit in pixels."""
    import xlformat

    return round(xlformat.digit_width(_windows_font(family, size), platform="windows") * 4 / 3)  # type: ignore[arg-type]


def text_for(px: int, adv: dict[str, int]) -> str | None:
    """A run of ``0`` and at most three letters whose advances sum to ``px``; None if none."""
    zero = adv["0"]
    tails: dict[int, str] = {0: ""}
    for n in (1, 2, 3):
        for combo in itertools.combinations_with_replacement(LETTERS, n):
            tails.setdefault(sum(adv[c] for c in combo), "".join(combo))
    for total, tail in sorted(tails.items(), key=lambda kv: len(kv[1])):
        rest = px - total
        if rest >= 0 and rest % zero == 0:
            return "0" * (rest // zero) + tail
    return None


def texts(normal_digit: int, adv: dict[str, int]) -> list[str]:
    """One text per whole pixel from LOW to HIGH widths of the Normal digit."""
    out = []
    for px in range(LOW * normal_digit, HIGH * normal_digit + 1):
        t = text_for(px, adv)
        if t is None:
            msg = f"no text of {px} px"
            raise ValueError(msg)
        out.append(t)
    return out


def suite_name(family: str, size: int) -> str:
    return f"autofitw_{faces._key(family)}_{size}"


def suite(family: str, size: int) -> xc.Suite:
    d = digit(family, size)
    by_font = {f"{family}|{size}|regular": texts(d, advances(family, size))}
    if (family, size) in ITALIC:
        by_font[f"{family}|{size}|italic"] = texts(d, advances(family, size, True))
    if (family, size) in CROSS:
        cf, cs = CROSS[(family, size)]
        by_font[f"{cf}|{cs}|regular"] = texts(d, advances(cf, cs))
    return faces._suite(suite_name(family, size), family, size, by_font)


SUITE_NAMES = tuple(suite_name(f, s) for f, s in NORMALS)
JOB_GROUPS: dict[str, tuple[str, ...]] = {GROUP: SUITE_NAMES}


def register() -> dict[str, tuple[str, ...]]:
    """Add the group's suites to :data:`capture.cases.SUITES`; the group's names, in job order."""
    for family, size in NORMALS:
        xc.SUITES[suite_name(family, size)] = lambda f=family, s=size: suite(f, s)
    return JOB_GROUPS


def main(argv: list[str]) -> None:
    if argv != ["count"]:
        sys.exit(__doc__)
    built = [suite(f, s) for f, s in NORMALS]
    for s in built:
        print(s.name, len(s.cases), "columns", max(len(c.value) for c in s.cases), "chars at most")
    print(f"{len(built)} workbooks, {sum(len(s.cases) for s in built)} columns")


if __name__ == "__main__":
    main(sys.argv[1:])
