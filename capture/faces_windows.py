"""Autofit captures on Excel for Windows from which whole-pixel glyph advances per ppem are read.

    uv run python -m capture.faces_windows count
    uv run python -m capture.job write faces_windows JOB_ID OUT_DIR
    uv run python -m capture.job read faces_windows JOB_DIR RESULT_DIR

One workbook per face of :data:`FACES` and whole size of :data:`SIZES`
(``facesw_<face>_<size>``), its Normal font the face's regular style at that size. Its cell
fonts are the face's four styles at the size and at the half size above it (none above 24), so
that the workbooks reach every ppem from 11 to 32 at 96 dpi (``ppem = size * 96 / 72``, a half
rounded up): 8 points is ppem 10.67, 8.5 is 11.33, 24 is 32. Per cell font, each in its own
column: every glyph of :data:`capture.faces.GLYPHS` five times (three times for
:data:`WIDE`), the space, the no-break space and the soft hyphen five times between two ``x``,
``0`` ten times and the mixed texts of :data:`CHECKS`. Every text stays under
:data:`capture.faces.LIMIT_DIGITS` widths of the Normal font's digit, so that no column earns
the autofit bonus. A five-run's width less the padding is five whole advances; the padding is
twice the ``0`` five-run's width less the ten-run's.

``capture/face_fit.py --windows`` reads the measure files the job's ``read`` writes to
``tests/captures/windows/faces/``.
"""

from __future__ import annotations

import sys
from pathlib import Path

from capture import cases as xc
from capture import faces

ROOT = Path(__file__).resolve().parents[1]
FACES_DIR = ROOT / "tests" / "captures" / "windows" / "faces"
GROUP = "faces_windows"
FACES = ("Aptos Narrow", "Calibri", "Arial", "Times New Roman")
SIZES = tuple(range(8, 25))
STYLES = faces.STYLES
WIDE = frozenset("WM@%\u00c6\u0152\u2026\u2030")  # W M @ % AE OE ellipsis per mille: three copies
BLANK = frozenset(" \u00a0\u00ad")  # the space, the no-break space, the soft hyphen
BRACKET = "x"
RUNS = {"wide": 3, "plain": 5, "blank": 5}
# the mixed texts of capture.faces but the two that reach 13 digit widths in Times New Roman bold
CHECKS = tuple(
    t
    for t in faces.MIXED
    if t not in ("\u00a9 2026 Soci\u00e9t\u00e9", "Z\u00fcrich\u2013Gen\u00e8ve")
)


def copies(glyph: str) -> int:
    """How many copies of a glyph its run holds."""
    return RUNS["wide"] if glyph in WIDE else RUNS["plain"]


def run_text(glyph: str) -> str:
    """The text whose autofit width gives the glyph's advance."""
    if glyph in BLANK:
        return BRACKET + glyph * RUNS["blank"] + BRACKET
    return glyph * copies(glyph)


def cell_sizes(size: int) -> tuple[float, ...]:
    """The cell font sizes of the workbook whose Normal font is ``size``."""
    return (float(size),) if size == SIZES[-1] else (float(size), size + 0.5)


def suite_name(family: str, size: int) -> str:
    return f"facesw_{faces._key(family)}_{size}"


def suite(family: str, size: int) -> xc.Suite:
    texts = [run_text(g) for g in faces.GLYPHS] + ["0" * 10, *CHECKS]
    return faces._suite(
        suite_name(family, size),
        family,
        size,
        {f"{family}|{s:g}|{style}": texts for s in cell_sizes(size) for style in STYLES},
    )


SUITE_NAMES = tuple(suite_name(f, s) for f in FACES for s in SIZES)


def register() -> tuple[str, ...]:
    """Add the group's suites to :data:`capture.cases.SUITES`; their names, in job order."""
    for family in FACES:
        for size in SIZES:
            xc.SUITES[suite_name(family, size)] = lambda f=family, s=size: suite(f, s)
    return SUITE_NAMES


def main(argv: list[str]) -> None:
    if argv != ["count"]:
        sys.exit(__doc__)
    built = [suite(f, s) for f in FACES for s in SIZES]
    columns = sum(len(s.cases) for s in built)
    fonts = max(len({c.font for c in s.cases}) for s in built)
    print(f"{len(built)} workbooks, {columns} columns, at most {fonts} cell fonts a workbook")


if __name__ == "__main__":
    main(sys.argv[1:])
