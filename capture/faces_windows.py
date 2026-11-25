"""Autofit captures on Excel for Windows from which whole-pixel glyph advances per ppem are read.

    uv run python -m capture.faces_windows count [GROUP ...]
    uv run python -m capture.job write GROUP JOB_ID OUT_DIR
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

``GROUP`` one of :data:`JOB_GROUPS`. The probe (``faces_windows_probe``) is one workbook naming
every family of :data:`PROBED`, whose header lists those the machine lacks. Each measuring group
(:data:`GROUP_FACES`) holds one workbook per face and whole size of :data:`SIZES`
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
PROBE_GROUP = "faces_windows_probe"
PROBE_SUITE = "facesw_probe"
# the measuring job groups and their families, in job order
FIRST = ("Aptos Narrow", "Calibri", "Arial", "Times New Roman")
# the other families of capture.faces, named by the probe so that its header lists the missing
PROBED = tuple(family for family in faces.FAMILIES if family not in FIRST)
MISSING = ("Courier", "Helvetica")  # the probe's fonts_missing: no font file on the machine
_MORE = tuple(family for family in PROBED if family not in MISSING)
GROUP_FACES: dict[str, tuple[str, ...]] = {
    GROUP: FIRST,
    **{f"{GROUP}_{k + 2}": _MORE[7 * k : 7 * k + 7] for k in range(3)},
}
FACES = tuple(family for families in GROUP_FACES.values() for family in families)
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


def probe_suite() -> xc.Suite:
    """One workbook naming every family of :data:`PROBED` in four styles at 11 points.

    The job's header lists the faces the machine has no font file for (``fonts_missing``).
    """
    texts = {f"{family}|11|{style}": ["0" * 5, "0" * 10] for family in PROBED for style in STYLES}
    return faces._suite(PROBE_SUITE, "Calibri", 11, texts)


SUITE_NAMES = tuple(suite_name(f, s) for f in FACES for s in SIZES)
JOB_GROUPS: dict[str, tuple[str, ...]] = {
    PROBE_GROUP: (PROBE_SUITE,),
    **{g: tuple(suite_name(f, s) for f in fs for s in SIZES) for g, fs in GROUP_FACES.items()},
}


def captured() -> tuple[tuple[str, ...], tuple[str, ...]]:
    """The families and suite names of the measuring groups whose job manifest is committed."""
    groups = [g for g in GROUP_FACES if (FACES_DIR.parent / f"{g}.job.json").exists()]
    families = tuple(f for g in groups for f in GROUP_FACES[g])
    return families, tuple(suite_name(f, s) for f in families for s in SIZES)


def register() -> dict[str, tuple[str, ...]]:
    """Add the groups' suites to :data:`capture.cases.SUITES`; each group's names, in job order."""
    for family in FACES:
        for size in SIZES:
            xc.SUITES[suite_name(family, size)] = lambda f=family, s=size: suite(f, s)
    xc.SUITES[PROBE_SUITE] = probe_suite
    return JOB_GROUPS


def main(argv: list[str]) -> None:
    if not argv or argv[0] != "count" or not set(argv[1:]) <= set(GROUP_FACES):
        sys.exit(__doc__)
    families = [f for g in argv[1:] or GROUP_FACES for f in GROUP_FACES[g]]
    built = [suite(f, s) for f in families for s in SIZES]
    columns = sum(len(s.cases) for s in built)
    fonts = max(len({c.font for c in s.cases}) for s in built)
    print(f"{len(built)} workbooks, {columns} columns, at most {fonts} cell fonts a workbook")


if __name__ == "__main__":
    main(sys.argv[1:])
