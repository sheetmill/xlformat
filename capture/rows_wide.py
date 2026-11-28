"""Row heights on Excel for Windows across sizes: the face's height H and descent D per size.

    uv run python -m capture.job write rows_wide|rows_ppem|rows_more JOB_ID OUT_DIR
    uv run python -m capture.job read rows_wide|rows_ppem|rows_more JOB_DIR RESULT_DIR
    uv run python -m capture.rows_wide [--check]     # fit; --check reports, writes nothing

Group ``rows_wide`` (CAPTURING.md item 86): one workbook per family and style of :data:`TARGETS`
(20), its Normal font Calibri 11 (H 20, D 5 pixels at 96 dpi, CAPTURING.md section 2 item 69).
Per size of the family, five rows: a cell in the face wrapped to 1, 2, 3 and 4 lines (one short
glyph a line), and a one-line cell in the face beside a Normal-font cell wrapped to
:func:`normal_lines` lines, tall enough to set the row's top. The job reads each row's height.
``read`` writes ``tests/captures/windows/row_height_wide.layout.tsv.gz``.

Group ``rows_ppem`` (CAPTURING.md item 89): the same design at one size for each ppem from 1 to
96 (1 to 72 points) that ``rows_wide`` leaves out, plus three half sizes per face whose ppems
``rows_wide`` already holds, as a check; its wrapped cells have lines enough to rise past the
Normal row at small sizes. ``read`` writes
``tests/captures/windows/row_height_ppem.layout.tsv.gz``.

Group ``rows_more`` (CAPTURING.md item 91): the same design at a size for each ppem from 1 to 96
for :data:`MORE_FAMILIES`, in four styles (a lower-case ``arial`` regular only, to test that
Excel matches a face's name ignoring case). ``read`` writes
``tests/captures/windows/row_height_more.layout.tsv.gz``.

Model (item 69): a row is ``max(lines * H - D) + max D`` over its cells, the Normal font a
one-line cell among them. Then, with the Normal font's ``Hn`` and ``Dn``: the wrapped rows step by
H; the row beside the tall Normal cell is ``m * Hn - Dn + max(Dn, D)``, which gives ``max(Dn, D)``;
the two-line row ``2H - D + max(Dn, D)`` gives D. The one-line row and the four-line row check
them.
"""

from __future__ import annotations

import csv
import gzip
import math
import sys
from pathlib import Path

from capture.cases import CellFont
from capture.row_cases import Cell, RowCase, _lines

ROOT = Path(__file__).resolve().parents[1]
LAYOUT = ROOT / "tests" / "captures" / "windows" / "row_height_wide.layout.tsv.gz"
GROUP = "rows_wide"
NORMAL = CellFont("Calibri", 11)
NORMAL_HD = (20, 5)  # Calibri 11's H and D on Windows (item 69)
STYLES = ((False, False), (True, False), (False, True), (True, True))
CALIBRI_SIZES = tuple(s / 2 for s in range(12, 81)) + tuple(float(s) for s in range(41, 73))
OTHER_SIZES = tuple(float(s) for s in range(6, 41))
TARGETS: dict[str, tuple[float, ...]] = {
    "Calibri": CALIBRI_SIZES,
    "Aptos Narrow": OTHER_SIZES,
    "Arial": OTHER_SIZES,
    "Times New Roman": OTHER_SIZES,
    "Courier New": OTHER_SIZES,
}
WRAPPED = (1, 2, 3, 4)
# group rows_ppem: one half-point size for each ppem from 1 to 96 at 96 dpi (1 to 72 points)
# that the rows_wide sizes leave out, and per face three half sizes whose ppems rows_wide holds
LAYOUT_PPEM = ROOT / "tests" / "captures" / "windows" / "row_height_ppem.layout.tsv.gz"
GROUP_PPEM = "rows_ppem"
PPEMS = range(1, 97)
CHECK_CANDIDATES = (6.5, 12.5, 30.5, 50.5, 60.5, 66.5)


def ppem(size: float) -> int:
    """The ppem a size draws at in Excel for Windows at 96 dpi, a half rounded up."""
    return math.floor(size * 4 / 3 + 0.5)


def _ppem_sizes(family: str) -> tuple[float, ...]:
    """A size for each ppem rows_wide leaves out (whole sizes first), then three checks."""
    held = {ppem(s) for s in TARGETS[family]}
    out = []
    for p in PPEMS:
        if p in held:
            continue
        sizes = [k / 2 for k in range(2, 145) if ppem(k / 2) == p]
        out.append(min(sizes, key=lambda s: (s != int(s), abs(s - p * 0.75))))
    checks = [s for s in CHECK_CANDIDATES if s not in TARGETS[family] and ppem(s) in held]
    return tuple(sorted(out)) + tuple(checks[:3])


PPEM_TARGETS = {family: _ppem_sizes(family) for family in TARGETS}

# group rows_more: the families xlformat measures widths for on Windows beyond the five of
# rows_wide, and MS Sans Serif and Courier, which the machine may draw as other faces; a size for
# each ppem from 1 to 96 (whole sizes first); "arial" in lower case, regular only, to test that
# Excel matches a face's name ignoring case
GROUP_MORE = "rows_more"
LAYOUT_MORE = ROOT / "tests" / "captures" / "windows" / "row_height_more.layout.tsv.gz"
MORE_FAMILIES = (
    "Aptos", "Arial Black", "Arial Narrow", "Book Antiqua", "Bookman Old Style", "Calibri Light",
    "Cambria", "Candara", "Century Gothic", "Comic Sans MS", "Consolas", "Garamond", "Georgia",
    "Haettenschweiler", "Impact", "Lucida Console", "Rockwell", "Tahoma", "Trebuchet MS",
    "Verdana", "MS Sans Serif", "Courier", "arial",
)  # fmt: skip
MORE_STYLES = {"arial": ((False, False),)}


def _every_ppem() -> tuple[float, ...]:
    """A size for each ppem from 1 to 96: a whole size where one draws at it, else the nearest."""
    out = []
    for p in PPEMS:
        sizes = [k / 2 for k in range(2, 145) if ppem(k / 2) == p]
        out.append(min(sizes, key=lambda s: (s != int(s), abs(s - p * 0.75))))
    return tuple(out)


MORE_TARGETS = dict.fromkeys(MORE_FAMILIES, _every_ppem())


def wrapped(group: str, size: float) -> tuple[int, ...]:
    """The line counts of a face's wrapped rows: four in a row, the first above the Normal row.

    rows_wide: 1 to 4. rows_ppem: from enough lines that the face's text rises past the Calibri
    11 Normal font's (20 pixels) with room, a line taken as at least the ppem's pixels high.
    """
    if group == GROUP:
        return WRAPPED
    first = max(1, math.ceil((NORMAL_HD[0] + 8) / ppem(size)))
    return tuple(range(first, first + 4))


def style_name(bold: bool, italic: bool) -> str:
    return " ".join(n for n, on in (("bold", bold), ("italic", italic)) if on) or "regular"


def normal_lines(size: float) -> int:
    """Normal-font lines whose top passes a face of ``size``'s one-line top, with room."""
    return max(2, math.ceil((size * 4 / 3 * 1.6 + 10) / NORMAL_HD[0]) + 1)


def workbooks(group: str = GROUP) -> list[tuple[str, str, tuple[RowCase, ...]]]:
    """(file name, family and style label, rows) per workbook of ``group``, in job order."""
    out = []
    targets, prefix = {
        GROUP: (TARGETS, "rw"),
        GROUP_PPEM: (PPEM_TARGETS, "rp"),
        GROUP_MORE: (MORE_TARGETS, "rm"),
    }[group]
    for family, sizes in targets.items():
        for bold, italic in MORE_STYLES.get(family, STYLES):
            style = style_name(bold, italic)
            cases = [RowCase("number", (Cell(1, 1234),))]
            for size in sizes:
                f = CellFont(family, size, bold, italic)
                cases += [
                    RowCase(f"{f.label}|{n}", (Cell(1, _lines(n), f, True),))
                    for n in wrapped(group, size)
                ]
                m = normal_lines(size)
                cases.append(
                    RowCase(
                        f"{f.label}|beside {m}", (Cell(1, _lines(m), None, True), Cell(2, 1, f))
                    )
                )
            lower = "_lower" if family != family.title() and family.lower() == family else ""
            name = f"{prefix}_{family}{lower}_{style}".replace(" ", "_").lower() + ".xlsx"
            out.append((name, f"{family}|{style}", tuple(cases)))
    return out


def layout_path(group: str) -> Path:
    return {GROUP: LAYOUT, GROUP_PPEM: LAYOUT_PPEM, GROUP_MORE: LAYOUT_MORE}[group]


def read_layout(group: str = GROUP) -> dict[tuple[str, str], int]:
    """(workbook, row tag) -> height in pixels."""
    with gzip.open(layout_path(group), "rt", encoding="utf-8", newline="") as f:
        reader = csv.DictReader((x for x in f if not x.startswith("#")), dialect="excel-tab")
        out = {}
        for r in reader:
            px = float(r["height"]) * 4 / 3
            if px != round(px):
                msg = f"{r['workbook']} {r['tag']}: {r['height']} points is not a pixel"
                raise ValueError(msg)
            out[(r["workbook"], r["tag"])] = round(px)
    return out


Face = tuple[str, str, float]  # family, style, size
MAX_ROW_PX = 546  # Excel's largest row, 409.5 points: a taller fitted row stops there


def fit(group: str = GROUP) -> tuple[dict[Face, tuple[int, int]], list[str]]:
    """(H, D) per face, and every row the model does not give."""
    got = read_layout(group)
    hn, dn = NORMAL_HD
    table: dict[Face, tuple[int, int]] = {}
    problems: list[str] = []
    for name, label, cases in workbooks(group):
        family, style = label.split("|")
        if got[(name, "number")] != hn:
            problems.append(f"{name}: the Normal row is {got[(name, 'number')]} px, not {hn}")
        for case in cases[1::5]:
            flabel = case.tag.split("|")[0]
            size = float(flabel.split(" ")[len(family.split(" "))])
            lines = tuple(
                int(c.tag.split("|")[1])
                for c in cases
                if c.tag.startswith(flabel + "|")
                and not c.tag.endswith(tuple(f"beside {k}" for k in range(1, 99)))
            )
            w = [got[(name, f"{flabel}|{n}")] for n in lines]
            m = normal_lines(size)
            beside = got[(name, f"{flabel}|beside {m}")]
            h = w[2] - w[1]
            top = beside - (m * hn - dn)  # max(Dn, D)
            d = lines[1] * h + top - w[1]
            table[(family, style, size)] = (h, d)

            def row(cells: list[tuple[int, int, int]]) -> int:
                fitted = max(n * hh - dd for n, hh, dd in cells) + max(dd for _, _, dd in cells)
                return min(fitted, MAX_ROW_PX)

            for n, width in zip(lines, w, strict=True):
                want = row([(1, hn, dn), (n, h, d)])
                if width != want:
                    problems.append(f"{flabel} {n} lines: {width} px, model {want}")
            want = row([(m, hn, dn), (1, h, d)])
            if beside != want or not 0 <= d <= h:
                problems.append(f"{flabel} beside: {beside} px, model {want} (H {h}, D {d})")
    return table, problems


def main(argv: list[str]) -> None:
    table, problems = fit()
    for line in problems:
        print(line)
    print(f"{len(table)} faces, {len(problems)} problems")
    if "--check" in argv or problems:
        return
    from capture import windows_rows

    windows_rows.main([])


if __name__ == "__main__":
    main(sys.argv[1:])
