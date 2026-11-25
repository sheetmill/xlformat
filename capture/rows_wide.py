"""Row heights on Excel for Windows across sizes: the face's height H and descent D per size.

    uv run python -m capture.job write rows_wide JOB_ID OUT_DIR
    uv run python -m capture.job read rows_wide JOB_DIR RESULT_DIR
    uv run python -m capture.rows_wide [--check]     # fit; --check reports, writes nothing

One workbook per family and style of :data:`TARGETS` (20), its Normal font Calibri 11 (H 20, D 5
pixels at 96 dpi, CAPTURING.md section 2 item 69). Per size of the family, five rows: a cell in
the face wrapped to 1, 2, 3 and 4 lines (one short glyph a line), and a one-line cell in the face
beside a Normal-font cell wrapped to :func:`normal_lines` lines, tall enough to set the row's top.
The job reads each row's height. ``read`` writes ``tests/captures/windows/row_height_wide.layout
.tsv.gz``.

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
from capture.row_heights import Cell, RowCase, _lines

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


def style_name(bold: bool, italic: bool) -> str:
    return " ".join(n for n, on in (("bold", bold), ("italic", italic)) if on) or "regular"


def normal_lines(size: float) -> int:
    """Normal-font lines whose top passes a face of ``size``'s one-line top, with room."""
    return max(2, math.ceil((size * 4 / 3 * 1.6 + 10) / NORMAL_HD[0]) + 1)


def workbooks() -> list[tuple[str, str, tuple[RowCase, ...]]]:
    """(file name, family and style label, rows) per workbook, in job order."""
    out = []
    for family, sizes in TARGETS.items():
        for bold, italic in STYLES:
            style = style_name(bold, italic)
            cases = [RowCase("number", (Cell(1, 1234),))]
            for size in sizes:
                f = CellFont(family, size, bold, italic)
                cases += [
                    RowCase(f"{f.label}|{n}", (Cell(1, _lines(n), f, True),)) for n in WRAPPED
                ]
                m = normal_lines(size)
                cases.append(
                    RowCase(
                        f"{f.label}|beside {m}", (Cell(1, _lines(m), None, True), Cell(2, 1, f))
                    )
                )
            name = f"rw_{family}_{style}".replace(" ", "_").lower() + ".xlsx"
            out.append((name, f"{family}|{style}", tuple(cases)))
    return out


def read_layout() -> dict[tuple[str, str], int]:
    """(workbook, row tag) -> height in pixels."""
    with gzip.open(LAYOUT, "rt", encoding="utf-8", newline="") as f:
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


def fit() -> tuple[dict[Face, tuple[int, int]], list[str]]:
    """(H, D) per face, and every row the model does not give."""
    got = read_layout()
    hn, dn = NORMAL_HD
    table: dict[Face, tuple[int, int]] = {}
    problems: list[str] = []
    for name, label, cases in workbooks():
        family, style = label.split("|")
        if got[(name, "number")] != hn:
            problems.append(f"{name}: the Normal row is {got[(name, 'number')]} px, not {hn}")
        for case in cases[1::5]:
            flabel = case.tag.split("|")[0]
            size = float(flabel.split(" ")[len(family.split(" "))])
            w = [got[(name, f"{flabel}|{n}")] for n in WRAPPED]
            m = normal_lines(size)
            beside = got[(name, f"{flabel}|beside {m}")]
            h = w[2] - w[1]
            top = beside - (m * hn - dn)  # max(Dn, D)
            d = 2 * h + top - w[1]
            table[(family, style, size)] = (h, d)

            def row(cells: list[tuple[int, int, int]]) -> int:
                return max(n * hh - dd for n, hh, dd in cells) + max(dd for _, _, dd in cells)

            for n, width in zip(WRAPPED, w, strict=True):
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
