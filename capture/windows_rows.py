"""Fit Excel for Windows' row heights and write ``src/xlformat/_rows_windows.py``.

    uv run python -m capture.windows_rows [--check]

Reads ``tests/captures/windows/row_height.layout.tsv.gz`` (CAPTURING.md section 2 item 69): one
workbook per Normal font (Calibri, Aptos Narrow, Arial; four styles; nine sizes), each row's
height in points at 96 dpi, which are whole pixels (``points * 96 / 72``). Model, in pixels: a
face (a family, style and size) has a height ``H`` and a descent ``D``; a row holding cells of
``lines`` lines in faces ``f`` (the Normal font counting as a one-line cell) is
``max(lines * H(f) - D(f)) + max D(f)``. ``H`` is each face's own row, the first row of the
workbook it is the Normal font of; ``D`` is searched face by face, each whole value from 0 to
``H``, keeping the one with the fewest misses over every row until no face changes.
``--check`` reports and writes nothing.
"""

from __future__ import annotations

import csv
import gzip
import sys
from collections import Counter
from pathlib import Path

from capture.cases import CellFont
from capture.row_heights import fonts, rows

ROOT = Path(__file__).resolve().parents[1]
LAYOUT = ROOT / "tests" / "captures" / "windows" / "row_height.layout.tsv.gz"
OUT = ROOT / "src" / "xlformat" / "_rows_windows.py"
SCALE = 96 / 72
Face = tuple[str, str, int]  # family, style, size


def _face(font: CellFont) -> Face:
    style = " ".join(n for n, on in (("bold", font.bold), ("italic", font.italic)) if on)
    return font.name, style or "regular", int(font.size)


def _captured() -> dict[tuple[str, int], float]:
    """(Normal font label, row) -> height in pixels."""
    with gzip.open(LAYOUT, "rt", encoding="utf-8", newline="") as f:
        reader = csv.DictReader((x for x in f if not x.startswith("#")), dialect="excel-tab")
        out = {}
        for r in reader:
            px = float(r["height"]) * SCALE
            if px != round(px):
                msg = f"{r['normal_font']} row {r['row']}: {r['height']} points is not a pixel"
                raise ValueError(msg)
            out[(r["normal_font"], int(r["row"]))] = px
    return out


def _cells(normal: CellFont) -> list[tuple[int, list[tuple[Face, int]]]]:
    """Per row: its number and its cells as (face, lines), the Normal font first."""
    out = []
    for k, case in enumerate(rows(), 1):
        cells = [(_face(normal), 1)]
        for cell in case.cells:
            lines = cell.value.count("\n") + 1 if isinstance(cell.value, str) and cell.wrap else 1
            cells.append((_face(cell.font) if cell.font else _face(normal), lines))
        out.append((k, cells))
    return out


def height(cells: list[tuple[Face, int]], table: dict[Face, tuple[int, int]]) -> int:
    top = max(lines * table[f][0] - table[f][1] for f, lines in cells)
    return top + max(table[f][1] for f, _ in cells)


def fit() -> tuple[dict[Face, tuple[int, int]], int, int, list[str]]:
    got = _captured()
    normals = fonts()
    cases = [(n, k, cells) for n in normals for k, cells in _cells(n)]
    heights: dict[Face, set[int]] = {}
    for n in normals:
        heights.setdefault(_face(n), set()).add(int(got[(n.label, 1)]))
    # a start near the typical descent; the search below decides each D
    table = {f: (h := next(iter(hs)), round(h * 0.22)) for f, hs in heights.items()}

    def misses(t: dict[Face, tuple[int, int]]) -> int:
        return sum(height(c, t) != got[(n.label, k)] for n, k, c in cases)

    changed = True
    while changed:
        changed = False
        for face in sorted(table):
            h, d = table[face]
            best = min(range(h + 1), key=lambda x: (misses({**table, face: (h, x)}), abs(x - d)))
            if best != d:
                table[face] = (h, best)
                changed = True
    wrong = [
        f"{n.label} row {k}: {got[(n.label, k)]:g} px, model {height(c, table)}"
        for n, k, c in cases
        if height(c, table) != got[(n.label, k)]
    ]
    return table, len(cases) - len(wrong), len(cases), wrong


def _write(table: dict[Face, tuple[int, int]]) -> None:
    families = sorted(
        {f for f, _, _ in table}, key=lambda f: ("Calibri", "Aptos Narrow", "Arial").index(f)
    )
    lines = [
        '"""Row heights in Excel for Windows at 96 dpi, in pixels (generated).',
        "",
        "Written by ``capture/windows_rows.py`` from",
        "``tests/captures/windows/row_height.layout.tsv.gz``. Per family and size, the face's",
        "height H and descent D, by style: a row is ``max(lines * H - D) + max D`` over its",
        "cells, the Normal font a one-line cell among them. Do not edit by hand.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "# family (lower case) -> style -> size -> (H, D) in pixels",
        "WINDOWS_ROWS: dict[str, dict[str, dict[int, tuple[int, int]]]] = {",
    ]
    for family in families:
        lines.append(f'    "{family.lower()}": {{')
        for style in ("regular", "bold", "italic", "bold italic"):
            sizes = sorted(z for f, st, z in table if f == family and st == style)
            pairs = ", ".join(f"{z}: {table[(family, style, z)]}" for z in sizes)
            lines.append(f'        "{style}": {{{pairs}}},')
        lines.append("    },")
    lines += ["}", ""]
    OUT.write_text("\n".join(lines))
    print(f"wrote {OUT.relative_to(ROOT)}")


def main(argv: list[str]) -> None:
    table, matched, total, wrong = fit()
    for family in ("Calibri", "Aptos Narrow", "Arial"):
        for style in ("regular", "bold", "italic", "bold italic"):
            row = sorted((z, hd) for (f, st, z), hd in table.items() if f == family and st == style)
            print(family, style, " ".join(f"{z}:{h}/{d}" for z, (h, d) in row))
    print(
        f"{matched} of {total} rows; misses "
        f"{Counter(w.split(' row')[0] for w in wrong).most_common(5)}"
    )
    for w in wrong[:10]:
        print(" ", w)
    if "--check" not in argv:
        _write(table)


if __name__ == "__main__":
    main(sys.argv[1:])
