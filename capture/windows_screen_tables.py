"""Excel for Windows' rules and ink as data: ``src/xlformat/_screen_windows.py``.

    uv run python -m capture.windows_screen_tables

Reads the committed window captures (CAPTURING.md section 2 items 74 and 78 to 80) and writes:

- ``WINDOWS_RULES``: (face, ppem) -> the eight values of ``RuleMetrics`` in device pixels at
  that ppem, measured from the baseline (the edge below the H stems' last ink row), ``None``
  where the rule gives the value. Tabled: every value at ppem 1 to 11 in Calibri, Arial and
  Aptos Narrow; Calibri's and Arial's double underline (top and thickness) at every ppem from 1
  to 384; Arial's accounting underline's top at every ppem (its hinted descent). A value drawn
  at several zooms is taken from the tallest row, where no row cuts the line.
- ``WINDOWS_INK``: (face, style, size) -> (cap height, descender) in device pixels at zoom 100:
  the ink of "H" above the baseline and of "g" below it, from the cells of the Calibri 11
  workbook of ``screen_vertical``.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WINDOWS = ROOT / "tests" / "captures" / "windows"
OUT = ROOT / "src" / "xlformat" / "_screen_windows.py"
TABLES = ("screen_rules_table", "screen_rules_table_arial", "screen_rules_table_high")
FACES = ("Calibri", "Arial", "Aptos Narrow")
SMALL = 11  # every value is tabled up to this ppem
FIELDS = (
    "underline_top",
    "underline_thickness",
    "double_top",
    "double_thickness",
    "double_gap",
    "accounting_top",
    "strikethrough_top",
    "strikethrough_thickness",
)
TOP_RUN_MIN = 2  # a run starting on a row's first two pixel rows is the row above's line


def _rows(path: Path) -> list[list[str]]:
    lines = path.read_text(encoding="utf-8").splitlines()
    return [x.split("\t") for x in lines if x and not x.startswith("#")][1:]


def cell_values(cell: list[str]) -> dict[str, int]:
    """A table cell's values in device pixels from the baseline's edge (``FIELDS`` names)."""
    _zoom, tag, _rh, _ppp, own, row_base, measured = cell
    rule = tag.rsplit(" ", 1)[1]
    base = row_base if rule == "double" and row_base else own
    if not base:
        return {}
    edge = int(base) + 1
    runs = []
    for r in measured.split():
        a, _, b = r.partition("-")
        if int(a) >= TOP_RUN_MIN:
            runs.append((int(a), int(b)))
    out: dict[str, int] = {}
    if rule == "single" and len(runs) == 1:
        out["underline_top"] = runs[0][0] - edge
        out["underline_thickness"] = runs[0][1] - runs[0][0] + 1
    elif rule == "strike" and len(runs) == 1:
        out["strikethrough_top"] = edge - runs[0][0]
        out["strikethrough_thickness"] = runs[0][1] - runs[0][0] + 1
    elif rule == "singleAccounting" and len(runs) == 1:
        out["accounting_top"] = runs[0][0] - edge
    elif rule == "double" and runs:
        out["double_top"] = runs[0][0] - edge
        out["double_thickness"] = runs[0][1] - runs[0][0] + 1
        if len(runs) == 2:
            out["double_gap"] = runs[1][0] - runs[0][0]
    elif rule == "doubleAccounting" and runs:
        out["double_accounting_top"] = runs[0][0] - edge
    return out


def measured() -> dict[tuple[str, int], dict[str, int]]:
    """(face, ppem) -> field -> value, each from the tallest row that holds it (the first such
    in file order on a tie); a later file's cell replaces an earlier one's at the same zoom."""
    cells: dict[tuple[str, str], list[str]] = {}
    for name in TABLES:
        for cell in _rows(WINDOWS / f"{name}.layout.tsv"):
            cells[(cell[0], cell[1])] = cell
    best: dict[tuple[str, int, str], tuple[int, int]] = {}
    for cell in cells.values():
        zoom, tag, rh, ppp = cell[0], cell[1], float(cell[2]), float(cell[3])
        face, size, _rule = tag.rsplit(" ", 2)
        p = math.floor(float(size) * 4 / 3 * int(zoom) / 100 + 0.5)
        row_px = math.floor(rh * ppp + 0.5)
        for field, value in cell_values(cell).items():
            key = (face, p, field)
            if key not in best or row_px > best[key][0]:
                best[key] = (row_px, value)
    out: dict[tuple[str, int], dict[str, int]] = {}
    for (face, p, field), (_, value) in best.items():
        out.setdefault((face, p), {})[field] = value
    return out


def rules() -> dict[tuple[str, int], tuple[int | None, ...]]:
    found = measured()
    out = {}
    for (face, p), values in sorted(found.items()):
        if p <= SMALL:
            keep = set(FIELDS)
        elif face in ("Calibri", "Arial"):
            keep = {"double_top", "double_thickness"} | (
                {"accounting_top"} if face == "Arial" else set()
            )
        else:
            continue
        if face == "Arial" and "accounting_top" not in values and "double_accounting_top" in values:
            values = {**values, "accounting_top": values["double_accounting_top"]}
        row = tuple(values.get(f) if f in keep else None for f in FIELDS)
        if any(v is not None for v in row):
            out[(face, p)] = row
    for face in ("Calibri", "Arial"):
        missing = [p for p in range(1, 385) if out.get((face, p), (None,) * 8)[2] is None]
        if missing:
            sys.exit(f"{face}: no double underline measured at ppem {missing[:10]}")
    missing = [p for p in range(1, 385) if out.get(("Arial", p), (None,) * 8)[5] is None]
    if missing:
        sys.exit(f"Arial: no accounting underline measured at ppem {missing[:10]}")
    return out


def ink() -> dict[tuple[str, str, int], tuple[int, int]]:
    out = {}
    for row in _rows(WINDOWS / "screen_vertical.layout.tsv"):
        normal, tag, _stated, _rh, _band, _shift, h, g = row
        if normal != "Calibri 11" or not tag.startswith("natural "):
            continue
        label = tag.removeprefix("natural ")
        face = next(f for f in FACES if label.startswith(f + " "))
        size, *style = label.removeprefix(face + " ").split(" ")
        cap_top, base = (int(v) for v in h.split()[0].split("-"))
        out[(face, " ".join(style) or "regular", int(size))] = (base + 1 - cap_top, int(g) - base)
    return out


def render() -> str:
    """The module's text."""
    lines = [
        '"""Excel for Windows\' rules and ink at 96 dpi, in pixels (generated; do not edit).',
        "",
        "Written by ``capture/windows_screen_tables.py`` from ``tests/captures/windows/``:",
        "``screen_rules_table``, ``screen_rules_table_arial`` and ``screen_rules_table_high``",
        "(CAPTURING.md section 2 items 78 to 80) and ``screen_vertical`` (item 74).",
        '"""',
        "",
        "# (face, ppem) -> (underline top, underline thickness, double underline's first line, its",
        "# thickness, its gap, accounting underline top, strikethrough top above the baseline,",
        "# strikethrough thickness), in pixels below the baseline; None where the rule gives it.",
        "# Arial's accounting top at ppem 255 is drawn a pixel above its hinted descent.",
        "WINDOWS_RULES: dict[tuple[str, int], tuple[int | None, ...]] = {",
    ]
    for key, row in rules().items():
        lines.append(f'    ("{key[0]}", {key[1]}): {row!r},')
    lines += [
        "}",
        "",
        '# (face, style, size) -> ("H" above the baseline, "g" below it) at zoom 100',
        "WINDOWS_INK: dict[tuple[str, str, int], tuple[int, int]] = {",
    ]
    for key, value in sorted(ink().items()):
        lines.append(f'    ("{key[0]}", "{key[1]}", {key[2]}): {value!r},')
    lines.append("}")
    return "\n".join(lines) + "\n"


def write() -> None:
    OUT.write_text(render(), encoding="utf-8")
    print(f"wrote {OUT}")


if __name__ == "__main__":
    write()
