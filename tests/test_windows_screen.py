"""Excel for Windows' rules, placement and ink against every captured row (96 dpi).

Rules: every cell of the rules captures (CAPTURING.md section 2 items 73, 75 and 77 to 80)
against ``rule_metrics(..., platform="windows")`` in device pixels from the baseline (the edge
below the H stems' last ink row). Placement: every row of ``screen_vertical`` and
``screen_wrapped`` (items 74 and 76) against ``vertical_metrics``. Ink: the faces' rows in every
workbook of ``screen_vertical`` against ``ink_metrics``.
"""

from __future__ import annotations

import math
from collections.abc import Iterator
from pathlib import Path

import pytest
from capture.windows_screen_tables import FIELDS, cell_values, render

from xlformat import Font, ink_metrics, rule_metrics, vertical_metrics
from xlformat._screen_windows import WINDOWS_RULES

WINDOWS = Path(__file__).parent / "captures" / "windows"
ONE_RULE = ("screen_pilot_rules", "screen_rules", "screen_rules_zoom", "screen_rules_fresh")
TABLES = ("screen_rules_table", "screen_rules_table_arial", "screen_rules_table_high")
STYLES = ("bold italic", "bold", "italic")


def _rows(name: str) -> list[list[str]]:
    lines = (WINDOWS / f"{name}.layout.tsv").read_text(encoding="utf-8").splitlines()
    return [x.split("\t") for x in lines if x and not x.startswith("#")][1:]


def _cells() -> Iterator[tuple[str, list[str]]]:
    """Every rule cell as the table's seven columns: an earlier one-rule row's double takes the
    baseline of its group's single underline row at the same zoom, size and row height."""
    for name in ONE_RULE:
        rows = _rows(name)
        single = {
            (r[0], r[1].rsplit(" ", 1)[0], r[2]): r[4] for r in rows if r[1].endswith(" single")
        }
        for z, tag, rh, ppp, base, measured in rows:
            if tag.endswith(" none") or tag.startswith("filler"):
                continue
            row_base = single.get((z, tag.rsplit(" ", 1)[0], rh), "")
            yield name, [z, tag, rh, ppp, base, row_base, measured]
    for name in TABLES:
        for cell in _rows(name):
            yield name, cell


def _cut(cell: list[str]) -> bool:
    """Whether a line of the cell ends on its row's last pixel row (cut by the row)."""
    last = math.floor(float(cell[2]) * float(cell[3]) + 0.5) - 1
    return any(int(r.split("-")[1]) >= last for r in cell[6].split())


def test_the_generated_table_is_the_scripts() -> None:
    assert render() == (Path(__file__).parents[1] / "src/xlformat/_screen_windows.py").read_text()


def test_every_rule_cell_matches_rule_metrics() -> None:
    misses = []
    checked = 0
    for name, cell in _cells():
        z, tag = int(cell[0]), cell[1]
        face, size, _rule = tag.rsplit(" ", 2)
        got = rule_metrics((face, float(size)), zoom=z, platform="windows")
        for field, value in cell_values(cell).items():
            if field == "double_accounting_top":
                field = "accounting_top"
            expected = getattr(got, field) * z / 75
            checked += 1
            if abs(expected - value) > 1e-9:
                misses.append((name, z, tag, field, value, expected, _cut(cell)))
    assert checked > 14000
    # the known deviations: lines cut by their row's last pixel row (small rows at zoom 25 and
    # the earlier groups' rows), and the cells measured where another zoom's taller row disagrees
    uncut = [m for m in misses if not m[6]]
    assert len(uncut) <= len(UNCUT_MISSES), uncut[:10]
    assert {(m[0], m[1], m[2], m[3]) for m in uncut} <= UNCUT_MISSES, uncut[:10]


# Cells whose measured line differs from the table while no line of the cell reaches its row's
# last pixel row. Aptos Narrow's double accounting at zoom 25 below ppem 10, in rows of 5 or 6
# pixels: its own lines lie past the row (CAPTURING.md section 2 item 80), and what the reader
# finds there, on the row's first pixel rows, are the row above's lines, drawn on past that row
# (zoom 25 draws no gridline over them); the same ppems' cells at zoom 75 give the table its
# values.
UNCUT_MISSES = {
    (
        "screen_rules_table_high",
        25,
        f"Aptos Narrow {size} doubleAccounting",
        "accounting_top",
    )
    for size in ("4.5", "7.5", "10.5", "13.5", "25.5")
}


@pytest.mark.parametrize("face", ["Calibri", "Arial", "Aptos Narrow"])
def test_rule_metrics_answers_every_whole_zoom(face: str) -> None:
    for zoom in (10, 37, 100, 255, 400):
        for size in (4, 8.5, 11, 72):
            got = rule_metrics((face, size), zoom=zoom, platform="windows")
            p = math.floor(size * 4 / 3 * zoom / 100 + 0.5)
            assert got.double_gap == 2 * got.double_thickness
            if face != "Aptos Narrow":
                assert (face, p) in WINDOWS_RULES
    with pytest.raises(LookupError, match="under one pixel an em"):
        rule_metrics((face, 1), zoom=10, platform="windows")
    for zoom in (9, 401, 100.5):
        with pytest.raises(LookupError, match="whole zooms"):
            rule_metrics((face, 11), zoom=zoom, platform="windows")


def test_rule_metrics_on_windows_knows_three_faces() -> None:
    with pytest.raises(LookupError, match="Excel for Windows"):
        rule_metrics(("Georgia", 11), platform="windows")


def _face(label: str) -> Font:
    style = next((s for s in STYLES if label.endswith(" " + s)), "")
    name, size = label.removesuffix(" " + style).rsplit(" ", 1) if style else label.rsplit(" ", 1)
    return Font(name, float(size), "bold" in style, "italic" in style)


def _placed() -> Iterator[tuple[str, list[str]]]:
    for name in ("screen_vertical", "screen_wrapped"):
        for row in _rows(name):
            yield name, row


def test_every_placed_row_matches_vertical_metrics() -> None:
    checked = 0
    for name, (normal, tag, _stated, rh, _band, shift, h_lines, _g) in _placed():
        kind, rest = tag.split(" ", 1)
        label = rest if kind == "natural" else rest.split(" ", 1)[1]
        font, normal_font = _face(label), _face(normal)
        r = math.floor(float(rh) * 4 / 3 + 0.5)  # the row in pixels
        edges = [int(x.split("-")[1]) + 1 - int(shift) for x in h_lines.split()]
        own = vertical_metrics(font, platform="windows")
        placed = vertical_metrics(font, normal_font=normal_font, platform="windows")
        px = 4 / 3
        if kind == "top":
            expected = [round(placed.baseline_from_top * px)]
        elif kind == "center":
            h = round(own.font_height * px)
            expected = [(r - h) // 2 + round(placed.baseline_from_top * px)]
        elif kind in ("bottom", "natural"):
            expected = [r - round(placed.baseline_from_bottom * px)]
        else:  # wrapped: lines a pitch apart, the last raised by the inset's difference
            n = int(rest.split(" ", 1)[0])
            raise_px = max(
                0,
                round(
                    (vertical_metrics(normal_font, platform="windows").line_inset - own.line_inset)
                    * px
                ),
            )
            last = r - round(placed.baseline_from_bottom * px) - raise_px
            pitch = round(own.line_pitch * px)
            expected = [last - pitch * (n - 1 - k) for k in range(n)]
        assert edges == expected, (name, normal, tag)
        checked += 1
    assert checked > 700


def test_every_face_row_matches_ink_metrics() -> None:
    checked = 0
    for normal, tag, _stated, _rh, _band, _shift, h_lines, g in _rows("screen_vertical"):
        if not tag.startswith("natural "):
            continue
        font = _face(tag.removeprefix("natural "))
        cap_top, last = (int(v) for v in h_lines.split()[0].split("-"))
        ink = ink_metrics(font, platform="windows")
        assert (ink.cap_height * 4 / 3, ink.descender * 4 / 3) == pytest.approx(
            (last + 1 - cap_top, int(g.split()[0]) - last)
        ), (normal, tag)
        checked += 1
    assert checked > 90
    with pytest.raises(LookupError, match="no ink measured"):
        ink_metrics(("Calibri", 13), platform="windows")


def test_the_tables_fields_are_rule_metrics_fields() -> None:
    from xlformat import RuleMetrics

    names = set(RuleMetrics.__dataclass_fields__)
    assert set(FIELDS) == names
