"""The screen capture jobs for Excel for Windows (``capture/screen_windows.py``).

No job of these groups has run yet: the tests check that each group's job is written the same
way twice, that its planned views show every row whole, and that the reader measures a result
drawn here as a job would return it (pictures at 4/3 device pixels a point, 96 dpi at zoom 100).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytest.importorskip("openpyxl")

from capture.cases import CellFont
from capture.job import write
from capture.screen_rules import Case as RuleCase
from capture.screen_vertical import write_png
from capture.screen_windows import (
    SCREEN_GROUPS,
    VIEW_BUDGET_PX,
    Book,
    books,
    estimated_points,
    measured_rows,
    plan_views,
    views,
)
from capture.vertical_pdf import Case as VerticalCase

WHITE, MAGENTA, CYAN, BLACK = (255, 255, 255), (255, 0, 255), (0, 255, 255), (0, 0, 0)
BORDER = (171, 171, 171)
Pixels = list[list[tuple[int, int, int]]]


@pytest.mark.parametrize("group", SCREEN_GROUPS)
def test_a_group_is_written_the_same_way_twice(group: str, tmp_path: Path) -> None:
    first = write(group, f"xlformat-test-{group}", tmp_path / "a")
    second = write(group, f"xlformat-test-{group}", tmp_path / "b")
    names = sorted(p.name for p in (first / "files").iterdir())
    assert names == sorted(p.name for p in (second / "files").iterdir())
    assert (first / "job.json").read_bytes() == (second / "job.json").read_bytes()
    for name in names:
        assert (first / "files" / name).read_bytes() == (second / "files" / name).read_bytes()
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert doc["purpose"].startswith(f"xlformat capture group {group}: window pictures of ")
    assert ";" not in doc["purpose"]  # no note: the purpose cites nothing
    assert doc["audit"] == (group == "screen_pilot")
    for book, workbook in zip(books(group), doc["workbooks"], strict=True):
        shots = [s for s in workbook["steps"] if s["op"] == "capture_window"]
        assert len(shots) == sum(len(views(book, zoom)) for zoom in book.zooms)
        assert {s["select"] for s in workbook["steps"] if s["op"] == "view"} == {"AZ1"}


@pytest.mark.parametrize("group", SCREEN_GROUPS)
def test_the_planned_views_show_every_row_whole(group: str) -> None:
    for book in books(group):
        for zoom in book.zooms:
            scale = 96 / 72 * zoom / 100  # as views() computes it
            px = [h * scale for h in estimated_points(book)]
            tops = views(book, zoom)
            assert tops[0] == 1
            assert tops == sorted(set(tops))
            seen: set[int] = set()
            for k, top in enumerate(tops):
                used, row = 0.0, top
                while row <= len(px) and (row == top or used + px[row - 1] <= VIEW_BUDGET_PX):
                    used += px[row - 1]
                    seen.add(row)
                    row += 1
                if k + 1 < len(tops):  # the next view begins at this one's last row
                    assert tops[k + 1] == max(row - 1, top + 1)
            assert seen == set(range(1, len(px) + 1)), (book.stem, zoom)


@pytest.mark.parametrize("group", SCREEN_GROUPS)
def test_every_row_is_whole_below_a_views_first_row(group: str) -> None:
    """A view's first row is cut by the column headings' border: every row is planned whole as
    another row of some view, but row 1, which is a filler in every group after the pilot."""
    for book in books(group):
        assert book.lead_filler == (group != "screen_pilot")
        if book.lead_filler:
            assert book.cases[0].tag.startswith("filler")
        for zoom in book.zooms:
            scale = 96 / 72 * zoom / 100  # as views() computes it
            px = [h * scale for h in estimated_points(book)]
            whole: set[int] = set()
            for top in views(book, zoom):
                used = px[top - 1]
                for row in range(top + 1, len(px) + 1):
                    used += px[row - 1]
                    if used > VIEW_BUDGET_PX:
                        break
                    whole.add(row)
            assert set(range(2, len(px) + 1)) <= whole, (book.stem, zoom)


def test_plan_views_overlaps_by_one_row() -> None:
    assert plan_views([300.0] * 10, budget=1000) == [1, 3, 5, 7, 9]
    assert plan_views([1200.0, 100.0], budget=1000) == [1, 2]


def _draw(
    heights_px: list[int], top: int, draw_row: Any, offset: int = 40, height: int = 1100
) -> Pixels:
    """A picture of the sheet from row ``top``: column A's bands, ``draw_row`` drawing column B."""
    pixels: Pixels = [[WHITE] * 400 for _ in range(offset)]
    row = top
    while row <= len(heights_px) and len(pixels) < height:
        fill = MAGENTA if row % 2 else CYAN
        band = [[fill] * 12 + [WHITE] * 388 for _ in range(heights_px[row - 1])]
        draw_row(row, band)
        if row == top:  # the column headings' border covers the first row's top pixel
            band[0] = [BORDER] * 400
        pixels += band
        row += 1
    return pixels[:height] + [[WHITE] * 400 for _ in range(max(0, height - len(pixels)))]


def _result(book: Book, heights_px: list[int], draw_row: Any, folder: Path) -> dict[str, Any]:
    """The step results a job of ``book`` returns, its pictures drawn into ``folder``."""
    steps: dict[str, Any] = {"heights": {"row_height": [[h * 0.75] for h in heights_px]}}
    for zoom in book.zooms:
        for top in views(book, zoom):
            name = f"{book.stem}_z{zoom}_r{top:03d}.png"
            write_png(folder / name, _draw(heights_px, top, draw_row))
            steps[f"view_z{zoom}_r{top}"] = {
                "zoom": float(zoom),
                "visible_range": f"$A${top}:$AQ$99",
            }
            steps[f"shot_z{zoom}_r{top}"] = {"artifact": name, "dpi": 96}
    return steps


def test_the_rules_reader_on_a_drawn_result(tmp_path: Path) -> None:
    rules = ("none", "single", "double", "strike")
    cases = [RuleCase("rule", "filler top none")]  # row 1: under the headings' border
    cases += [RuleCase("rule", f"Calibri 11 {r}", "Calibri", 11.0, r) for r in rules]
    cases += [RuleCase("rule", f"filler {k} none") for k in range(20)]
    book = Book("rules_test", "rules", tuple(cases), (100,))
    heights_px = [24] * len(cases)
    base = 16  # the H stems' last pixel row, from the row's top

    def draw_row(row: int, band: Pixels) -> None:
        if not 2 <= row <= len(rules) + 1:
            return
        for x0 in (30, 44, 58, 72):  # four H glyphs: two stems and a bar
            for y in range(6, base + 1):
                band[y][x0] = band[y][x0 + 6] = BLACK
            band[11][x0 : x0 + 7] = [BLACK] * 7
        lines = {"single": [18], "double": [17, 19], "strike": [12], "none": []}[rules[row - 2]]
        for y in lines:
            band[y][30:79] = [BLACK] * 49

    steps = _result(book, heights_px, draw_row, tmp_path)
    got = measured_rows(book, steps, tmp_path)
    assert got == [
        "100\tCalibri 11 none\t18\t1.33333\t16\t",
        "100\tCalibri 11 single\t18\t1.33333\t16\t18-18",
        "100\tCalibri 11 double\t18\t1.33333\t16\t17-17 19-19",
        "100\tCalibri 11 strike\t18\t1.33333\t16\t12-12",
    ]


def test_the_vertical_reader_on_a_drawn_result(tmp_path: Path) -> None:
    font = CellFont("Calibri", 11)
    cases = [
        VerticalCase("filler top", font),
        VerticalCase("natural Calibri 11", font),
        VerticalCase("top H+13 Calibri 11", font, "top", 24.75),
        VerticalCase("bottom H+13 Calibri 11", font, "bottom", 24.75),
    ]
    book = Book("vertical_test", "vertical", tuple(cases), (100,))
    heights_px = [20, 20, 33, 33]
    # per row: the H's cap top and baseline, and the g's bottom, in pixels from the row's top
    ink = {1: (5, 15, 18), 2: (5, 15, 18), 3: (3, 13, 16), 4: (18, 28, 31)}

    def draw_row(row: int, band: Pixels) -> None:
        cap, base, tail = ink[row]
        for y in range(cap, base + 1):
            band[y][30] = band[y][36] = BLACK  # H
        band[base - 4][44:50] = [BLACK] * 6  # x
        for y in range(base - 4, tail + 1):
            band[y][56] = BLACK  # g, with its tail
        band[(cap + base) // 2][30:37] = [BLACK] * 7

    steps = _result(book, heights_px, draw_row, tmp_path)
    got = measured_rows(book, steps, tmp_path)
    assert got == [
        "Calibri 11\tnatural Calibri 11\t\t15\t20\t0\t5-15\t18",
        "Calibri 11\ttop H+13 Calibri 11\t24.75\t24.75\t33\t0\t3-13\t16",
        "Calibri 11\tbottom H+13 Calibri 11\t24.75\t24.75\t33\t0\t18-28\t31",
    ]


def test_a_frame_that_is_not_one_offset_stops_the_reader(tmp_path: Path) -> None:
    cases = [RuleCase("rule", "filler top none")]
    cases += [RuleCase("rule", f"filler {k} none") for k in range(6)]
    book = Book("frame_test", "rules", tuple(cases), (100,))
    drawn = [24, 24, 24, 25, 24, 24, 24]  # row 4 drawn a pixel taller than Excel reports it
    steps = _result(book, drawn, lambda row, band: None, tmp_path)
    steps["heights"] = {"row_height": [[24 * 0.75]] * len(cases)}
    with pytest.raises(SystemExit, match="not at one offset"):
        measured_rows(book, steps, tmp_path)


def test_the_table_reader_takes_the_rows_baseline_from_its_single_underline(
    tmp_path: Path,
) -> None:
    """Five rule cells side by side; the double's first line lies on the stems' last pixel row,
    so its own cell reads the baseline a row high, and the row's single underline gives it."""
    from capture.screen_table import TABLE_RULES, TableRow

    cases = [TableRow("filler top")]
    cases.append(TableRow("Calibri 11 all", "Calibri", 11.0, TABLE_RULES))
    cases += [TableRow(f"filler {k}") for k in range(20)]
    book = Book("table_test", "table", tuple(cases), (100,))
    heights_px = [30] * len(cases)
    gridlines = (12, 72, 132, 192, 252, 312)  # columns A to F's right edges: 12 and 60 pixels
    base = 20
    lines = {
        "single": [22],
        "double": [20, 22],  # the first on the stems' last row
        "strike": [14],
        "singleAccounting": [24],
        "doubleAccounting": [24, 26],
    }

    def draw_row(row: int, band: Pixels) -> None:
        for y in range(len(band)):
            for x in gridlines:
                band[y][x] = (218, 220, 224)
        if row != 2:
            return
        for c, rule in enumerate(TABLE_RULES):
            left = gridlines[c] + 6
            for x0 in (left, left + 12):  # two H glyphs: two stems and a bar
                for y in range(8, base + 1):
                    band[y][x0] = band[y][x0 + 6] = BLACK
                band[14][x0 : x0 + 7] = [BLACK] * 7
            for y in lines[rule]:
                band[y][left : left + 19] = [BLACK] * 19

    steps = _result(book, heights_px, draw_row, tmp_path)
    steps["columns"] = {"left": [[0, 9, 54, 99, 144, 189]], "width": [[9, 45, 45, 45, 45, 45]]}
    got = measured_rows(book, steps, tmp_path)
    assert got == [
        "100\tCalibri 11 single\t22.5\t1.33333\t20\t20\t22-22",
        "100\tCalibri 11 double\t22.5\t1.33333\t19\t20\t20-20 22-22",
        "100\tCalibri 11 strike\t22.5\t1.33333\t20\t20\t14-14",
        "100\tCalibri 11 singleAccounting\t22.5\t1.33333\t20\t20\t24-24",
        "100\tCalibri 11 doubleAccounting\t22.5\t1.33333\t20\t20\t24-24 26-26",
    ]


def test_the_table_draws_every_ppem_to_192_at_two_zooms() -> None:
    from capture.screen_table import FIRST, MAX_PPEM, columns, cover, ppem, wanted

    chosen = cover()
    for zoom, sizes in FIRST.items():
        assert chosen[zoom] == tuple(sizes)
    zooms_of: dict[int, set[int]] = {}
    for zoom, sizes in chosen.items():
        for size in sizes:
            zooms_of.setdefault(ppem(size, zoom), set()).add(zoom)
    for p in range(1, 382):
        assert len(zooms_of.get(p, ())) >= wanted(p), p
    assert all(len(zooms_of[p]) >= 2 for p in range(1, 193))
    assert all(len(zooms_of[p]) >= 2 for p in range(200, 382, 8))
    assert zooms_of[MAX_PPEM] == {400}  # 72 points at zoom 400, the one size and zoom
    for book in books("screen_rules_table"):
        assert len(book.cases) < 1300  # Excel fits only about the first 1,400 rows it opens
        (zoom,) = book.zooms
        assert columns(zoom, chosen[zoom]) >= 3 or zoom < 300


def test_a_table_workbook_stays_under_the_points_seen_to_fail() -> None:
    """Under 900 fonts (cells past the 1,000th were drawn in the default font, CAPTURING.md item
    78), under 1,400 rows (item 58) and well under 9,300 points of rows (where item 78's cut fell,
    not separated from the fonts there). screen_rules_table ran before this was known."""
    from capture.screen_table import MAX_FONTS

    for book in books("screen_rules_table_arial") + books("screen_rules_table_high"):
        fonts = {(c.face, c.size, r) for c in book.cases if hasattr(c, "rules") for r in c.rules}
        assert len(fonts) + 1 < MAX_FONTS
        assert len(book.cases) < 1400
        assert sum(estimated_points(book)) < 7000
    assert MAX_FONTS <= 900


def test_the_high_table_draws_each_left_out_ppem_once() -> None:
    from capture.screen_table import HIGH_RULES, high_wanted, ppem

    drawn: dict[tuple[str, int, str], int] = {}
    for book in books("screen_rules_table_high"):
        (zoom,) = book.zooms
        for row in book.cases:
            for rule in getattr(row, "rules", ()):
                if row.face in HIGH_RULES:
                    key = (row.face, ppem(row.size, zoom), rule)
                    drawn[key] = drawn.get(key, 0) + 1
    wanted = {p for p in range(193, 385) if high_wanted(p)}
    assert wanted == {p for p in range(193, 384) if p % 8}
    for face, rules in HIGH_RULES.items():
        for rule in rules:
            assert {p for (f, p, r) in drawn if f == face and r == rule} == wanted
    assert set(drawn.values()) == {1}


def test_the_high_group_draws_aptos_narrow_small_and_sampled_at_two_zooms() -> None:
    from capture.screen_table import APTOS, APTOS_MIN_ZOOM, TABLE_RULES, ppem

    zooms: dict[tuple[int, str], set[int]] = {}
    for book in books("screen_rules_table_high"):
        (zoom,) = book.zooms
        for row in book.cases:
            for rule in getattr(row, "rules", ()):
                if row.face == APTOS:
                    assert zoom >= APTOS_MIN_ZOOM
                    zooms.setdefault((ppem(row.size, zoom), rule), set()).add(zoom)
    expected = [*range(1, 12), *range(200, 384, 8)]
    for rule in TABLE_RULES:
        assert {p for p, r in zooms if r == rule} == {*expected, 384}
        assert all(len(zooms[(p, rule)]) >= 2 for p in expected)
