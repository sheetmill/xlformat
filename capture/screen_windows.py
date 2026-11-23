"""Window captures on Excel for Windows as capture jobs: text placement and rules.

    uv run python -m capture.job write GROUP JOB_ID OUT_DIR       # GROUP one of SCREEN_GROUPS
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

The Mac's screen suites (``capture/screen_vertical.py``, ``capture/screen_rules.py``) for the faces
with Windows tables (Calibri, Aptos Narrow, Arial): each workbook holds a case a row in column B
and column A in magenta and cyan in turn, so each row is a band of its own on screen. A job opens
the workbook, reads every row's height in points, and per zoom shows the sheet from a planned list
of first rows (``view``, with the active cell far to the right, off the picture) and captures
Excel's window (``capture_window``). The views are planned from each row's estimated height in
pixels so that every row is whole in at least one of them, and each view begins at the previous
view's last row, which the sheet tabs may cover. ``read`` measures the pictures with the Mac's
measuring code at the picture's device pixels a sheet point (``read_views``) and writes
``tests/captures/windows/<group>[_<kind>].layout.tsv`` under the job's result header; the pictures
stay with the result.

Groups:

- ``screen_pilot``: Calibri and Arial at 9, 11 and 20 points under no rule, each underline and a
  strikethrough at zoom 75, 100 and 200; Calibri and Arial 11 placed top, center and bottom in
  rows of the face's height, 13 and 14 pixels taller and twice as tall, fitted and wrapped, at
  zoom 100. The job asks for the format's audit of every capture.
- ``screen_vertical``: "Hxg" in the three faces at nine sizes, placed as in the pilot, with bold,
  italic and bold italic at 11, 12 and 20 points, in a Calibri 11 workbook; and the faces at 9, 11
  and 20 points fitted and placed in rows 13 pixels taller in workbooks of five other Normal fonts.
- ``screen_rules``: the Mac's ``rules_double`` cases for the three faces (every size from 9 to 36
  and from 40 to 72 in steps of 4) at zoom 100 and 200.
- ``screen_rules_zoom``: Calibri and Arial at the nine sizes under each rule at zoom 75, 90, 100,
  110, 125 and 150, where the pixels a point are whole, a fraction, or neither.
- ``screen_rules_fresh``: the three faces at sizes from 37 to 57 points none of the groups above
  holds, under each underline and a strikethrough, at zoom 80, 120, 140 and 175.
- ``screen_wrapped``: bottom-aligned cells wrapped to two and three lines in faces whose descent
  is smaller than the Normal font's, and in two whose is not, in Calibri 20 and Arial 14
  workbooks, at zoom 100.
- ``screen_rules_table``: Calibri and Arial under each underline and a strikethrough at every
  ppem from 1 to 384, a workbook a zoom, the rules side by side in a row
  (``capture/screen_table.py``).
"""

from __future__ import annotations

import json
import math
import re
import sys
import tempfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from capture.cases import CellFont
from capture.screen_rules import DOUBLE_RULES, DOUBLE_SIZES, measure_rules
from capture.screen_rules import Case as RuleCase
from capture.screen_rules import build as build_rules
from capture.screen_table import TableRow, column_chars, cover, measure_table, table_rows
from capture.screen_table import build as build_table
from capture.screen_vertical import build as build_vertical
from capture.screen_vertical import measure_vertical
from capture.vertical_pdf import SIZES
from capture.vertical_pdf import Case as VerticalCase

FAMILIES = ("Calibri", "Aptos Narrow", "Arial")
DEFAULT_NORMAL = CellFont("Calibri", 11)  # openpyxl's workbook default
RULES = ("none", *DOUBLE_RULES)
FAR_CELL = "AZ1"  # the active cell, right of every picture at zoom 200
VIEW_BUDGET_PX = 1000  # rows planned into one view; Excel shows about 1,100 pixels of cells
DPI = 96
SCREEN_GROUPS = (
    "screen_pilot",
    "screen_vertical",
    "screen_rules",
    "screen_rules_zoom",
    "screen_wrapped",
    "screen_rules_fresh",
    "screen_rules_table",
)
_PURPOSE = {
    "screen_pilot": "window pictures of text placement and of underlines and strikethrough",
    "screen_vertical": "window pictures of text placement",
    "screen_rules": "window pictures of underlines and strikethrough",
    "screen_rules_zoom": "window pictures of underlines and strikethrough at six zooms",
    "screen_wrapped": "window pictures of wrapped text under a larger Normal font",
    "screen_rules_fresh": "window pictures of underlines and strikethrough at new sizes and zooms",
    "screen_rules_table": "window pictures of underlines and strikethrough at every ppem",
}


@dataclass(frozen=True, slots=True)
class Book:
    """One workbook of a screen job."""

    stem: str
    kind: Literal["vertical", "rules", "table"]
    cases: tuple[VerticalCase, ...] | tuple[RuleCase, ...] | tuple[TableRow, ...]
    zooms: tuple[int, ...]
    normal: CellFont = DEFAULT_NORMAL
    # row 1 is every first view's first row, whose top pixel the column headings' border covers:
    # a filler there, except in the pilot, which ran before this was seen (its row 1 is unmeasured)
    lead_filler: bool = True

    @property
    def sheet(self) -> str:
        return "vertical" if self.kind == "vertical" else "Sheet"


# --- cases --------------------------------------------------------------------------------------


def _windows_hd(font: CellFont) -> tuple[int, int]:
    from xlformat._rows_windows import WINDOWS_ROWS

    style = {
        (False, False): "regular",
        (True, False): "bold",
        (False, True): "italic",
        (True, True): "bold italic",
    }[(font.bold, font.italic)]
    return WINDOWS_ROWS[font.name.lower()][style][int(font.size)]


def _placed(
    font: CellFont, extra: tuple[int, ...] = (0, 13, 14), double: bool = True
) -> list[VerticalCase]:
    """``font`` fitted, and placed top, center and bottom in rows of its height H plus ``extra``
    pixels (13 and 14: an odd and an even slack) and of 2H, as whole pixels in points."""
    h, _ = _windows_hd(font)
    out = [VerticalCase(f"natural {font.label}", font)]
    heights = [(f"H+{e}" if e else "H", h + e) for e in extra] + ([("2H", 2 * h)] if double else [])
    for vertical in ("top", "center", "bottom"):
        for label, px in heights:
            out.append(VerticalCase(f"{vertical} {label} {font.label}", font, vertical, px * 0.75))
    return out


def vertical_cases() -> list[VerticalCase]:
    out: list[VerticalCase] = []
    for family in FAMILIES:
        for size in SIZES:
            f = CellFont(family, size)
            out += _placed(f)
            out += [VerticalCase(f"wrapped {n} {f.label}", f, "bottom", None, n) for n in (2, 3)]
    for family in FAMILIES:
        for size in (11, 12, 20):
            for bold, italic in ((True, False), (False, True), (True, True)):
                out += _placed(CellFont(family, size, bold, italic), (13,), double=False)
    return out


def mix_cases() -> list[VerticalCase]:
    """The cases of a workbook of another Normal font: what the Normal font moves."""
    out: list[VerticalCase] = []
    for family in FAMILIES:
        for size in (9, 11, 20):
            out += _placed(CellFont(family, size), (13,), double=False)
    return out


def _rule_cases(
    faces: tuple[str, ...], sizes: tuple[float, ...], rules: tuple[str, ...]
) -> list[RuleCase]:
    found = [
        RuleCase("rule", f"{face} {s:g} {rule}", face, float(s), rule)
        for face in faces
        for s in sizes
        for rule in rules
    ]
    return found + [RuleCase("rule", f"filler {k} none") for k in range(20)]


# screen_rules_fresh: sizes and zooms none of the other rules groups holds
FRESH_SIZES = (37.0, 38.0, 39.0, 41.0, 43.0, 45.0, 49.0, 53.0, 57.0)
FRESH_ZOOMS = (80, 120, 140, 175)
# screen_wrapped: wrapped cells whose face's descent is smaller than the Normal font's (and two
# controls each whose is not), bottom-aligned
WRAPPED = (
    (
        CellFont("Calibri", 20),
        (
            CellFont("Calibri", 11),
            CellFont("Arial", 11),
            CellFont("Aptos Narrow", 14),
            CellFont("Arial", 18),
            CellFont("Calibri", 24),
        ),
    ),
    (
        CellFont("Arial", 14),
        (
            CellFont("Arial", 9),
            CellFont("Calibri", 9),
            CellFont("Calibri", 14),
            CellFont("Arial", 16),
        ),
    ),
)
MIX_NORMALS = (
    CellFont("Calibri", 9),
    CellFont("Calibri", 20),
    CellFont("Arial", 9),
    CellFont("Arial", 11),
    CellFont("Aptos Narrow", 11),
)


def books(group: str) -> list[Book]:
    """The workbooks of a screen group, in the job's order, a filler first where one is taken."""
    out = []
    for book in _books(group):
        if book.lead_filler:
            filler: VerticalCase | RuleCase | TableRow = (
                VerticalCase("filler top", DEFAULT_NORMAL)
                if book.kind == "vertical"
                else TableRow("filler top")
                if book.kind == "table"
                else RuleCase("rule", "filler top none")
            )
            book = Book(
                book.stem,
                book.kind,
                (filler, *book.cases),
                book.zooms,
                book.normal,
                True,
            )
        out.append(book)
    return out


def _books(group: str) -> list[Book]:
    if group == "screen_pilot":
        pilot_vertical = [
            c
            for family in ("Calibri", "Arial")
            for c in [
                *_placed(CellFont(family, 11)),
                VerticalCase(f"wrapped 2 {family} 11", CellFont(family, 11), "bottom", None, 2),
            ]
        ]
        return [
            Book(
                "pilot_rules",
                "rules",
                tuple(_rule_cases(("Calibri", "Arial"), (9, 11, 20), RULES)),
                (75, 100, 200),
                lead_filler=False,
            ),
            Book("pilot_vertical", "vertical", tuple(pilot_vertical), (100,), lead_filler=False),
        ]
    if group == "screen_vertical":
        out = [Book("vertical_calibri_11", "vertical", tuple(vertical_cases()), (100,))]
        for normal in MIX_NORMALS:
            stem = "vertical_" + normal.label.replace(" ", "_").lower()
            out.append(Book(stem, "vertical", tuple(mix_cases()), (100,), normal))
        return out
    if group == "screen_rules":
        cases = _rule_cases(FAMILIES, tuple(float(s) for s in DOUBLE_SIZES), DOUBLE_RULES)
        return [Book("rules", "rules", tuple(cases), (100, 200))]
    if group == "screen_rules_zoom":
        cases = _rule_cases(("Calibri", "Arial"), tuple(float(s) for s in SIZES), DOUBLE_RULES)
        return [Book("rules_zoom", "rules", tuple(cases), (75, 90, 100, 110, 125, 150))]
    if group == "screen_rules_fresh":
        cases = _rule_cases(FAMILIES, FRESH_SIZES, DOUBLE_RULES)
        return [Book("rules_fresh", "rules", tuple(cases), FRESH_ZOOMS)]
    if group == "screen_rules_table":
        return [
            Book(f"table_z{zoom}", "table", tuple(table_rows(zoom, sizes)), (zoom,))
            for zoom, sizes in cover().items()
        ]
    if group == "screen_wrapped":
        return [
            Book(
                f"wrapped_{normal.label.replace(' ', '_').lower()}",
                "vertical",
                tuple(
                    VerticalCase(f"wrapped {n} {f.label}", f, "bottom", None, n)
                    for f in fonts
                    for n in (2, 3)
                ),
                (100,),
                normal,
            )
            for normal, fonts in WRAPPED
        ]
    msg = f"not a screen group: {group}"
    raise KeyError(msg)


# --- views --------------------------------------------------------------------------------------


def estimated_points(book: Book) -> list[float]:
    """Each row's height in points as planned: a stated height, else the row model's (the Mac's
    for rules rows, a fifth more, since Windows' accounting rows are not modeled)."""
    from xlformat import Font, row_metrics

    normal = Font(book.normal.name, book.normal.size, book.normal.bold, book.normal.italic)
    out: list[float] = []
    for case in book.cases:
        if isinstance(case, VerticalCase):
            if case.height is not None:
                out.append(case.height)
                continue
            f = case.font
            font = Font(f.name, f.size, f.bold, f.italic)
            wrapped = [(font, case.lines)] if case.lines > 1 else []
            cells = [] if wrapped else [font]
            out.append(
                row_metrics(normal, cell_fonts=cells, wrapped=wrapped, platform="windows").height
            )
        elif isinstance(case, TableRow):
            # the row model holds accounting rows at whole sizes from 6 only: a row is planned
            # at the next whole size, and at 6 below it, a row no shorter
            fonts = [
                Font(
                    case.face, max(6, math.ceil(case.size)), underline=None if r == "strike" else r
                )
                for r in case.rules
            ]
            out.append(1.2 * row_metrics(normal, cell_fonts=fonts).height)
        else:
            underline = None if case.rule in ("none", "strike") else case.rule
            font = Font(case.face, case.size, underline=underline)
            out.append(1.2 * row_metrics(normal, cell_fonts=[font]).height)
    return out


def plan_views(heights_px: list[float], budget: float = VIEW_BUDGET_PX) -> list[int]:
    """First rows of views showing every row whole in at least one: each view holds the rows
    that fit ``budget`` pixels, and the next begins at its last row."""
    tops: list[int] = []
    top, n = 1, len(heights_px)
    while True:
        tops.append(top)
        used, last = 0.0, top
        for row in range(top, n + 1):
            if row > top and used + heights_px[row - 1] > budget:
                break
            used += heights_px[row - 1]
            last = row
        if last >= n:
            return tops
        top = max(last, top + 1)


def views(book: Book, zoom: int) -> list[int]:
    scale = DPI / 72 * zoom / 100
    return plan_views([h * scale for h in estimated_points(book)])


# --- the job ------------------------------------------------------------------------------------


def _steps(book: Book) -> list[dict[str, Any]]:
    n = len(book.cases)
    steps: list[dict[str, Any]] = [
        {"op": "open"},
        {"op": "read", "id": "heights", "range": f"{book.sheet}!B1:B{n}", "props": ["row_height"]},
    ]
    if book.kind == "table":  # where each rule's column lies
        steps.append({"op": "read", "id": "columns", "range": f"{book.sheet}!A1:F1",
                      "props": ["left", "width"]})  # fmt: skip
    for zoom in book.zooms:
        for top in views(book, zoom):
            steps.append({
                "op": "view", "id": f"view_z{zoom}_r{top}", "sheet": book.sheet, "zoom": zoom,
                "select": FAR_CELL, "scroll_row": top, "scroll_column": 1,
            })  # fmt: skip
            steps.append({
                "op": "capture_window", "id": f"shot_z{zoom}_r{top}",
                "name": f"{book.stem}_z{zoom}_r{top:03d}.png",
            })  # fmt: skip
    steps.append({"op": "close", "save": False})
    return steps


def workbooks(group: str, files: Path) -> list[dict[str, Any]]:
    """Each workbook of the group built into ``files`` (fixed zip times) with its steps."""
    from capture.job import _fixed_copy

    out = []
    for book in books(group):
        name = f"{book.stem}.xlsx"
        with tempfile.TemporaryDirectory() as tmp:
            raw = Path(tmp) / name
            if book.kind == "vertical":
                build_vertical(book.normal, list(book.cases), raw)  # type: ignore[arg-type]
            elif book.kind == "table":
                sizes = [c.size for c in book.cases if isinstance(c, TableRow) and c.rules]
                build_table(list(book.cases), column_chars(sizes), raw)  # type: ignore[arg-type]
            else:
                build_rules("rules_double", list(book.cases), raw)  # type: ignore[arg-type]
            _fixed_copy(raw, files / name)
        out.append({"file": name, "steps": _steps(book)})
    return out


def manifest(group: str, job_id: str, books_: list[dict[str, Any]], note: str) -> dict[str, Any]:
    from capture.job import manifest_for

    doc = manifest_for(group, job_id, books_, note)
    doc["purpose"] = doc["purpose"].replace(
        "displayed text, widths and autofit per suite", _PURPOSE[group]
    )
    doc["audit"] = group == "screen_pilot"
    return doc


# --- reading the result -------------------------------------------------------------------------


def _half_up(scale: float) -> Callable[[float], int]:
    """A row's height in whole device pixels at ``scale``, as Excel draws it: half up (the
    pilot: every band of every view begins where the rows above, so rounded, end)."""
    return lambda points: math.floor(points * scale + 0.5 + 1e-9)


def _first_row(visible_range: str) -> int:
    match = re.match(r"\$?[A-Z]+\$?(\d+)", visible_range)
    if match is None:
        msg = f"visible range {visible_range!r}"
        raise ValueError(msg)
    return int(match[1])


def measured_rows(book: Book, steps: dict[str, Any], artifacts: Path) -> list[str]:
    """The layout lines of one workbook from its step results (by id) and its pictures."""
    heights = [float(row[0]) for row in steps["heights"]["row_height"]]
    if len(heights) != len(book.cases):
        msg = f"{book.stem}: {len(heights)} heights for {len(book.cases)} cases"
        raise ValueError(msg)
    lines: list[str] = []
    for zoom in book.zooms:
        pictures: list[tuple[int, Path]] = []
        scale = 0.0
        for top in views(book, zoom):
            seen = steps[f"view_z{zoom}_r{top}"]
            shot = steps[f"shot_z{zoom}_r{top}"]
            if float(seen["zoom"]) != zoom:
                msg = f"{book.stem}: zoom {seen['zoom']} where {zoom} was asked"
                raise ValueError(msg)
            scale = float(shot["dpi"]) / 72 * zoom / 100
            pictures.append((_first_row(seen["visible_range"]), artifacts / shot["artifact"]))
        # Excel draws each row's height rounded half up to whole pixels, the rows stacked; every
        # view's first row is cut by the headings' border, so row 1 is never measured
        frame = {"row_px": _half_up(scale), "unmeasured_ok": frozenset({1})}
        if book.kind == "vertical":
            for m in measure_vertical(
                pictures,
                heights,
                f"{book.stem} zoom {zoom}",
                scale,
                list(book.cases),  # type: ignore[arg-type]
                **frame,  # type: ignore[arg-type]
            ):
                if m.case.tag.startswith("filler"):
                    continue
                band_px = m.band[1] - m.band[0] + 1
                h = " ".join(f"{a}-{b}" for a, b in m.h_lines)
                g = " ".join(str(b) for b in m.g_bottoms)
                stated = "" if m.case.height is None else f"{m.case.height:g}"
                lines.append(
                    f"{book.normal.label}\t{m.case.tag}\t{stated}\t{m.height:g}\t{band_px}"
                    f"\t{m.shift}\t{h}\t{g}"
                )
        elif book.kind == "table":
            columns = steps["columns"]
            lines += measure_table(
                pictures, heights, zoom, scale, list(book.cases),  # type: ignore[arg-type]
                [float(v) for v in columns["left"][0]], [float(v) for v in columns["width"][0]],
                **frame,
            )  # fmt: skip
        else:
            rows = measure_rules(
                pictures, heights, zoom, "rules_double", match_scale=scale,
                pixels_per_point=scale, case_list=list(book.cases), **frame,  # type: ignore[arg-type]
            )  # fmt: skip
            lines += [line for line in rows if "\tfiller " not in line]
    return lines


COLUMNS = {
    "vertical": (
        "normal_font\ttag\tstated_height\trow_height\tband_px\tshift_px\th_cap_base_px\tg_bottom_px"
    ),
    "rules": "zoom\ttag\trow_height\tpixels_per_point\tbaseline_px\tmeasured",
    "table": "zoom\ttag\trow_height\tpixels_per_point\tbaseline_px\trow_baseline_px\tmeasured",
}
METHOD = {
    "vertical": (
        "window capture at zoom 100; device pixels from each row's top, computed from the view's "
        "first marker band (column A); H: cap top-baseline per line; g: descender bottom per line"
    ),
    "rules": (
        "window captures; device pixels from the row's top: the H stems' last ink row, each "
        "rule's first and last row"
    ),
    "table": (
        "window captures; a cell a rule, the rules side by side in a row; device pixels from the "
        "row's top: the cell's H stems' last ink row, the row's (its single underline's cell, "
        "else its strikethrough's), each line's first and last row"
    ),
}


def read(group: str, job: Path, result: Path, outcome: dict[str, Any], out_dir: Path) -> list[Path]:
    """The layout files of a screen group's result (checked by ``capture.job.read``)."""
    from capture.job import _results

    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    header = outcome["header"]
    lines: dict[str, list[str]] = {}
    for book in books(group):
        steps = _results(by_file[f"{book.stem}.xlsx"])
        lines.setdefault(book.kind, []).extend(measured_rows(book, steps, result / "artifacts"))
    written = []
    for kind, rows in lines.items():
        name = f"{group}.layout.tsv" if len(lines) == 1 else f"{group}_{kind}.layout.tsv"
        path = out_dir / name
        with path.open("w", encoding="utf-8", newline="\n") as f:
            f.write(f"# suite: {name.removesuffix('.layout.tsv')} (capture/screen_windows.py)\n")
            f.write("# platform: windows\n")
            f.write(f"# excel_build: {header['environment']['excel']['build']}\n")
            f.write(f"# job: {group}.job.json\n# job_id: {header['job_id']}\n")
            f.write(f"# job_header: {json.dumps(header, sort_keys=True, ensure_ascii=False)}\n")
            f.write(f"# method: {METHOD[kind]}\n")
            f.write(COLUMNS[kind] + "\n")
            f.write("\n".join(rows) + "\n")
        written.append(path)
    return written


if __name__ == "__main__":
    sys.exit(__doc__)
