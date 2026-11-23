"""The rules table on Excel for Windows: every underline and a strikethrough by ppem, a cell each.

Group ``screen_rules_table`` of ``capture/screen_windows.py``. Calibri and Arial are drawn at
every ppem (the face's pixels per em, ``round(size * 4/3 * zoom / 100)`` half up) from 1 to 384,
the range of Excel's font sizes (1 to 72 points in half points) and zooms (10 to 400), so that
the lines Excel draws can be read by ppem alone. Every ppem up to 192 is drawn at two zooms, so
the table shows whether an equal ppem draws equal lines, and every eighth ppem above 192 at two
(``wanted``): above 192 a row is too tall for more than one or two to a picture, and the
sample shows whether those ppems follow a rule. Where two rows of a size may not fit a picture,
a plain row goes before each.

One workbook per zoom. A sheet row holds one face at one size, its five rules side by side in
columns B onward (``TABLE_RULES`` in that order), as many columns as fit Excel's window at the
zoom (``columns``); rules that do not fit continue on the next row. A double underline's first
line can lie on the H stems' last pixel row, which then reads as part of the line: the row's
single underline or strikethrough cell, drawn at the same baseline, gives the baseline instead.
Each cell holds ``TABLE_TEXT``, left- and bottom-aligned, so that the cells stay narrow enough
for several to fit at zoom 400.

Zooms 75, 100, 150 and 200 hold sizes the earlier rules groups measured with one rule a row, so
a row that holds every rule can be compared with them.

Group ``screen_rules_table_arial`` draws again the cells the zoom 150 and 200 workbooks drew in
the default font (Arial from 29 points), a workbook a zoom, each well under ``MAX_FONTS``.
"""

from __future__ import annotations

import math
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

from capture.screen_vertical import (
    Image,
    _marker_column,
    _marker_right,
    read_views,
)

TABLE_TEXT = "HH"
TABLE_FACES = ("Calibri", "Arial")
TABLE_RULES = ("single", "double", "strike", "singleAccounting", "doubleAccounting")
SIZES = tuple(k / 2 for k in range(2, 145))  # 1 to 72 points in half points
ZOOMS = range(10, 401)
MAX_PPEM = 384
# drawn first, whole: every size at zoom 150 and 200 (ppem 2 to 192), and sizes the earlier rules
# groups measured at zoom 75 and 100
FIRST = {150: SIZES, 200: SIZES, 75: (9.0, 10.0, 11.0), 100: (9.0, 11.0, 20.0, 36.0)}
# screen_rules_table_arial: the cells of screen_rules_table drawn in the default font (from the
# 1,000th font of the zoom 150 and 200 workbooks, CAPTURING.md item 78), Arial 29 on, again
# in two workbooks a zoom, 29 to 50 and 50.5 to 72 points, each about half the rows' height
ARIAL_AGAIN = tuple(
    (zoom, part, tuple(s for s in SIZES if lo <= s <= hi))
    for zoom in (150, 200)
    for part, (lo, hi) in enumerate(((29.0, 50.0), (50.5, 72.0)), 1)
)
MAX_FONTS = 900  # a workbook's fonts: cells past the 1,000th were drawn in the default font
# the advance of "H" in em, the wider face's setting the columns' width
ADVANCE = {"Calibri": 1276 / 2048, "Arial": 1479 / 2048}
WINDOW_PX = 2560  # Excel's window at 96 dpi, maximized on the runner's display
COLUMN_A_CHARS = 4
# column A narrower than this in device pixels holds no marker the reader finds (zoom 15: about 5)
MIN_MARKER_PX = 7


def ppem(size: float, zoom: int) -> int:
    return math.floor(size * 4 / 3 * zoom / 100 + 0.5)


def wanted(p: int) -> int:
    """How many zooms draw ppem ``p``: two up to 192 (zoom 200's largest), and above, where a
    row is too tall for more than one or two to a picture, two for every eighth ppem."""
    return 2 if p <= 192 or p % 8 == 0 else 0


def _text_px(size: float, zoom: int) -> float:
    """The widest face's ``TABLE_TEXT`` at ``size`` and ``zoom`` in device pixels."""
    em = size * 4 / 3 * zoom / 100
    return len(TABLE_TEXT) * max(ADVANCE.values()) * em


def column_chars(sizes: Sequence[float]) -> int:
    """Columns B onward in characters of the Normal font (7 pixels and 5 of padding at zoom
    100): the text at the largest size and 24 pixels to spare."""
    px = _text_px(max(sizes), 100) + 24
    return math.ceil((px - 5) / 7) + 1


def columns(zoom: int, sizes: Sequence[float]) -> int:
    """How many rule cells a row holds at ``zoom``: as many as fit the window beside the
    headings and column A, from 1 to 5."""
    width = (column_chars(sizes) * 7 + 5) * zoom / 100
    room = WINDOW_PX - (30 + COLUMN_A_CHARS * 7 + 5) * zoom / 100 - 60
    return max(1, min(len(TABLE_RULES), int(room // width)))


def _cost(zoom: int, sizes: Sequence[float]) -> float:
    """A zoom's pictures, roughly: rows a fifth taller than 1.2 em, and one more per face."""
    k = columns(zoom, sizes)
    rows = math.ceil(len(TABLE_RULES) / k)
    px = sum(rows * 1.45 * s * 4 / 3 * zoom / 100 for s in sizes)
    return len(TABLE_FACES) * (px / 1000 + 1)


def cover() -> dict[int, tuple[float, ...]]:
    """Each zoom's sizes: ``FIRST``, then zooms chosen greedily, most ppems still wanted per
    picture first (the lower zoom on a tie), until every ppem is drawn as often as ``wanted``
    asks and can be (at different zooms; a size a ppem at a zoom, the smallest)."""
    reach: dict[int, dict[int, float]] = {}
    for z in ZOOMS:
        found: dict[int, float] = {}
        for s in SIZES:
            p = ppem(s, z)
            if 1 <= p <= MAX_PPEM:
                found.setdefault(p, s)
        reach[z] = found
    want = {p: min(wanted(p), sum(p in r for r in reach.values())) for p in range(1, MAX_PPEM + 1)}
    drawn = dict.fromkeys(want, 0)
    chosen: dict[int, tuple[float, ...]] = {}
    for z, sizes in FIRST.items():
        chosen[z] = tuple(sizes)
        for p in {ppem(s, z) for s in sizes}:
            drawn[p] += 1
    while True:
        need = {p for p, n in want.items() if drawn[p] < n}
        if not need:
            break
        best, best_gain = 0, 0.0
        for z in ZOOMS:
            if z in chosen:
                continue
            got = need & reach[z].keys()
            if not got:
                continue
            gain = len(got) / _cost(z, [reach[z][p] for p in got])
            if gain > best_gain:
                best, best_gain = z, gain
        got = need & reach[best].keys()
        chosen[best] = tuple(sorted(reach[best][p] for p in got))
        for p in got:
            drawn[p] += 1
    return dict(sorted(chosen.items()))


@dataclass(frozen=True, slots=True)
class TableRow:
    """A sheet row: one face at one size under ``rules``, a cell each from column B; a filler
    (no rules) is a plain row."""

    tag: str
    face: str = "Calibri"
    size: float = 11.0
    rules: tuple[str, ...] = ()


def table_rows(
    zoom: int, sizes: Sequence[float], faces: Sequence[str] = TABLE_FACES
) -> list[TableRow]:
    """The rows of zoom ``zoom``'s workbook: each face at each size, the rules ``columns`` a
    row, then 20 fillers, so that the sheet does not end on rows taller than a view."""
    k = columns(zoom, sizes)
    out: list[TableRow] = []
    for face in faces:
        for size in sizes:
            for i in range(0, len(TABLE_RULES), k):
                rules = TABLE_RULES[i : i + k]
                if _tall(size, zoom) or (out and out[-1].rules and _tall(out[-1].size, zoom)):
                    out.append(TableRow(f"gap {len(out)}"))
                out.append(TableRow(f"{face} {size:g} {'+'.join(rules)}", face, size, rules))
    return out + [TableRow(f"filler {k}") for k in range(20)]


def _tall(size: float, zoom: int) -> bool:
    """Whether two rows of ``size`` may not fit a picture: then a plain row goes before each such
    row and the row after it, so that a view beginning at the plain row (its top pixel under the
    headings' border) shows the next whole."""
    return 1.6 * size * 4 / 3 * zoom / 100 > 450


def build(rows: Sequence[TableRow], chars: int, path: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    for r, row in enumerate(rows, 1):
        ws.cell(row=r, column=1).fill = markers[r % 2]
        for c, rule in enumerate(row.rules, 2):
            cell = ws.cell(row=r, column=c, value=TABLE_TEXT)
            cell.data_type = "s"
            underline = None if rule == "strike" else rule
            cell.font = Font(
                name=row.face, sz=row.size, family=2, u=underline, strike=rule == "strike"
            )
            cell.alignment = Alignment(horizontal="left", vertical="bottom")
    ws.column_dimensions["A"].width = COLUMN_A_CHARS
    for c in range(2, 2 + len(TABLE_RULES)):
        ws.column_dimensions[get_column_letter(c)].width = chars
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --- reading ------------------------------------------------------------------------------------


def _is_grid(rgb: tuple[int, int, int]) -> bool:
    """A gridline pixel: light gray, neither white nor ink."""
    r, g, b = rgb
    return max(r, g, b) - min(r, g, b) <= 12 and 170 <= (r + g + b) / 3 <= 245


def _gridline(image: Image, x_expected: int, y0: int, y1: int) -> int:
    """The gridline nearest ``x_expected`` (within 4 pixels) down the band, or exit."""
    import sys

    best, best_n = None, 0
    for x in sorted(range(x_expected - 4, x_expected + 5), key=lambda x: abs(x - x_expected)):
        if not 0 <= x < image.width:
            continue
        n = sum(_is_grid(image.rgb(x, y)) for y in range(y0, y1 + 1))
        if n > best_n:
            best, best_n = x, n
    if best is None or best_n < 0.8 * (y1 - y0 + 1):
        sys.exit(f"no gridline near x {x_expected} in rows {y0}-{y1}")
    return best


def cell_spans(
    image: Image,
    y0: int,
    y1: int,
    lefts: Sequence[float],
    widths: Sequence[float],
    scale: float,
    count: int,
) -> list[tuple[int, int]]:
    """``count`` columns from B between their gridlines, as (first, last) pixel columns.

    ``lefts`` and ``widths`` are the columns' (A onward) left edges and widths in points as Excel
    reports them; column A's right edge is found in the picture as the marker's end, each
    gridline from there by the points, snapped to the gray line nearest."""
    mid = (y0 + y1) // 2
    a_right = _marker_right(image, mid, _marker_column(image))  # column A's last pixel + 1
    spans = []
    left_px = a_right
    for c in range(1, count + 1):
        right_pts = lefts[c] + widths[c] - lefts[1]
        right_px = _gridline(image, a_right + math.floor(right_pts * scale + 0.5), y0, y1)
        spans.append((left_px + 2, right_px - 1))
        left_px = right_px
    return spans


def measure_table(
    views: Iterable[tuple[int, Path]],
    heights: list[float],
    zoom: int,
    scale: float,
    rows: Sequence[TableRow],
    lefts: Sequence[float],
    widths: Sequence[float],
    **frame: object,
) -> list[str]:
    """Every rule cell of a table workbook at one zoom: per cell the zoom, tag (face, size,
    rule), the row's height in points, device pixels a point, the cell's own baseline (the H
    stems' last ink row), the row's baseline (its single underline's cell, else its
    strikethrough's, else empty) and each line's first and last pixel row, from the row's top."""
    from capture.screen_rules import _rule_rows

    def measure_row(
        image: Image, row: int, band: tuple[int, int], top_px: int, bottom_px: int, s: float
    ) -> list[str]:
        case = rows[row - 1]
        if not case.rules:
            return []
        spans = cell_spans(image, top_px, bottom_px, lefts, widths, s, len(case.rules))
        found = {}
        for rule, (x0, x1) in zip(case.rules, spans, strict=True):
            stems, runs = _rule_rows(image, x0, x1 + 1, top_px, bottom_px)
            found[rule] = (None if stems is None else stems - top_px, runs)
        sibling = next((found[r][0] for r in ("single", "strike") if r in found), None)
        out = []
        for rule in case.rules:
            base, runs = found[rule]
            text = " ".join(f"{a - top_px}-{b - top_px}" for a, b in runs)
            out.append(
                f"{zoom}\t{case.face} {case.size:g} {rule}\t{heights[row - 1]:g}\t{s:g}"
                f"\t{'' if base is None else base}\t{'' if sibling is None else sibling}\t{text}"
            )
        return out

    found = read_views(
        views, heights, measure_row, label=f"table zoom {zoom}", match_scale=scale,
        pixels_per_point=scale, **frame,  # type: ignore[arg-type]
    )  # fmt: skip
    return [line for r in sorted(found) for line in found[r]]


def unmeasured_table(
    zoom: int, heights: list[float], scale: float, rows: Sequence[TableRow]
) -> list[str]:
    """The lines of a workbook whose pictures cannot be read (column A too narrow for its
    marker): every rule cell with nothing measured."""
    return [
        f"{zoom}\t{row.face} {row.size:g} {rule}\t{heights[r]:g}\t{scale:g}\t\t\t"
        for r, row in enumerate(rows)
        for rule in row.rules
    ]
