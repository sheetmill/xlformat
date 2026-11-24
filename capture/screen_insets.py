"""Where Excel starts and ends a cell's text, by the cell's font: the text insets on screen.

    uv run python -m capture.screen_insets run      # Excel for Mac (in the macOS VM): capture, read
    uv run python -m capture.screen_insets read     # read the Mac's saved pictures again

The same workbooks run on Excel for Windows as the group ``screen_insets`` of
``capture/screen_windows.py``. A row holds one face at one size, "HH" in four cells side by side:
left-aligned and right-aligned under a single underline, which runs from the text's origin to its
end, then left- and right-aligned under a single accounting underline, which spans the cell, so
those cells' text edges come from the H's ink less the bearing the plain cells show. Faces
Calibri, Arial and Aptos Narrow at every size from 9 to 36 points and every fourth from 40 to 72,
in workbooks of two Normal fonts (Calibri 11, Aptos Narrow 20), so that the cell's font can be told
from the Normal font's; seen at zoom 100 and 150.

The reader finds each row's gridlines in the picture (columns of light gray down the row's band)
(B's left one is column A's fill's last pixel column, which covers it) and gives, per cell, the
gridlines' pixel columns, the underline's first and last pixel column and the H's ink's first
and last, in device pixels.
"""

from __future__ import annotations

import itertools
import sys
import tarfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

from capture.cases import CellFont
from capture.screen_vertical import Image

FACES = ("Calibri", "Arial", "Aptos Narrow")
SIZES = (*range(9, 37), *range(40, 73, 4))
NORMALS = (CellFont("Calibri", 11), CellFont("Aptos Narrow", 20))
ZOOMS = (100, 150)
CELLS = (("left", "single"), ("right", "single"), ("left", "singleAccounting"),
         ("right", "singleAccounting"))  # fmt: skip
TEXT = "HH"
COLUMN_CHARS = 24  # wide enough for "HH" at 72 points with room either side


@dataclass(frozen=True, slots=True)
class InsetRow:
    """A sheet row: ``face`` at ``size`` in the four cells of ``CELLS``; a filler when no face."""

    tag: str
    face: str = ""
    size: float = 0.0

    @property
    def rules(self) -> tuple[str, ...]:  # the planner's and limit test's view of the cells
        return tuple(f"{a} {u}" for a, u in CELLS) if self.face else ()


@dataclass(frozen=True, slots=True)
class InsetBook:
    stem: str
    normal: CellFont
    rows: tuple[InsetRow, ...]


def books() -> list[InsetBook]:
    """Two workbooks a Normal font: sizes to 36, and above (each well under 1,000 fonts)."""
    out = []
    for normal in NORMALS:
        label = normal.label.replace(" ", "_").lower()
        for part, sizes in (
            ("small", [s for s in SIZES if s <= 36]),
            ("large", [s for s in SIZES if s > 36]),
        ):
            rows = [InsetRow("filler top")]
            rows += [InsetRow(f"{face} {s:g}", face, float(s)) for face in FACES for s in sizes]
            rows += [InsetRow(f"filler {k}") for k in range(20)]
            out.append(InsetBook(f"insets_{label}_{part}", normal, tuple(rows)))
    return out


def build(book: InsetBook, path: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.indexed_list import IndexedList

    wb = Workbook()
    n = book.normal
    base = Font(name=n.name, sz=n.size, family=2, scheme="minor" if n.name == "Calibri" else None)
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    for r, row in enumerate(book.rows, 1):
        ws.cell(row=r, column=1).fill = markers[r % 2]
        if not row.face:
            continue
        for c, (align, underline) in enumerate(CELLS, 2):
            cell = ws.cell(row=r, column=c, value=TEXT)
            cell.data_type = "s"
            cell.font = Font(name=row.face, sz=row.size, family=2, u=underline)
            cell.alignment = Alignment(horizontal=align, vertical="bottom")
    ws.column_dimensions["A"].width = 4
    for c in range(2, 2 + len(CELLS)):
        ws.column_dimensions[get_column_letter(c)].width = COLUMN_CHARS
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --- reading ------------------------------------------------------------------------------------


def _gridlines(image: object, x0: int, y0: int, y1: int, count: int) -> list[int]:
    """The first ``count`` gridlines right of ``x0`` down the band: each a run of light gray
    columns (two pixels on the Mac's display), given by its first column."""
    from capture.screen_table import _is_grid

    img = image  # Image
    out: list[int] = []
    x, height = x0, y1 - y0 + 1
    while x < img.width and len(out) < count:  # type: ignore[attr-defined]
        n = sum(_is_grid(img.rgb(x, y)) for y in range(y0, y1 + 1))  # type: ignore[attr-defined]
        if n >= 0.8 * height and (not out or x > out[-1] + 3):
            out.append(x)
        x += 1
    if len(out) < count:
        sys.exit(f"found {len(out)} of {count} gridlines in rows {y0}-{y1}")
    return out


def measure_cell(image: object, x0: int, x1: int, y0: int, y1: int) -> tuple[int, ...] | None:
    """(line first, line last, ink first, ink last) between columns x0 and x1, or None: the line
    is the pixel row whose ink is one run longest; the ink, every other row's."""
    from capture.screen_vertical import is_ink

    img = image
    line, ink_lo, ink_hi = None, None, None
    rows = []
    for y in range(y0, y1 + 1):
        xs = [x for x in range(x0, x1 + 1) if is_ink(img.rgb(x, y))]  # type: ignore[attr-defined]
        if xs:
            rows.append(xs)
    if not rows:
        return None
    for xs in rows:
        if len(xs) == xs[-1] - xs[0] + 1 and (line is None or len(xs) > line[1] - line[0] + 1):
            line = (xs[0], xs[-1])
    for xs in rows:
        if line is not None and (xs[0], xs[-1]) == line and len(xs) == line[1] - line[0] + 1:
            continue
        ink_lo = xs[0] if ink_lo is None else min(ink_lo, xs[0])
        ink_hi = xs[-1] if ink_hi is None else max(ink_hi, xs[-1])
    if line is None or ink_lo is None or ink_hi is None:
        return None
    return (line[0], line[1], ink_lo, ink_hi)


def measure_insets(
    views: Iterable[tuple[int, Path]],
    heights: list[float],
    zoom: int,
    rows: Sequence[InsetRow],
    normal: CellFont,
    **frame: object,
) -> list[str]:
    """Per face row: the Normal font, zoom, tag, row height, then per cell (``CELLS``) the left
    and right gridlines, the line's ends and the ink's ends in device pixels, ``-`` where none
    reads."""
    from capture.screen_vertical import _marker_column, _marker_right, read_views

    def measure_row(
        image: Image, row: int, band: tuple[int, int], top_px: int, bottom_px: int, s: float
    ) -> str | None:
        case = rows[row - 1]
        if not case.face:
            return None
        a_right = _marker_right(image, (top_px + bottom_px) // 2, _marker_column(image))
        # column A's fill covers the gridline between A and B: its last pixel column is B's
        # left gridline; the others are drawn
        lines = [a_right - 1, *_gridlines(image, a_right, top_px + 1, bottom_px - 1, len(CELLS))]
        cells = []
        for gl, gr in itertools.pairwise(lines):
            # the cell's pixels: past the left gridline's run, up to the right gridline
            x0 = gl + 1
            while x0 < gr and _grid(image, x0, top_px, bottom_px):
                x0 += 1
            m = measure_cell(image, x0, gr - 1, top_px, bottom_px)
            cells.append(f"{gl} {gr} " + ("- - - -" if m is None else " ".join(map(str, m))))
        return f"{normal.label}\t{zoom}\t{case.tag}\t{heights[row - 1]:g}\t{s:g}\t" + "\t".join(
            cells
        )

    found = read_views(views, heights, measure_row, label=f"insets zoom {zoom}", **frame)  # type: ignore[arg-type]
    return [line for r in sorted(found) if (line := found[r]) is not None]


def _grid(image: object, x: int, y0: int, y1: int) -> bool:
    from capture.screen_table import _is_grid

    n = sum(_is_grid(image.rgb(x, y)) for y in range(y0, y1 + 1))  # type: ignore[attr-defined]
    return n >= 0.8 * (y1 - y0 + 1)


COLUMNS = "normal_font\tzoom\ttag\trow_height\tpixels_per_point\t" + "\t".join(
    f"{a}_{u}" for a, u in CELLS
)
METHOD = (
    "window captures; per cell: left gridline, right gridline (first pixel column of each), the "
    "underline's first and last pixel column, the H's ink's first and last, in device pixels"
)


# --- Excel for Mac (run in the macOS VM) ---------------------------------------------------------


def mac_output() -> Path:
    from capture.screen_rules import CAPTURES

    return CAPTURES / "insets.layout.tsv"


def mac_shots() -> Path:
    from capture.build import BUILD_DIR

    return BUILD_DIR / "screen_insets"


def run() -> None:
    """Every workbook at every zoom on Excel for Mac, its pictures saved, then ``read``."""
    from capture.build import WORKBOOK_DIR
    from capture.excel_watch import answers, excel_windows
    from capture.screen_vertical import (
        _capture,
        _request,
        _request_count,
        _window_id,
        read_png,
        rows_in_view,
    )

    if not answers():
        sys.exit("Excel does not answer: nothing run")
    if _request_count() != "0" or excel_windows():
        sys.exit("Excel must hold no workbook and show no window: nothing run")
    for book in books():
        name = f"{book.stem}.xlsx"
        path = build(book, (WORKBOOK_DIR / name).resolve())
        n = len(book.rows)
        _request("open", str(path), name, seconds=120.0)
        try:
            heights = _request("heights", name, str(n), seconds=120.0)
            window = _window_id(book.stem)
            for zoom in ZOOMS:
                folder = mac_shots() / book.stem / f"zoom{zoom}"
                folder.mkdir(parents=True, exist_ok=True)
                for old in folder.glob("*.png"):
                    old.unlink()
                (folder / "heights.txt").write_text(heights + "\n")
                if float(_request("zoom", name, str(zoom))) != zoom:
                    sys.exit(f"zoom {zoom} not set")
                top = 1
                while top <= n:
                    if _request("scroll", name, str(top)) != str(top):
                        sys.exit(f"row {top} not at the top")
                    shot = folder / f"top_{top:03d}.png"
                    _capture(window, shot)
                    hs = [float(h) for h in heights.split(",")]
                    found = rows_in_view(
                        read_png(shot), top, hs, None if zoom == 100 else 2 * zoom / 100
                    )
                    if not found or found[0][0] != top:
                        sys.exit(f"{shot.name}: row {top} not found whole at the top")
                    last = found[-1][0]
                    top = last if top < last < n else last + 1
        finally:
            _request("close", name)
    with tarfile.open(mac_shots().with_suffix(".tar"), "w") as tar:
        tar.add(mac_shots(), arcname="screen_insets")
    read()


def read() -> None:
    """The Mac's layout file from its saved pictures."""
    from capture.edition import edition_line, excel_build
    from capture.screen_vertical import numbered_views

    lines = []
    for book in books():
        for zoom in ZOOMS:
            folder = mac_shots() / book.stem / f"zoom{zoom}"
            heights = [float(h) for h in (folder / "heights.txt").read_text().strip().split(",")]
            own = zoom != 100  # rows rounded to whole pixels: each row's own band
            lines += measure_insets(
                numbered_views(folder), heights, zoom, book.rows, book.normal,
                match_scale=None if zoom == 100 else 2 * zoom / 100, own_bands=own,
            )  # fmt: skip
    with mac_output().open("w", encoding="utf-8", newline="\n") as out:
        out.write("# suite: insets (capture.screen_insets); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(f"# method: {METHOD}; a display of scale 2\n")
        out.write(COLUMNS + "\n" + "\n".join(lines) + "\n")
    print(f"wrote {mac_output()}: {len(lines)} rows")


if __name__ == "__main__":
    {"run": run, "read": read}[sys.argv[1]]()
