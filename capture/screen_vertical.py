"""Where Excel puts text vertically in a cell, from the screen (CAPTURING.md section 2 item 47).

    uv run python -m capture.screen_vertical build    # the workbooks, no Excel
    uv run python -m capture.screen_vertical run      # open, capture every view, close; then read
    uv run python -m capture.screen_vertical read     # measure the saved captures again

The cases of ``capture/vertical_pdf.py`` (the text "Hxg" in every measured face, aligned top,
center and bottom in rows of three heights, rows Excel fits, wrapped lines, the styles), one per
row in column B, with column A filled magenta and cyan in turn: on screen each row's marker band
is the row between its gridlines, so no geometry is mapped. Excel opens the workbook at zoom 100,
in front; the run scrolls so that each row in turn is at the top, captures Excel's window with
``screencapture -l`` (CoreGraphics finds it; the image is at the display's scale), and repeats
a capture a second later until two are identical, Excel possibly still drawing. Bands are matched
to rows by the heights Excel reports. In each row the first glyph's ink gives the cap top and the
baseline (the bottom of "H"), the last glyph's the descender's bottom ("g"), per wrapped line.

The run needs the display on, Excel in front and nobody at the keyboard; it sends Excel one
request a step, starts only when Excel answers with no workbook open and no window of its own,
and closes the workbook without saving, alerts on, on every path. No clipboard. Writes
``tests/captures/vertical.layout.tsv``, every measured distance in image pixels from the row's top,
the scale in the header; the captures stay in ``capture/build/screen_vertical/``.
"""

from __future__ import annotations

import datetime as dt
import struct
import subprocess
import sys
import time
import zlib
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import TypeVar

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.cases import CellFont
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, excel_windows, osascript_or_exit
from capture.vertical_pdf import NORMALS, TEXT, Case, cases

SCRIPT = Path(__file__).with_name("screen_vertical.applescript")
SHOTS = BUILD_DIR / "screen_vertical"
OUTPUT = BUILD_DIR.parent.parent / "tests" / "captures" / "vertical.layout.tsv"
_WINDOWS_JXA = r"""
ObjC.import("CoreGraphics");
const list = ObjC.castRefToObject(
  $.CGWindowListCopyWindowInfo($.kCGWindowListOptionOnScreenOnly, 0));
const out = [];
for (let i = 0; i < list.count; i++) {
  const w = list.objectAtIndex(i);
  if (ObjC.unwrap(w.objectForKey("kCGWindowOwnerName")) !== "Microsoft Excel") continue;
  if (ObjC.unwrap(w.objectForKey("kCGWindowLayer")) !== 0) continue;
  const title = ObjC.unwrap(w.objectForKey("kCGWindowName")) || "";
  out.push([ObjC.unwrap(w.objectForKey("kCGWindowNumber")), title].join("\t"));
}
out.join("\n");
"""


# --- the workbook ------------------------------------------------------------------------------


def _paths(normal: CellFont) -> tuple[str, Path]:
    name = "screen_vertical_" + normal.label.replace(" ", "_").lower() + ".xlsx"
    return name, (WORKBOOK_DIR / name).resolve()


def build(normal: CellFont, case_list: list[Case] | None = None, path: Path | None = None) -> Path:
    """The cases of the PDF probe in column B, column A in marker fills, in Excel's container
    (or ``case_list`` at ``path``)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils.indexed_list import IndexedList

    wb = Workbook()
    base = Font(
        name=normal.name, sz=normal.size, family=2,
        scheme="minor" if normal.name == "Calibri" else None,
    )  # fmt: skip
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    ws.title = "vertical"
    # two marker colors in turn: Excel draws no gridline between cells of one fill, so rows of
    # a single color would merge into one band on screen
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    for r, case in enumerate(cases() if case_list is None else case_list, 1):
        ws.cell(row=r, column=1).fill = markers[r % 2]
        cell = ws.cell(row=r, column=2, value="\n".join([TEXT] * case.lines))
        cell.data_type = "s"
        f = case.font
        cell.font = Font(name=f.name, sz=f.size, b=f.bold, i=f.italic, family=2)
        cell.alignment = Alignment(vertical=case.vertical, wrap_text=case.lines > 1)
        if case.height is not None:
            ws.row_dimensions[r].height = case.height
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    path = _paths(normal)[1] if path is None else path
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --- images ------------------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Image:
    """An 8-bit RGB(A) image: ``rgb(x, y)`` gives a pixel's color."""

    width: int
    height: int
    channels: int
    rows: list[bytes]

    def rgb(self, x: int, y: int) -> tuple[int, int, int]:
        i = x * self.channels
        row = self.rows[y]
        return row[i], row[i + 1], row[i + 2]


def read_png(path: Path) -> Image:
    """Decode a non-interlaced 8-bit RGB or RGBA PNG (what ``screencapture`` writes)."""
    data = path.read_bytes()
    assert data[:8] == b"\x89PNG\r\n\x1a\n", path
    pos, idat = 8, b""
    width = height = channels = 0
    while pos < len(data):
        (length,) = struct.unpack(">I", data[pos : pos + 4])
        kind = data[pos + 4 : pos + 8]
        body = data[pos + 8 : pos + 8 + length]
        if kind == b"IHDR":
            width, height, depth, color, _, _, interlace = struct.unpack(">IIBBBBB", body)
            if depth != 8 or color not in (2, 6) or interlace:
                sys.exit(f"{path.name}: PNG depth {depth}, color {color}, interlace {interlace}")
            channels = 3 if color == 2 else 4
        elif kind == b"IDAT":
            idat += body
        pos += 12 + length
    raw = zlib.decompress(idat)
    stride = width * channels
    rows: list[bytes] = []
    prev = bytearray(stride)
    for y in range(height):
        start = y * (stride + 1)
        kind, line = raw[start], bytearray(raw[start + 1 : start + 1 + stride])
        for i in range(stride):
            a = line[i - channels] if i >= channels else 0
            b = prev[i]
            c = prev[i - channels] if i >= channels else 0
            if kind == 1:
                line[i] = (line[i] + a) & 255
            elif kind == 2:
                line[i] = (line[i] + b) & 255
            elif kind == 3:
                line[i] = (line[i] + (a + b) // 2) & 255
            elif kind == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append(bytes(line))
        prev = line
    return Image(width, height, channels, rows)


def write_png(path: Path, pixels: list[list[tuple[int, int, int]]]) -> None:
    """An 8-bit RGB PNG of ``pixels`` (rows of colors), every row unfiltered: for pictures drawn
    in tests."""

    def chunk(kind: bytes, body: bytes) -> bytes:
        return (
            struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
        )

    raw = b"".join(b"\0" + bytes(c for px in row for c in px) for row in pixels)
    head = struct.pack(">IIBBBBB", len(pixels[0]), len(pixels), 8, 2, 0, 0, 0)
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", head)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def marker(rgb: tuple[int, int, int]) -> str:
    """``"m"`` for magenta, ``"c"`` for cyan, ``""`` for anything else."""
    r, g, b = rgb
    if g < 90 and r > 200 and b > 200:
        return "m"
    if r < 160 and g > 200 and b > 200:  # cyan shows as about (117, 251, 253) on screen
        return "c"
    return ""


def is_marker(rgb: tuple[int, int, int]) -> bool:
    """Either marker color, magenta or cyan."""
    return marker(rgb) != ""


def is_ink(rgb: tuple[int, int, int]) -> bool:
    """Dark enough to be text, light gray gridlines and white excluded."""
    r, g, b = rgb
    return 0.299 * r + 0.587 * g + 0.114 * b < 140


# --- Excel -------------------------------------------------------------------------------------


def _request(*arguments: str, seconds: float = 60.0) -> str:
    proc = osascript_or_exit(str(SCRIPT), *arguments, seconds=seconds)
    if proc.returncode != 0:
        sys.exit(f"Excel request {arguments[0]} failed: {proc.stderr.strip()[:200]}")
    return proc.stdout.strip()


def _window_id(title: str) -> str:
    for _ in range(20):
        out = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", _WINDOWS_JXA],
            capture_output=True, text=True, timeout=10, check=False,
        ).stdout  # fmt: skip
        for line in out.strip().splitlines():
            number, _, name = line.partition("\t")
            if title in name:
                return number
        time.sleep(0.5)
    sys.exit(f"no window of Excel's titled like {title!r} on screen")


def _capture(window: str, out: Path) -> None:
    """Capture the window until two shots a second apart are identical."""
    previous = b""
    for _ in range(10):
        subprocess.run(["screencapture", "-x", "-o", "-l", window, str(out)], check=True)
        data = out.read_bytes()
        if data == previous:
            return
        previous = data
        time.sleep(1.0)
    sys.exit(f"{out.name}: the window kept changing")


def run_normal(normal: CellFont) -> None:
    """Open one workbook, capture a view per scroll position, close it."""
    name, path = _paths(normal)
    build(normal)
    folder = SHOTS / normal.label.replace(" ", "_").lower()
    folder.mkdir(parents=True, exist_ok=True)
    for old in folder.glob("*.png"):
        old.unlink()
    _request("open", str(path), name, seconds=120.0)
    try:
        n = len(cases())
        heights = _request("heights", name, str(n), seconds=120.0)
        (folder / "heights.txt").write_text(heights + "\n")
        window = _window_id(name.removesuffix(".xlsx"))
        top = 1
        while top <= n:
            reported = _request("scroll", name, str(top))
            if reported != str(top):
                sys.exit(f"asked for row {top} at the top, Excel reports {reported}")
            shot = folder / f"top_{top:03d}.png"
            _capture(window, shot)
            rows = rows_in_view(read_png(shot), top, [float(h) for h in heights.split(",")])
            if not rows or rows[0][0] != top:
                sys.exit(f"{shot.name}: row {top} not found whole at the top")
            if len(rows) < 3 and rows[-1][0] < n:
                sys.exit(f"{shot.name}: only {len(rows)} rows recognized in a view: stopping")
            # the view's last row can lie under the sheet-tab bar at the window's bottom edge, so
            # the next view starts at it: every row is also captured above the edge
            last = rows[-1][0]
            top = last if last > top and last < n else last + 1
    finally:
        _request("close", name)
    print(f"{normal.label}: captured {len(list(folder.glob('*.png')))} views")


# --- measuring ---------------------------------------------------------------------------------


def _marker_column(image: Image) -> int:
    """An x inside column A's marker fill, searched from the image's top (the last view of a
    sheet holds only a few rows, near the top). The fill runs down at least 8 pixels, which no
    icon of Excel for Windows' ribbon in a marker's colors does."""
    for y in range(image.height - 8):
        for x in range(image.width // 2):
            if (
                is_marker(image.rgb(x, y))
                and is_marker(image.rgb(x + 4, y))
                and all(is_marker(image.rgb(x + 2, y + k)) for k in range(1, 8))
            ):
                return x + 2
    sys.exit("no marker column in the capture")


MIN_BAND_PX = 4  # a row's band is taller: a shorter run is an icon above the cells


def _bands(image: Image) -> list[tuple[int, int]]:
    """(first, last) pixel rows of each marker band, top to bottom, at least ``MIN_BAND_PX``
    tall (Excel for Windows' ribbon has a light blue icon in the cyan marker's range)."""
    x = _marker_column(image)
    out: list[tuple[int, int]] = []
    start, current = None, ""
    for y in range(image.height + 1):
        m = marker(image.rgb(x, y)) if y < image.height else ""
        if m != current:
            if current:
                out.append((start or 0, y - 1))
            start, current = (y if m else None), m
    return [b for b in out if b[1] - b[0] + 1 >= MIN_BAND_PX]


def rows_in_view(
    image: Image, top: int, heights: list[float], scale: float | None = None
) -> list[tuple[int, tuple[int, int]]]:
    """(row, band) for each band whose height in pixels fits its row's height at the display's
    scale (given, or the median over the view); a band the window's edge cuts off does not."""
    bands = _bands(image)
    pairs = [(top + k, b) for k, b in enumerate(bands) if top + k <= len(heights)]
    if not pairs:
        return []
    if scale is None:
        ratios = sorted((b[1] - b[0] + 1) / heights[r - 1] for r, b in pairs)
        scale = ratios[len(ratios) // 2]
    return [(r, b) for r, b in pairs if abs(b[1] - b[0] + 1 - heights[r - 1] * scale) <= 3]


T = TypeVar("T")
# measure(image, row, band, top_px, bottom_px, pixels_per_point) -> the row's measurement
RowMeasure = Callable[[Image, int, tuple[int, int], int, int, float], T]


def read_views(
    views: Iterable[tuple[int, Path]],
    heights: list[float],
    measure_row: RowMeasure[T],
    *,
    label: str,
    match_scale: float | None = None,
    pixels_per_point: float | None = None,
    own_bands: bool = False,
    row_px: Callable[[float], int] | None = None,
    unmeasured_ok: frozenset[int] = frozenset(),
) -> dict[int, T]:
    """Measure every row of a sheet once, from views whose first visible row is given.

    ``views`` are (first row, picture) pairs; ``heights`` every row's height in points as Excel
    reports it. Bands are matched to rows by ``rows_in_view`` at ``match_scale`` (the view's median
    when None). A row's pixels run from its top, computed from the view's first band plus the
    heights above it at ``pixels_per_point`` device pixels a point (when None, the view's median
    rounded to whole pixels: 2 on a Mac display of scale 2), or with ``own_bands`` from its own
    band (for a zoom at which Excel rounds each row to whole pixels). A row seen whole in a later
    view replaces one seen only as a view's last band. Exits naming ``label`` if a row is never
    measured, other than those of ``unmeasured_ok``.

    With ``row_px`` (each row's height in whole device pixels, as Excel draws it) the view's frame
    is anchored on its bands after the first instead: every one must begin at the same offset
    from the heights above it (or the reader exits: the frame is not a constant offset), and the
    first band, whose top pixel the column headings' border covers, is not measured.
    """
    found: dict[int, T] = {}
    at_edge: set[int] = set()  # rows measured only as a view's last band so far
    n_rows = len(heights)
    for top, path in sorted(views):
        image = read_png(path)
        rows = rows_in_view(image, top, heights, match_scale)
        if not rows:
            continue
        first_row, first_band = rows[0]
        anchor = None
        if row_px is not None:
            offsets = {
                band[0] - sum(row_px(h) for h in heights[first_row - 1 : row - 1])
                for row, band in rows[1:]
            }
            if len(offsets) != 1:
                sys.exit(f"{label}: {path.name}: bands not at one offset from the rows: {offsets}")
            (anchor,) = offsets
        if pixels_per_point is None:
            ratios = sorted((b[1] - b[0] + 1) / heights[r - 1] for r, b in rows)
            scale: float = round(ratios[len(ratios) // 2])
        else:
            scale = pixels_per_point
        for k, (row, band) in enumerate(rows):
            if anchor is not None and k == 0:
                continue  # the headings' border covers its top pixel
            edge = k == len(rows) - 1 and row < n_rows
            if row in found and (row not in at_edge or edge):
                continue
            if anchor is not None and row_px is not None:
                top_px = anchor + sum(row_px(h) for h in heights[first_row - 1 : row - 1])
                bottom_px = top_px + row_px(heights[row - 1]) - 1
            elif own_bands:
                top_px, bottom_px = band
            else:
                # boundaries between filled cells are sometimes drawn a pixel row or two off, so
                # distances are taken from the row's top computed from the view's first band
                top_px = first_band[0] + round(scale * sum(heights[first_row - 1 : row - 1]))
                bottom_px = top_px + round(scale * heights[row - 1]) - 1
            found[row] = measure_row(image, row, band, top_px, bottom_px, scale)
            if edge:
                at_edge.add(row)
            else:
                at_edge.discard(row)
    missing = [r for r in range(1, n_rows + 1) if r not in found and r not in unmeasured_ok]
    if missing:
        sys.exit(f"{label}: rows never measured: {missing[:10]}")
    return found


def numbered_views(folder: Path) -> list[tuple[int, Path]]:
    """The Mac run's views: ``top_<first row>.png`` in ``folder``."""
    return [(int(p.stem.split("_")[1]), p) for p in folder.glob("top_*.png")]


def _marker_right(image: Image, y: int, x: int) -> int:
    while x < image.width and is_marker(image.rgb(x, y)):
        x += 1
    return x


def _glyph_runs(image: Image, x0: int, x1: int, y0: int, y1: int) -> list[tuple[int, int]]:
    """Runs of columns holding ink between rows y0 and y1."""
    runs: list[tuple[int, int]] = []
    start = None
    for x in range(x0, x1):
        has = any(is_ink(image.rgb(x, y)) for y in range(y0, y1 + 1))
        if has and start is None:
            start = x
        elif not has and start is not None:
            runs.append((start, x - 1))
            start = None
    return runs


def _ink_rows(image: Image, x0: int, x1: int, y0: int, y1: int) -> list[tuple[int, int]]:
    """Runs of pixel rows holding ink between columns x0 and x1."""
    runs: list[tuple[int, int]] = []
    start = None
    for y in range(y0, y1 + 2):
        has = y <= y1 and any(is_ink(image.rgb(x, y)) for x in range(x0, x1 + 1))
        if has and start is None:
            start = y
        elif not has and start is not None:
            runs.append((start, y - 1))
            start = None
    return runs


def measure(image: Image, band: tuple[int, int]) -> tuple[list[tuple[int, int]], list[int]]:
    """Per line of text in the band: (cap top, baseline) of "H", and the bottom of "g", in pixels
    from the band's first row."""
    y0, y1 = band
    left = _marker_right(image, (y0 + y1) // 2, _marker_column(image)) + 2
    runs = _glyph_runs(image, left, min(image.width, left + 600), y0, y1)
    if not runs:
        return [], []
    if len(runs) == 1:
        # italic glyphs touch: H's foot is in the run's leftmost columns, its top (the tallest
        # glyph of "Hxg") the run's highest ink, g's tail the run's lowest ink, per line
        x0, x1 = runs[0]
        feet = _ink_rows(image, x0, min(x1, x0 + 3), y0, y1)
        g_runs = _ink_rows(image, x0, x1, y0, y1)
        h_runs = [(whole[0], foot[1]) for whole, foot in zip(g_runs, feet, strict=False)]
    else:
        h_runs = _ink_rows(image, runs[0][0], runs[0][1], y0, y1)
        g_runs = _ink_rows(image, runs[-1][0], runs[-1][1], y0, y1)
    return [(a - y0, b - y0) for a, b in h_runs], [b - y0 for _, b in g_runs]


@dataclass(frozen=True, slots=True)
class Measured:
    case: Case
    height: float
    band: tuple[int, int]
    h_lines: list[tuple[int, int]]
    g_bottoms: list[int]
    shift: int  # the band's first row less the row's top computed from the view's first band


def read_normal(normal: CellFont) -> list[Measured]:
    """Match each view's bands to rows by height and measure every row once."""
    folder = SHOTS / normal.label.replace(" ", "_").lower()
    heights = [float(h) for h in (folder / "heights.txt").read_text().strip().split(",")]
    return measure_vertical(numbered_views(folder), heights, normal.label)


def measure_vertical(
    views: Iterable[tuple[int, Path]],
    heights: list[float],
    label: str,
    pixels_per_point: float | None = None,
    case_list: list[Case] | None = None,
    row_px: Callable[[float], int] | None = None,
    unmeasured_ok: frozenset[int] = frozenset(),
) -> list[Measured]:
    """Every row of the vertical cases (or ``case_list``) measured from its views
    (``read_views``), in row order."""
    all_cases = cases() if case_list is None else case_list

    def measure_row(
        image: Image, row: int, band: tuple[int, int], top_px: int, bottom_px: int, _: float
    ) -> Measured:
        h_lines, g_bottoms = measure(image, (top_px, bottom_px))
        return Measured(
            all_cases[row - 1], heights[row - 1], band, h_lines, g_bottoms, band[0] - top_px
        )

    found = read_views(
        views,
        heights,
        measure_row,
        label=label,
        pixels_per_point=pixels_per_point,
        row_px=row_px,
        unmeasured_ok=unmeasured_ok,
    )
    if len(found) + len(unmeasured_ok - set(found)) != len(all_cases):
        sys.exit(f"{label}: {len(found)} rows measured for {len(all_cases)} cases")
    # the band heights must follow the row heights at one scale
    off = [(m.case.tag, m.shift) for m in found.values() if abs(m.shift) > 2]
    if off:
        sys.exit(f"{label}: bands off their computed rows by more than 2 px: {off[:3]}")
    return [found[r] for r in sorted(found)]


def write() -> None:
    lines = []
    for normal in NORMALS:
        for m in read_normal(normal):
            band_px = m.band[1] - m.band[0] + 1
            h = " ".join(f"{a}-{b}" for a, b in m.h_lines)
            g = " ".join(str(b) for b in m.g_bottoms)
            stated = "" if m.case.height is None else f"{m.case.height:g}"
            lines.append(
                f"{normal.label}\t{m.case.tag}\t{stated}\t{m.height:g}\t{band_px}\t{m.shift}"
                f"\t{h}\t{g}"
            )
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: vertical (capture.screen_vertical); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; method: window capture at zoom 100; "
            "pixels (2 a point) from each row's top, computed from the view's first marker band "
            "(column A); H: cap top-baseline per "
            "line, g: descender bottom per line\n"
        )
        out.write(
            "normal_font\ttag\tstated_height\trow_height\tband_px\tband_shift\tH_top-bottom"
            "\tg_bottom\n"
        )
        out.write("\n".join(lines) + "\n")
    print(f"wrote {OUTPUT}: {len(lines)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        for normal in NORMALS:
            print(build(normal))
    elif command == "run":
        if not answers():
            sys.exit("Excel does not answer: nothing run")
        count = _request_count()
        if count != "0" or excel_windows():
            sys.exit(f"Excel must hold no workbook and show no window ({count}): nothing run")
        for normal in NORMALS:
            run_normal(normal)
        write()
    elif command == "read":
        write()
    else:
        sys.exit(__doc__)


def _request_count() -> str:
    proc = osascript_or_exit(
        "-e", 'tell application "Microsoft Excel" to get count of workbooks', seconds=20.0
    )
    return proc.stdout.strip()


if __name__ == "__main__":
    main(sys.argv[1:])
