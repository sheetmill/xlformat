"""The window-capture reader at a scale that is not a whole number of pixels a point.

Excel for Windows at 96 dpi draws a sheet point as 4/3 device pixels, where the Mac's display of
scale 2 draws it as 2. ``read_views`` takes the scale as a parameter; these pictures are drawn
here, with column A's marker bands and an "H" of known pixel rows, and read back.
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

import pytest

pytest.importorskip("openpyxl")  # capture.screen_vertical imports the workbook builder

from capture.screen_vertical import measure, read_views

WHITE, MAGENTA, CYAN, BLACK = (255, 255, 255), (255, 0, 255), (0, 255, 255), (0, 0, 0)
SCALE = 4 / 3  # device pixels a sheet point at 96 dpi and zoom 100
HEIGHTS = [15.0, 18.75, 15.0, 30.0, 15.0]  # points; 20, 25, 20, 40, 20 pixels
WIDTH = 120


def _png(path: Path, pixels: list[list[tuple[int, int, int]]]) -> None:
    """A minimal 8-bit RGB PNG, every row filtered with type 0."""

    def chunk(kind: bytes, body: bytes) -> bytes:
        crc = zlib.crc32(kind + body)
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", crc)

    raw = b"".join(b"\0" + bytes(c for px in row for c in px) for row in pixels)
    header = struct.pack(">IIBBBBB", len(pixels[0]), len(pixels), 8, 2, 0, 0, 0)
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _view(first_row: int, offset: int) -> tuple[list[list[tuple[int, int, int]]], dict[int, int]]:
    """A picture whose first visible row starts ``offset`` pixels down; per row, its top pixel.

    Each row's band fills columns 0 to 9 in magenta or cyan in turn, and holds an "H" (columns 20
    to 26) from 4 pixels below the row's top to 6 pixels above its bottom.
    """
    rows = [[WHITE] * WIDTH for _ in range(offset)]
    tops: dict[int, int] = {}
    for row in range(first_row, len(HEIGHTS) + 1):
        tops[row] = len(rows)
        height = round(HEIGHTS[row - 1] * SCALE)
        fill = MAGENTA if row % 2 else CYAN
        for y in range(height):
            line = [fill] * 10 + [WHITE] * (WIDTH - 10)
            if 4 <= y <= height - 7:
                line[20] = line[26] = BLACK  # the stems
                if y == height // 2:
                    line[20:27] = [BLACK] * 7  # the bar
            rows.append(line)
    return rows + [[WHITE] * WIDTH for _ in range(6)], tops


def test_rows_and_ink_at_four_thirds_of_a_pixel_a_point(tmp_path: Path) -> None:
    views = []
    for first_row, offset in ((1, 3), (3, 7)):
        pixels, _ = _view(first_row, offset)
        path = tmp_path / f"view{first_row}.png"
        _png(path, pixels)
        views.append((first_row, path))

    def measure_row(image, row, band, top_px, bottom_px, scale):  # type: ignore[no-untyped-def]
        assert scale == SCALE
        return band, top_px, bottom_px, measure(image, (top_px, bottom_px))

    found = read_views(views, HEIGHTS, measure_row, label="test", pixels_per_point=SCALE)
    assert sorted(found) == [1, 2, 3, 4, 5]
    for row, (band, top_px, bottom_px, (h_lines, g_bottoms)) in found.items():
        height = round(HEIGHTS[row - 1] * SCALE)
        assert band == (top_px, bottom_px)
        assert bottom_px - top_px + 1 == height
        # "H" alone: its cap top and baseline, and its bottom as the last glyph's
        assert h_lines == [(4, height - 7)]
        assert g_bottoms == [height - 7]
