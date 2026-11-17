"""``section_rgb`` against ``tests/captures/palette.layout.tsv`` (CAPTURING.md §2 item 50).

Each row is a cell's text color as measured on screen (sRGB), under a custom 64-color workbook
palette or the default one, with the candidates it was matched against; the workbook's palette
is rebuilt from those candidates.
"""

from __future__ import annotations

import csv
from pathlib import Path

import pytest

from xlformat import DEFAULT_PALETTE, section_rgb

LAYOUT = Path(__file__).parent / "captures" / "palette.layout.tsv"
NAMES = ("Black", "White", "Red", "Green", "Blue", "Yellow", "Magenta", "Cyan")
TOLERANCE = 3.5  # the measured colors are within 3.2 of their source after the conversion to sRGB


def _rows() -> list[list[str]]:
    with LAYOUT.open(encoding="utf-8") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.reader(lines, dialect="excel-tab"))[1:]


def _index(tag: str) -> int:
    return NAMES.index(tag) if tag in NAMES else int(tag[5:]) + 7


def _distance(a: str, b: str) -> float:
    pairs = zip(bytes.fromhex(a), bytes.fromhex(b), strict=True)
    return float(sum((x - y) ** 2 for x, y in pairs) ** 0.5)


def test_every_measured_cell() -> None:
    rows = _rows()
    assert len(rows) == 128
    custom = [""] * 64
    for kind, tag, *_, candidates in rows:
        if kind == "custom":
            workbook = dict(c.split("=") for c in candidates.split())["workbook"]
            custom[_index(tag)] = workbook.split(":")[0]
    misses = []
    for kind, tag, measured, *_ in rows:
        want = section_rgb(tag, custom if kind == "custom" else None)
        if _distance(measured, want) > TOLERANCE:
            misses.append((kind, tag, measured, want))
    assert not misses, misses[:10]


def test_entries_and_short_palettes() -> None:
    palette = ["FF000000"] * 64
    palette[12] = "ff123456"  # Color5 is entry 12, in AARRGGBB as openpyxl writes it
    assert section_rgb("Color5", palette) == "123456"
    assert section_rgb("Red", palette) == "FF0000"  # names ignore the palette
    assert section_rgb("Color5", palette[:10]) == DEFAULT_PALETTE["Color5"]  # past a short one
    assert section_rgb("Color5") == DEFAULT_PALETTE["Color5"]
    with pytest.raises(ValueError, match="not an RRGGBB"):
        section_rgb("Color1", ["x"] * 64)
    with pytest.raises(KeyError):
        section_rgb("Purple")
