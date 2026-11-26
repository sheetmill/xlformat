"""``row_metrics`` against ``tests/captures/row_height.layout.tsv.gz`` (CAPTURING.md §2 item 44).

The capture holds, per Normal font (three families, nine sizes, four styles), the height Excel
gave each row of a workbook: a row in the Normal font, a row without an element, the sheet's
standard height, a number in each measured cell font, styled empty cells, two fonts in one row,
and one to six wrapped lines in the Normal font and in each regular cell font. Each row's tag
says what it held, and the test rebuilds the call from it.
"""

from __future__ import annotations

import csv
import gzip
import re
from functools import partial
from pathlib import Path

import pytest

from xlformat import Font, RowMetrics, row_metrics
from xlformat import row_metrics as _row_metrics

CAPTURES = Path(__file__).parent / "captures"
# Excel for Mac's capture, and Excel for Windows' at 96 dpi (CAPTURING.md section 2 item 69)
LAYOUTS = {
    "mac": CAPTURES / "row_height.layout.tsv.gz",
    "windows": CAPTURES / "windows" / "row_height.layout.tsv.gz",
}
_FACE = re.compile(r"(Calibri|Aptos Narrow|Arial) (\d+)((?: bold)?)((?: italic)?)")


def _font(label: str) -> Font:
    m = _FACE.fullmatch(label)
    assert m, label
    return Font(m.group(1), float(m.group(2)), bool(m.group(3)), bool(m.group(4)))


def _call(normal: Font, tag: str, platform: str = "mac") -> RowMetrics:
    """The ``row_metrics`` call a captured row's tag describes."""
    row_metrics = partial(_row_metrics, platform=platform)
    if tag in ("standard height", "number", "no element", "number again", "number at the end"):
        return row_metrics(normal)
    if m := re.fullmatch(r"wrapped (\d+) lines (.+) beside (.+)", tag):
        face = None if m.group(2) == "normal" else _font(m.group(2))
        return row_metrics(
            normal, cell_fonts=[_font(m.group(3))], wrapped=[(face, int(m.group(1)))]
        )
    if m := re.fullmatch(r"wrapped (\d+) lines(?: (.+))?", tag):
        face = _font(m.group(2)) if m.group(2) else None
        return row_metrics(normal, wrapped=[(face, int(m.group(1)))])
    if m := re.fullmatch(r"(?:cell|empty styled|normal and) (.+)", tag):
        return row_metrics(normal, cell_fonts=[_font(m.group(1))])
    first, second = tag.split(" and ")
    return row_metrics(normal, cell_fonts=[_font(first), _font(second)])


def _rows(platform: str = "mac") -> list[list[str]]:
    with gzip.open(LAYOUTS[platform], "rt", encoding="utf-8", newline="") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.reader(lines, dialect="excel-tab"))[1:]


@pytest.mark.parametrize("platform", sorted(LAYOUTS))
def test_every_captured_row_height(platform: str) -> None:
    rows = _rows(platform)
    per_font: dict[str, list[str]] = {}
    for normal, _, tag, _ in rows:
        per_font.setdefault(normal, []).append(tag)
    assert len(per_font) == 108  # three families, nine sizes, four styles
    assert len({tuple(tags) for tags in per_font.values()}) == 1  # the same rows in every workbook
    misses = [
        (normal, tag, height, got)
        for normal, _, tag, height in rows
        if (got := _call(_font(normal), tag, platform).height) != float(height)
    ]
    assert not misses, f"{len(misses)} heights differ: {misses[:10]}"


def test_font_and_line_height() -> None:
    m = row_metrics(Font("Aptos Narrow", 16, bold=True), platform="mac")
    assert (m.height, m.font_height, m.line_height) == (22.0, 22.0, 23.0)
    assert row_metrics(normal_font="Calibri", platform="mac").font_height == 15.0


def test_unmeasured_faces_and_bad_lines() -> None:
    with pytest.raises(LookupError, match="no row heights for Calibri 73"):
        row_metrics(("Calibri", 73), platform="mac")
    with pytest.raises(LookupError, match="no row heights for font 'Comic Sans'"):
        row_metrics(cell_fonts=[("Comic Sans", 11)], platform="mac")
    # sizes and faces between the tables follow the slope model (font_height.layout.tsv)
    assert row_metrics(("Calibri", 13), platform="mac").font_height == 17.0
    assert row_metrics(("Calibri", 12.5), platform="mac").font_height == 17.0
    assert row_metrics(("Times New Roman", 11), platform="mac").font_height == 14.0


def test_faces_of_two_families() -> None:
    # Arial 11 and Calibri 10 are both 14 points tall; Arial's descent is smaller, so the row
    # holds Arial's ascent and Calibri's descent: 15 points (row_height.layout.tsv.gz)
    assert row_metrics(("Arial", 11), cell_fonts=[("Calibri", 10)], platform="mac").height == 15.0
    assert row_metrics(("Calibri", 11), cell_fonts=[("Arial", 11)], platform="mac").height == 15.0
    assert row_metrics(("Arial", 10), platform="mac").font_height == 13.0
    with pytest.raises(ValueError, match="at least one line"):
        row_metrics(wrapped=[(None, 0)], platform="mac")
