"""``row_metrics`` against ``tests/captures/row_height.layout.tsv`` (CAPTURING.md §2 item 44).

The capture holds, per Normal font (both families, six sizes, four styles), the height Excel
gave each row of a workbook: a row in the Normal font, a row without an element, the sheet's
standard height, a number in each measured cell font, styled empty cells, two fonts in one row,
and one to six wrapped lines in the Normal font and in each regular cell font. Each row's tag
says what it held, and the test rebuilds the call from it.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

import pytest

from xlformat import Font, RowMetrics, row_metrics

LAYOUT = Path(__file__).parent / "captures" / "row_height.layout.tsv"
_FACE = re.compile(r"(Calibri|Aptos Narrow) (\d+)((?: bold)?)((?: italic)?)")


def _font(label: str) -> Font:
    m = _FACE.fullmatch(label)
    assert m, label
    return Font(m.group(1), float(m.group(2)), bool(m.group(3)), bool(m.group(4)))


def _call(normal: Font, tag: str) -> RowMetrics:
    """The ``row_metrics`` call a captured row's tag describes."""
    if tag in ("standard height", "number", "no element", "number again", "number at the end"):
        return row_metrics(normal)
    if m := re.fullmatch(r"wrapped (\d+) lines(?: (.+))?", tag):
        face = _font(m.group(2)) if m.group(2) else None
        return row_metrics(normal, wrapped=[(face, int(m.group(1)))])
    if m := re.fullmatch(r"(?:cell|empty styled|normal and) (.+)", tag):
        return row_metrics(normal, cell_fonts=[_font(m.group(1))])
    first, second = tag.split(" and ")
    return row_metrics(normal, cell_fonts=[_font(first), _font(second)])


def _rows() -> list[list[str]]:
    with LAYOUT.open(encoding="utf-8") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.reader(lines, dialect="excel-tab"))[1:]


def test_every_captured_row_height() -> None:
    rows = _rows()
    assert len(rows) == 48 * 148
    misses = [
        (normal, tag, height, got)
        for normal, _, tag, height in rows
        if (got := _call(_font(normal), tag).height) != float(height)
    ]
    assert not misses, f"{len(misses)} heights differ: {misses[:10]}"


def test_font_and_line_height() -> None:
    m = row_metrics(Font("Aptos Narrow", 16, bold=True))
    assert (m.height, m.font_height, m.line_height) == (22.0, 22.0, 23.0)
    assert row_metrics(normal_font="Calibri").font_height == 15.0


def test_unmeasured_faces_and_bad_lines() -> None:
    with pytest.raises(LookupError, match="no row heights for Calibri 13"):
        row_metrics(("Calibri", 13))
    with pytest.raises(LookupError, match="no row heights for Calibri 18"):
        row_metrics(cell_fonts=[("Calibri", 18)])
    with pytest.raises(LookupError):
        row_metrics("Arial")
    with pytest.raises(ValueError, match="at least one line"):
        row_metrics(wrapped=[(None, 0)])
