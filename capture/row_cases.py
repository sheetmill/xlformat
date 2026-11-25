"""The rows of a row-height workbook: cells, fonts and wrapped lines (no openpyxl needed).

Shared by ``capture/row_heights.py``, which writes the workbooks, and ``capture/rows_wide.py``,
whose fit reads the rows back without building them.
"""

from __future__ import annotations

from dataclasses import dataclass

from capture.cases import CellFont


@dataclass(frozen=True, slots=True)
class Cell:
    """One cell of a row: column (1 = A), value (None: empty but styled), font, wrap."""

    column: int
    value: object
    font: CellFont | None = None
    wrap: bool = False


@dataclass(frozen=True, slots=True)
class RowCase:
    """One row of every workbook; ``cells`` empty means the row has no element."""

    tag: str
    cells: tuple[Cell, ...]


def _lines(n: int) -> str:
    return "\n".join(["x"] * n)  # one short glyph a line: no line wraps at any width
