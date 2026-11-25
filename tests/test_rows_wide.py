"""Excel for Windows' row heights across sizes (CAPTURING.md section 2 item 86).

Every row of ``tests/captures/windows/row_height_wide.layout.tsv.gz`` must be the row model's
height, and every face it fits must be in the package's table (``src/xlformat/_rows_windows.py``,
whose other faces ``capture/windows_rows.py`` fits from item 69's capture).
"""

from __future__ import annotations

import pytest

import xlformat
from capture import rows_wide
from xlformat._rows_windows import WINDOWS_ROWS


def test_every_row_is_the_model_and_in_the_table() -> None:
    table, problems = rows_wide.fit()
    assert problems == []
    assert len(table) == 964
    for (family, style, size), hd in table.items():
        assert WINDOWS_ROWS[family.lower()][style][size] == hd, (family, style, size)


@pytest.mark.parametrize(
    "font",
    [
        xlformat.Font("Calibri", 72),
        xlformat.Font("Calibri", 8.5),
        xlformat.Font("Courier New", 6, bold=True),
        xlformat.Font("Times New Roman", 40, italic=True),
    ],
    ids=lambda f: f"{f.name} {f.size:g} {f.style}",
)
def test_row_metrics_answers_the_captured_sizes(font: xlformat.Font) -> None:
    h, _d = WINDOWS_ROWS[font.name.lower()][font.style][font.size]
    got = xlformat.row_metrics(font, platform="windows").height
    assert got * 96 / 72 == h  # a one-line row in its own Normal font is H
