"""Excel for Windows' row heights across sizes (CAPTURING.md section 2 items 86 and 89).

Every row of ``tests/captures/windows/row_height_wide.layout.tsv.gz`` and
``row_height_ppem.layout.tsv.gz`` must be the row model's height, and every face it fits must be
in the package's table (``src/xlformat/_rows_windows.py``, whose other faces
``capture/windows_rows.py`` fits from item 69's capture).
"""

from __future__ import annotations

import pytest

import xlformat
from capture import rows_wide
from xlformat._rows_windows import WINDOWS_ROWS


@pytest.mark.parametrize(
    ("group", "faces"), [("rows_wide", 964), ("rows_ppem", 1108), ("rows_more", 8544)]
)
def test_every_row_is_the_model_and_in_the_table(group: str, faces: int) -> None:
    # a fitted row stops at Excel's largest, 409.5 points (rows_more: 55 four-line rows)
    table, problems = rows_wide.fit(group)
    assert problems == []
    assert len(table) == faces
    for (family, style, size), hd in table.items():
        assert WINDOWS_ROWS[family.lower()][style][size] == hd, (family, style, size)


def test_a_face_name_matches_ignoring_case() -> None:
    # CAPTURING.md section 2 item 91: "arial" draws as "Arial" at every ppem from 1 to 96
    table, _ = rows_wide.fit("rows_more")
    lower = {size: hd for (family, _, size), hd in table.items() if family == "arial"}
    assert len(lower) == 96
    by_ppem = {rows_wide.ppem(z): hd for z, hd in WINDOWS_ROWS["arial"]["regular"].items()}
    assert all(by_ppem[rows_wide.ppem(z)] == hd for z, hd in lower.items())
    assert xlformat.row_metrics("tahoma", platform="windows") == xlformat.row_metrics(
        "Tahoma", platform="windows"
    )


def test_each_ppem_has_one_height_and_every_size_to_72_answers() -> None:
    # CAPTURING.md section 2 item 89: per face and style one (H, D) at each ppem from 1 to 96
    for family, styles in WINDOWS_ROWS.items():
        for style, sizes in styles.items():
            by_ppem: dict[int, set[tuple[int, int]]] = {}
            for size, hd in sizes.items():
                by_ppem.setdefault(rows_wide.ppem(size), set()).add(hd)
            assert all(len(v) == 1 for v in by_ppem.values()), (family, style)
            assert set(by_ppem) >= set(range(1, 97)), (family, style)  # every family captured
    for size in (1, 1.5, 5.5, 40.5, 55.5, 72):
        font = xlformat.Font("Courier New", size, italic=True)
        h, d = WINDOWS_ROWS["courier new"]["italic"].get(size) or next(
            hd
            for z, hd in WINDOWS_ROWS["courier new"]["italic"].items()
            if rows_wide.ppem(z) == rows_wide.ppem(size)
        )
        metrics = xlformat.vertical_metrics(font, platform="windows")
        assert metrics.font_height * 96 / 72 == h
        assert metrics.baseline_from_top * 96 / 72 == h - d + 1
    with pytest.raises(LookupError, match="1 to 72 points"):
        xlformat.row_metrics(xlformat.Font("Arial", 72.5), platform="windows")


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
