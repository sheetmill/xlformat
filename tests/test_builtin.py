"""``builtin_format`` against ``tests/captures/builtin_formats.layout.tsv`` (CAPTURING.md §2 49).

Each row is a built-in id as Excel for Mac (en-US) reported it: its number format and the text
of seven values under it at a column width of 40.
"""

from __future__ import annotations

import csv
from ast import literal_eval
from pathlib import Path

import pytest

from xlformat import BUILTIN_FORMATS, builtin_format, format_value

CAPTURES = Path(__file__).parent / "captures"
# Excel for Mac's capture, and Excel for Windows' at 96 dpi (CAPTURING.md section 2 item 70)
LAYOUTS = {
    "mac": CAPTURES / "builtin_formats.layout.tsv",
    "windows": CAPTURES / "windows" / "builtin_formats.layout.tsv",
}
VALUES = (45292.5625, 1234.5678, -1234.5678, 0.25, 0, 1.5, "abc")


def _rows(platform: str = "mac") -> list[list[str]]:
    with LAYOUTS[platform].open(encoding="utf-8") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.reader(lines, dialect="excel-tab"))[1:]


@pytest.mark.parametrize("platform", sorted(LAYOUTS))
def test_every_id_and_its_texts(platform: str) -> None:
    rows = _rows(platform)
    assert [int(r[0]) for r in rows] == list(range(164))
    misses = []
    for ident, stored, _, *texts in rows:
        fmt = builtin_format(int(ident), platform=platform)
        if fmt != literal_eval(stored):
            misses.append((ident, fmt, literal_eval(stored)))
        for value, text in zip(VALUES, texts, strict=True):
            got = format_value(value, fmt, column_width=40, platform=platform)
            if got != literal_eval(text):
                misses.append((ident, value, got, literal_eval(text)))
    assert not misses, f"{len(misses)} differ: {misses[:10]}"


def test_typed_date_and_table() -> None:
    assert builtin_format(14, platform="mac") == "m/d/yy"
    assert format_value(45292, builtin_format(14, platform="mac"), platform="mac") == "1/1/24"
    assert len(BUILTIN_FORMATS) == 164
    with pytest.raises(TypeError):
        BUILTIN_FORMATS[14] = "x"  # type: ignore[index]


@pytest.mark.parametrize("ident", [-1, 164, 200])
def test_not_builtin(ident: int) -> None:
    with pytest.raises(ValueError, match="not a built-in number format id"):
        builtin_format(ident, platform="mac")
