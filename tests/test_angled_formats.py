"""Angled text is the width-less rendering (``tests/captures/angled_formats.layout.tsv``).

Excel fits each row of turned text to the text's width, so the captured rows show which text it
turns: rows of formats whose width-less texts are equally wide are equal, at every angle captured
(15 to 90 degrees and -45 and -90), whatever padding or fill the format holds, and a fill never
widens the text to the column (CAPTURING.md section 2 item 56).
"""

from __future__ import annotations

import ast
import collections
from pathlib import Path

from xlformat import format_value, text_width

LAYOUT = Path(__file__).parent / "captures" / "angled_formats.layout.tsv"


def _rows() -> list[tuple[int, object, str, float]]:
    out = []
    with LAYOUT.open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "angle\t")):
                continue
            angle, value, fmt, height = line.rstrip("\n").split("\t")
            out.append((int(angle), ast.literal_eval(value), fmt, float(height)))
    return out


def test_rows_follow_the_width_less_text() -> None:
    rows = _rows()
    assert len(rows) == 128
    heights: dict[tuple[int, float], set[float]] = collections.defaultdict(set)
    for angle, value, fmt, height in rows:
        heights[(angle, text_width(format_value(value, fmt, platform="mac"), platform="mac"))].add(
            height
        )
    assert all(len(h) == 1 for h in heights.values()), heights


def test_vertical_text_is_the_text_width_plus_seven() -> None:
    """Text turned 90 or -90 degrees fits a row as tall as its width plus 7 points."""
    for angle, value, fmt, height in _rows():
        if abs(angle) == 90:
            assert (
                height == text_width(format_value(value, fmt, platform="mac"), platform="mac") + 7
            ), (angle, value, fmt)


def test_negative_angles_fit_as_positive_ones() -> None:
    by = {(angle, repr(value), fmt): height for angle, value, fmt, height in _rows()}
    for (angle, value, fmt), height in by.items():
        if angle < 0:
            assert by[(-angle, value, fmt)] == height
