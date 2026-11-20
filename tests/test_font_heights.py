"""``row_metrics`` against ``font_height.layout.tsv`` and ``wrapped_height.layout.tsv``.

The first holds, for six faces in workbooks whose Normal style is the face at 1 point, the row
Excel fits to a number in the face at every size from 1 to 72 in half steps, and to the face at
9, 11, 14 and 20 beside Calibri at 9 to 24. The second holds one to three wrapped lines of each
face at 9, 11, 14 and 20 under a Calibri 20 and a Calibri 9 Normal font (CAPTURING.md section 2
item 57).
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from xlformat import Font, row_metrics

CAPTURES = Path(__file__).parent / "captures"
_FACE = re.compile(r"(.+?) (\d+(?:\.5)?)")


def _rows(name: str) -> list[list[str]]:
    with (CAPTURES / name).open(encoding="utf-8") as f:
        return [line.rstrip("\n").split("\t") for line in f if not line.startswith("#")][1:]


def test_every_face_at_every_size_and_beside_calibri() -> None:
    rows = _rows("font_height.layout.tsv")
    assert len(rows) == 6 * 179
    misses = []
    for face, tag, height in rows:
        if " beside " in tag:
            own, other = tag.split(" beside ")
            m, n = _FACE.fullmatch(own), _FACE.fullmatch(other)
            assert m, tag
            assert n, tag
            got = row_metrics(
                (face, 1.0), cell_fonts=[(m[1], float(m[2])), (n[1], float(n[2]))]
            ).height
        else:
            m = _FACE.fullmatch(tag)
            assert m, tag
            got = row_metrics((face, 1.0), cell_fonts=[(m[1], float(m[2]))]).height
        if got != float(height):
            misses.append((face, tag, height, got))
    assert not misses, f"{len(misses)} rows differ: {misses[:10]}"


def test_wrapped_lines_of_every_face() -> None:
    rows = _rows("wrapped_height.layout.tsv")
    assert len(rows) == 144
    misses = []
    for normal, tag, height in rows:
        m = re.fullmatch(r"(.+?) (\d+) wrapped (\d)", tag)
        assert m, tag
        n = _FACE.fullmatch(normal)
        assert n, normal
        got = row_metrics(
            ("Calibri", float(n[2])), wrapped=[((m[1], float(m[2])), int(m[3]))]
        ).height
        if got != float(height):
            misses.append((normal, tag, height, got))
    assert not misses, f"{len(misses)} rows differ: {misses[:10]}"


def test_every_underlined_row() -> None:
    # underline_height.layout.tsv: every face at every whole size from 6 to 72 under each
    # underline style; Courier New's accounting rows and Times New Roman's at 37 to 40 points
    # follow no known rule and raise
    rows = _rows("underline_height.layout.tsv")
    assert len(rows) == 2814
    misses, raised = [], 0
    for tag, height in rows:
        m = re.fullmatch(r"(.+?) (\d+) (\w+) (\d)", tag)
        assert m, tag
        face, size, underline, lines = m[1], float(m[2]), m[3], int(m[4])
        font = Font(face, size, underline=None if underline == "none" else underline)
        try:
            if lines == 1:
                got = row_metrics(("Calibri", 1.0), cell_fonts=[font]).height
            else:
                got = row_metrics(("Calibri", 1.0), wrapped=[(font, lines)]).height
        except LookupError:
            assert "Accounting" in underline, tag
            assert face == "Courier New" or (face == "Times New Roman" and 37 <= size <= 40), tag
            raised += 1
            continue
        if got != float(height):
            misses.append((tag, height, got))
    assert not misses, f"{len(misses)} rows differ: {misses[:10]}"
    assert raised == 67 * 3 + 4 * 3


def test_underline_names() -> None:
    with pytest.raises(ValueError, match="underline must be one of"):
        Font("Calibri", 11, underline="accounting")
