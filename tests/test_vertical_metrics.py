"""``vertical_metrics`` against ``tests/captures/vertical.layout.tsv`` (CAPTURING.md items 47, 55).

The capture holds, per Normal font (Calibri 11, Aptos Narrow 11, Arial 10), each row's height
and, in image pixels from the row's top at 2 a point, the top and the bottom of the ink of "H"
per line (the bottom being the baseline) and the bottom of "g". Each row's tag names its
alignment and face; the test rebuilds every baseline from the face's metrics and the row's
height, to half a point, and the ink exactly.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

import pytest

from xlformat import Font, InkMetrics, VerticalMetrics, ink_metrics, vertical_metrics

LAYOUT = Path(__file__).parent / "captures" / "vertical.layout.tsv"
PIXELS_PER_POINT = 2
# The file's distances run from each row's top as computed from the view's first marker band, which
# starts below the column header's separator; every other band starts on its own top gridline, 2 px
# higher (band_shift -2). So each row's own top is 2 px above the file's origin (CAPTURING.md
# section 2 item 54).
FRAME_PX = 2
# Rows the reader measured only as the last row of a view, where the sheet-tab bar at the window's
# bottom edge cuts them: its pixels read as an extra or merged ink line (a crop of "center 2F
# Aptos Narrow 16" under Arial 10 shows one clean "Hxg" above the bar). The capture script now
# starts each view at the previous view's last row, so a new run measures them clear of the edge.
SCREEN_EDGE_MISREADS = {
    ("Calibri 11", "center 2F Arial 14"),
    ("Aptos Narrow 11", "center 2F Arial 14"),
    ("Arial 10", "center 2F Aptos Narrow 11"),
    ("Arial 10", "center 2F Aptos Narrow 16"),
}
_FAMILIES = "Calibri|Aptos Narrow|Arial|Times New Roman|Courier New|Georgia"
_FACE = re.compile(rf"({_FAMILIES}) (\d+)((?: bold)?)((?: italic)?)$")


def _font(label: str) -> Font:
    m = _FACE.search(label)
    assert m, label
    return Font(m.group(1), float(m.group(2)), bool(m.group(3)), bool(m.group(4)))


def _baselines(
    normal: VerticalMetrics, face: VerticalMetrics, tag: str, height: float
) -> list[float]:
    """Each line's baseline below the row's top, as the metrics place it."""
    kind = tag.split()[0]
    lines = int(tag.split()[1]) if kind == "wrapped" else 1
    # wrapped text in a face with a smaller line inset sits higher by the difference
    lift = max(0.0, normal.line_inset - face.line_inset) if kind == "wrapped" else 0.0
    # face is measured with the workbook's Normal font passed, so its baseline_from_bottom
    # already holds the Normal font's floor
    last = height - face.baseline_from_bottom - lift
    bottom = [last - (lines - 1 - i) * face.line_pitch for i in range(lines)]
    if kind == "top":
        return [face.baseline_from_top]
    if kind == "center":
        return [(face.baseline_from_top + bottom[0]) / 2]
    return bottom


def _rows() -> list[list[str]]:
    with LAYOUT.open(encoding="utf-8") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.reader(lines, dialect="excel-tab"))[1:]


def test_every_captured_placement() -> None:
    rows = _rows()
    assert len(rows) == 1080
    misses = []
    for normal_label, tag, _, height, _, _, h_ink, g_ink in rows:
        if (normal_label, tag) in SCREEN_EDGE_MISREADS:
            continue
        normal = vertical_metrics(_font(normal_label))
        face = vertical_metrics(_font(tag), normal_font=_font(normal_label))
        ink = [tuple(int(v) for v in line.split("-")) for line in h_ink.split()]
        tails = [int(v) for v in g_ink.split()]
        got = [(bottom + 1 + FRAME_PX) / PIXELS_PER_POINT for _, bottom in ink]
        want = _baselines(normal, face, tag, float(height))
        if len(got) != len(want) or any(abs(a - b) > 0.5 for a, b in zip(got, want, strict=False)):
            misses.append((normal_label, tag, got, want))
        for (top, bottom), tail in zip(ink, tails, strict=True):
            cap = (bottom + 1 - top) / PIXELS_PER_POINT
            descender = (tail - bottom) / PIXELS_PER_POINT
            # Arial 11 and 16 draw the descender half a point either way by the baseline's pixel
            # position (the same in all three workbooks, by alignment)
            slack = (
                0.5
                if (face_key := _FACE.search(tag))
                and face_key.group(1) == "Arial"
                and face_key.group(2) in ("11", "16")
                else 0.0
            )
            ink = ink_metrics(_font(tag))
            if cap != ink.cap_height or abs(descender - ink.descender) > slack:
                misses.append((normal_label, tag, (cap, descender)))
    assert not misses, f"{len(misses)} rows differ: {misses[:10]}"


def test_fields() -> None:
    m = vertical_metrics()
    assert m == VerticalMetrics(
        baseline_from_top=12.0,
        baseline_from_bottom=3.0,
        font_height=15.0,
        line_height=16.0,
        line_pitch=15.0,
        line_inset=1.0,
    )
    assert vertical_metrics(Font("Calibri", 16, italic=True)) == vertical_metrics(("Calibri", 16))
    assert ink_metrics() == InkMetrics(cap_height=7.0, descender=2.0)


def test_the_ink_fields_of_vertical_metrics_are_deprecated() -> None:
    m = vertical_metrics(("Arial", 14))
    with pytest.warns(DeprecationWarning, match="ink_metrics"):
        assert m.cap_height == ink_metrics(("Arial", 14)).cap_height
    with pytest.warns(DeprecationWarning, match="ink_metrics"):
        assert m.descender == ink_metrics(("Arial", 14)).descender
    # placement covers every supported size; ink only what was measured
    assert vertical_metrics(("Times New Roman", 13)).baseline_from_top == 14.0
    with pytest.raises(LookupError, match="no ink measured for Times New Roman 13"):
        ink_metrics(("Times New Roman", 13))


def test_unmeasured_faces() -> None:
    with pytest.raises(LookupError, match="no ink measured for Calibri 22"):
        ink_metrics(("Calibri", 22))
    with pytest.raises(LookupError, match="no row heights for Calibri 73"):
        vertical_metrics(("Calibri", 73))
    with pytest.raises(LookupError, match="Comic Sans"):
        vertical_metrics("Comic Sans")


def test_the_normal_font_floors_a_bottom_aligned_baseline() -> None:
    # Calibri 9 and Arial 9 to 11 sit 2 points up of their own and 3 under a Calibri 11 Normal
    # font, 2 under Arial 10 (vertical.layout.tsv, CAPTURING.md section 2 item 55)
    assert vertical_metrics(("Calibri", 9)).baseline_from_bottom == 2.0
    assert vertical_metrics(("Calibri", 9), normal_font=("Calibri", 11)).baseline_from_bottom == 3.0
    assert vertical_metrics(("Calibri", 9), normal_font=("Arial", 10)).baseline_from_bottom == 2.0
    assert (
        vertical_metrics(("Arial", 11), normal_font=("Aptos Narrow", 11)).baseline_from_bottom
        == 3.0
    )
    # a face that sits higher than the Normal font keeps its own
    assert vertical_metrics(("Calibri", 24), normal_font=("Arial", 10)).baseline_from_bottom == 6.0
    # the default is the cell's own font
    assert vertical_metrics(("Calibri", 11)) == vertical_metrics(
        ("Calibri", 11), normal_font=("Calibri", 11)
    )
    with pytest.raises(LookupError, match="Comic Sans"):
        vertical_metrics(("Calibri", 11), normal_font="Comic Sans")
