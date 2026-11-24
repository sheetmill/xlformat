"""Text insets against every row of the inset captures (CAPTURING.md section 2 item 81).

Each row holds four cells: left- and right-aligned under a single underline, whose ends are the
text's, then the same under a single accounting underline, whose text edges are the H's ink less
the bearings the plain cells show. In device pixels: the left inset from the left gridline's first
pixel column (column B's is under column A's fill, which ends on its last), the right one from the
text's last pixel column to the right gridline's first. A gridline is a pixel wide on Windows and
two on the Mac's 2x display.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from xlformat import ColumnMetrics, column_metrics, text_insets

CAPTURES = Path(__file__).parent / "captures"
LAYOUTS = {
    "windows": CAPTURES / "windows" / "screen_insets.layout.tsv",
    "mac": CAPTURES / "insets.layout.tsv",
}
# the Mac's first face row of each workbook: its right accounting cell's ink reads past its line
# (item 81); the line agrees
MAC_INK_MISREADS = {"Calibri 9", "Calibri 40"}


def _rows(path: Path) -> list[list[str]]:
    lines = path.read_text(encoding="utf-8").splitlines()
    return [x.split("\t") for x in lines if x and not x.startswith("#")][1:]


@pytest.mark.parametrize("platform", ["windows", "mac"])
def test_every_captured_row_matches_text_insets(platform: str) -> None:
    px = 4 / 3 if platform == "windows" else 2.0  # device pixels a point at zoom 100
    width = 1 if platform == "windows" else 2  # a gridline's device pixels
    checked = 0
    for normal, zoom, tag, _rh, _ppp, *cells in _rows(LAYOUTS[platform]):
        face, size = tag.rsplit(" ", 1)
        z = int(zoom)
        got = text_insets((face, float(size)), normal_font=_font(normal), zoom=z, platform=platform)
        c = [[int(v) for v in cell.split()] for cell in cells]
        gl1, ls1, il1 = c[0][0], c[0][2], c[0][4]
        gr2, le2, ih2 = c[1][1], c[1][3], c[1][5]
        gl3, gr3, ls3, le3, il3 = c[2][0], c[2][1], c[2][2], c[2][3], c[2][4]
        gr4, ih4 = c[3][1], c[3][5]
        unit = px * z / 100
        assert ls1 - (gl1 - width + 1) == pytest.approx(got.left * unit), (platform, zoom, tag)
        assert gr2 - le2 - 1 == pytest.approx(got.right * unit), (platform, normal, zoom, tag)
        bearing_left, bearing_right = il1 - ls1, le2 - ih2
        assert il3 - bearing_left - gl3 == pytest.approx(got.accounting_left * unit)
        assert ls3 - gl3 == pytest.approx(got.accounting_left * unit)  # the line starts there
        assert gr3 - le3 - 1 == pytest.approx(got.accounting_right * unit)
        if not (platform == "mac" and tag in MAC_INK_MISREADS):
            assert gr4 - (ih4 + bearing_right) - 1 == pytest.approx(got.accounting_right * unit)
        checked += 1
    assert checked == 444


def _font(label: str) -> tuple[str, float]:
    name, size = label.rsplit(" ", 1)
    return name, float(size)


@pytest.mark.parametrize("platform", ["windows", "mac"])
def test_the_column_insets_are_the_normal_fonts(platform: str) -> None:
    for font in (("Calibri", 11), ("Arial", 11), ("Aptos Narrow", 12), ("Calibri", 24)):
        m: ColumnMetrics = column_metrics(8.43, font, platform=platform)
        insets = text_insets(font, platform=platform)
        assert (m.left_inset, m.right_inset) == (insets.left, insets.right)


def test_text_insets_raises_outside_its_captures() -> None:
    with pytest.raises(LookupError, match="no text insets"):
        text_insets(("Georgia", 11), platform="windows")
    with pytest.raises(LookupError):
        text_insets(("Geneva", 11), platform="mac")
    with pytest.raises(LookupError, match="under one screen unit"):
        text_insets(("Calibri", 1), zoom=10, platform="windows")


# the faces capture (CAPTURING.md section 2 item 83): Verdana 36 italic at zoom 150 sits at screen
# size 54, whose overhang is scaled from 56's (11 screen points) where Excel shows 10
FACES_DEVIATIONS = {("Verdana 36 italic", 150)}


def test_the_faces_capture_matches_text_insets() -> None:
    from xlformat import Font

    checked = 0
    for _normal, zoom, tag, _rh, _ppp, *cells in _rows(CAPTURES / "insets_faces.layout.tsv"):
        words = tag.split(" ")
        italic = words[-1] == "italic"
        if italic:
            words = words[:-1]
        z = int(zoom)
        font = Font(" ".join(words[:-1]), float(words[-1]), italic=italic)
        got = text_insets(font, zoom=z, platform="mac")
        c = [[int(v) for v in cell.split()] for cell in cells]
        unit = 2 * z / 100  # device pixels a sheet point
        clear = (got.left - 100 / z) * unit  # the accounting line ends at the clear space
        assert c[0][2] - (c[0][0] - 1) == pytest.approx(got.left * unit), (tag, z)
        assert c[2][2] - c[2][0] == pytest.approx(got.accounting_left * unit), (tag, z)
        assert c[2][1] - c[2][3] - 1 == pytest.approx(clear), (tag, z)
        if (tag, z) not in FACES_DEVIATIONS:
            assert c[1][1] - c[1][3] - 1 == pytest.approx(got.right * unit), (tag, z)
            bearing = c[1][3] - c[1][5]
            if tag != "Times New Roman 11":  # the first face row: item 81's ink misread
                text_end = c[3][5] + bearing
                assert c[3][1] - text_end - 1 == pytest.approx(got.accounting_right * unit)
        checked += 1
    assert checked == 30
