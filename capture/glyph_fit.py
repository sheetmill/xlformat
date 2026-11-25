"""Fit glyph advances from the per-size autofit measure suites (``tests/test_glyph_tables.py``).

Reads a ``measure*.measure.tsv`` suite: Excel for Mac's in ``tests/captures/`` (points), Excel
for Windows' at 96 dpi in ``tests/captures/windows/`` (pixels: points times 96 / 72; CAPTURING.md
section 2 item 65). Each holds every printable ASCII character in runs of 1, 2, 5 and 10, digit
pairs, month and day names, AM, PM, E+07, ####, mixed strings, autofitted one per column, with
the column width read back in points (CAPTURING.md section 2 items 24, 26 and 30). The header
names the workbook's Normal font (family and size); a ``cell_font`` column names the styled face
the cells carried, else the cells are the Normal font. The package's advances come from
``capture/face_fit.py`` (the Mac) and ``capture/face_fit_windows.py`` (Windows); these suites
check them.

Model: the autofit width is ``floor(sum of glyph advances) + PAD``; PAD comes out of the digit
runs (``2 * width(5 digits) - width(10 digits)``): the column's padding plus 2 for an upright face
(``xlformat._measure.column_padding``), more for an italic one, whose slant reaches half the
excess past its advances on each side (the overhang). The advance of a glyph is
``(width of its ten-run - PAD) / 10``: a whole advance, plus a tenth per unit of autofit bonus a
wide glyph's ten-run earns.
"""

from __future__ import annotations

import ast
import math
import re
from pathlib import Path

from xlformat import Font
from xlformat._platform import Platform

_LABEL = re.compile(r"(.+?) (\d+(?:\.\d+)?)((?: bold)?)((?: italic)?)$")
# the platform being fitted: the Mac (points) or Windows at 96 dpi (pixels), set by the test
PLATFORM = Platform.MAC


def _font(label: str) -> Font:
    match = _LABEL.match(label)
    if not match:
        msg = f"not a font label: {label!r}"
        raise ValueError(msg)
    return Font(match[1], float(match[2]), bool(match[3]), bool(match[4]))


def load(path: Path) -> tuple[Font, Font, dict[str, float]]:
    """The cell face, the Normal font (from the header) and text to autofit points."""
    lines = path.read_text().splitlines()
    match = re.search(r"font: ([^;]+)", lines[0])
    if not match:
        msg = f"{path.name}: header does not name the font"
        raise ValueError(msg)
    normal = _font(match[1].strip())
    faces: set[str] = set()
    out: dict[str, float] = {}
    for line in lines[1:]:
        if line.startswith(("#", "value_type")):
            continue
        fields = line.split("\t")
        device = float(fields[4]) * PLATFORM.scale
        if device != round(device):
            msg = f"{path.name}: {fields[4]} points is not a whole device unit"
            raise ValueError(msg)
        out[ast.literal_eval(fields[1])] = device
        if len(fields) > 5 and fields[5]:
            faces.add(fields[5])
    if len(faces) > 1:
        msg = f"{path.name}: more than one cell font {sorted(faces)}"
        raise ValueError(msg)
    face = _font(faces.pop()) if faces else normal
    return face, normal, out


def padding(measured: dict[str, float]) -> float:
    """Autofit padding in points, from the five- and ten-runs of every digit (must agree)."""
    pads = {2 * measured[d * 5] - measured[d * 10] for d in "0123456789"}
    if len(pads) != 1:
        msg = f"digit runs disagree on the padding: {sorted(pads)}"
        raise ValueError(msg)
    return pads.pop()


def fit(measured: dict[str, float]) -> tuple[dict[str, float], float, int, int, list[str]]:
    """Advances per glyph (space unset), padding, exact matches, off-by-one, worse strings."""
    pad = padding(measured)
    adv = {chr(c): (measured[chr(c) * 10] - pad) / 10 for c in range(33, 127)}
    exact = off = 0
    worse: list[str] = []
    for text, pts in measured.items():
        if " " in text:
            continue  # scored once the space is known
        diff = _predict(text, adv, pad) - pts
        if diff == 0:
            exact += 1
        elif abs(diff) == 1:
            off += 1
        else:
            worse.append(text)
    return adv, pad, exact, off, worse


def _predict(text: str, adv: dict[str, float], pad: float) -> float:
    return math.floor(sum(adv[c] for c in text) + 1e-9) + pad
