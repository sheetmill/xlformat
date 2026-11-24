"""The glyph advances xlformat uses are what the committed autofit captures give, on both platforms.

Every ``measure*.measure.tsv`` (Excel for Mac in ``tests/captures/``, Excel for Windows at 96 dpi
in ``tests/captures/windows/``) is fitted again as ``capture/glyph_fit.py`` fits it: each glyph's
advance from its ten-run less the autofit padding, the italic overhang from the padding. On
Windows the result must equal the table and overhang the package ships for that face. On the
Mac it must equal the rule's whole advance plus a tenth per point of autofit bonus the ten-run
earns (``max(0, floor(W / D / 5) - 2)``, W the run's width, D the Normal font's digit), and the
overhang the package ships. The space, which autofit cannot resolve, is checked by the
``space_fill`` suites in ``tests/test_captures.py``.
"""

from __future__ import annotations

import math
from pathlib import Path

import pytest

from capture import glyph_fit
from xlformat import _measure
from xlformat._measure import column_padding
from xlformat._platform import Platform

CAPTURES = Path(__file__).parent / "captures"
FILES = [
    *((Platform.MAC, p) for p in sorted(CAPTURES.glob("measure*.measure.tsv"))),
    *((Platform.WINDOWS, p) for p in sorted((CAPTURES / "windows").glob("measure*.measure.tsv"))),
]


@pytest.mark.parametrize(
    ("platform", "path"),
    FILES,
    ids=[f"{p.name}/{path.name.split('.')[0]}" for p, path in FILES],
)
def test_shipped_table_is_the_fit(
    platform: Platform, path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(glyph_fit, "PLATFORM", platform)
    face, normal, measured = glyph_fit.load(path)
    adv, pad, *_ = glyph_fit.fit(measured)
    over = (pad - column_padding(adv["0"]) - 2) / 2
    if platform is Platform.WINDOWS:
        fonts, overhangs = _measure._FONTS_WINDOWS, _measure._OVERHANG_WINDOWS
        shipped = fonts[face.name.lower()][1][face.style][int(face.size)]
        assert {c: a for c, a in shipped.items() if c != " "} == adv
        assert overhangs[face.name.lower()].get(face.style, {}).get(int(face.size), 0) == over
        return
    whole = _measure._whole_advances(face.name, face.size, face.style)
    digit = _measure._whole_advances(normal.name, normal.size, normal.style)[""]
    expected = {c: whole[c] + max(0, math.floor(10 * whole[c] / digit / 5) - 2) / 10 for c in adv}
    assert {c: round(a, 9) for c, a in adv.items()} == expected
    assert _measure.overhang(face) == over
