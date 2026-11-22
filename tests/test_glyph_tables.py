"""The shipped glyph tables are what the committed autofit captures give, on both platforms.

Every ``measure*.measure.tsv`` (Excel for Mac in ``tests/captures/``, Excel for Windows at 96 dpi
in ``tests/captures/windows/``) is fitted again as ``capture/glyph_fit.py`` fits it: each glyph's
advance from its ten-run less the autofit padding, the italic overhang from the padding. The
result must equal the table and overhang the package ships for that face. The space, which
autofit cannot resolve, is checked by the ``space_fill`` suites in ``tests/test_captures.py``.
"""

from __future__ import annotations

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
    face, _, measured = glyph_fit.load(path)
    adv, pad, *_ = glyph_fit.fit(measured)
    fonts, overhangs = _measure._REGISTRY[platform.name]
    shipped = fonts[face.name.lower()][1][face.style][int(face.size)]
    assert {c: a for c, a in shipped.items() if c != " "} == adv
    over = (pad - column_padding(adv["0"]) - 2) / 2
    assert overhangs[face.name.lower()].get(face.style, {}).get(int(face.size), 0) == over
