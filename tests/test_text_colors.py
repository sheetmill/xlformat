"""Text and number colors as the screen shows them (``tests/captures/rules.layout.tsv``).

The rows hold "Total" and 12 in Calibri 11 under formats of one to four sections, read from window
captures at zoom 100: the mean color of the text's colored ink, or ``black`` (CAPTURING.md
section 2 item 60). Excel's ``display format`` read-back misreports colored text, so these rows,
not the ``colors`` suite's, decide the color of text.
"""

from __future__ import annotations

import ast
from pathlib import Path

from xlformat import format_cell

LAYOUT = Path(__file__).parent / "captures" / "rules.layout.tsv"
_HUES = {"Blue": (0, 0, 255), "Red": (255, 0, 0), "Magenta": (255, 0, 255)}


def _nearest(rgb: str) -> str:
    r, g, b = (int(v) for v in rgb.split())
    return min(
        _HUES, key=lambda n: sum((a - c) ** 2 for a, c in zip((r, g, b), _HUES[n], strict=True))
    )


def test_text_takes_a_color_only_from_a_text_section_or_a_lone_general() -> None:
    rows = []
    with LAYOUT.open(encoding="utf-8") as f:
        for line in f:
            fields = line.rstrip("\n").split("\t")
            if line.startswith("#") or fields[0] != "100" or not fields[1].startswith(("'", "12 ")):
                continue
            value_text, _, fmt = fields[1].partition(" ")
            rows.append((ast.literal_eval(value_text), fmt, fields[5]))
    assert len(rows) == 18
    for value, fmt, seen in rows:
        color = format_cell(value, fmt).color
        assert (color or "black") == (seen if seen == "black" else _nearest(seen)), (value, fmt)
