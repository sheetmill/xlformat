"""Every text capture holds no control character other than tab and newline.

The check reads each file's raw text, gzipped files decompressed, and splits it on ``\\n``
only: ``str.splitlines`` would also split on, and so hide, ``\\r``, ``\\x0b``, ``\\x0c``,
``\\x1c`` to ``\\x1e`` and ``\\x85``.
"""

from __future__ import annotations

import gzip
import re
from pathlib import Path

import pytest
from capture.schema import read_rows

CAPTURES = Path(__file__).parent / "captures"
TEXT_SUFFIXES = (".tsv", ".json")
FILES = sorted(
    p
    for p in CAPTURES.rglob("*")
    if p.is_file() and p.name.removesuffix(".gz").endswith(TEXT_SUFFIXES)
)

# C0 controls except tab and newline, DEL, and the C1 controls.
CONTROL = re.compile("[\x00-\x08\x0b-\x1f\x7f-\x9f]")

# Suites whose cells are built with CHAR(13): Excel displays the carriage returns, so each
# one is captured in the row's quoted expected_text field, and nowhere else.
CARRIAGE_RETURN_IN_TEXT = frozenset({"line_breaks.tsv.gz", "line_breaks_aptos_narrow.tsv.gz"})


def _text(path: Path) -> str:
    data = path.read_bytes()
    if path.suffix == ".gz":
        data = gzip.decompress(data)
    return data.decode("utf-8")


def test_every_capture_is_checked() -> None:
    assert len(FILES) > 800
    assert {p.name for p in FILES} >= CARRIAGE_RETURN_IN_TEXT
    assert not [p for p in CAPTURES.rglob("*") if p.is_file() and p not in FILES]


@pytest.mark.parametrize("path", FILES, ids=lambda p: str(p.relative_to(CAPTURES)))
def test_a_text_capture_holds_no_control_character(path: Path) -> None:
    allowed = "\r" if path.name in CARRIAGE_RETURN_IN_TEXT else ""
    found = [
        (number, repr(match.group()), line[:60])
        for number, line in enumerate(_text(path).split("\n"), 1)
        for match in CONTROL.finditer(line)
        if match.group() not in allowed
    ]
    assert not found, found


@pytest.mark.parametrize("name", sorted(CARRIAGE_RETURN_IN_TEXT))
def test_carriage_returns_occur_only_in_expected_text(name: str) -> None:
    path = CAPTURES / name
    _, rows = read_rows(path)
    in_text = sum(row.expected_text.count("\r") for row in rows)
    assert in_text > 0
    assert _text(path).count("\r") == in_text
