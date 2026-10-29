"""What Excel's repair does to a refused format, against ``tests/captures/repair.layout.tsv``.

Captured by ``capture/repair.py`` with the repair dialogs accepted by hand (CAPTURING.md section 2
item 33): every cell whose format Excel refuses is reset to General and shows General in its
column, and every other cell keeps its format. So a reader that catches ``FormatError`` and
renders the cell as General shows what Excel shows after the repair.
"""

from __future__ import annotations

import csv
from ast import literal_eval
from pathlib import Path

import pytest

from xlformat import FormatError, NumberFormat, file_column_metrics, format_value

LAYOUT = Path(__file__).parent / "captures" / "repair.layout.tsv"
WIDTH = file_column_metrics().width  # the workbooks set no column widths: 8 units


def _rows() -> list[dict[str, str]]:
    with LAYOUT.open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader([ln for ln in f if not ln.startswith("#")], dialect="excel-tab"))


@pytest.mark.excel
def test_repair_resets_refused_cells_to_general_and_keeps_the_rest() -> None:
    rows = _rows()
    refused = [r for r in rows if r["tag"] == "refused"]
    assert len(refused) == 120
    failures = []
    for row in rows:
        fmt, value = literal_eval(row["format"]), literal_eval(row["value_repr"])
        if row["tag"] == "refused":
            with pytest.raises(FormatError):
                NumberFormat.parse(fmt)
            expected_format, shown = "General", format_value(value, "General", column_width=WIDTH)
        else:
            expected_format, shown = fmt, format_value(value, fmt, column_width=WIDTH)
        if row["format_after"] != expected_format or row["text"] != shown:
            failures.append(f"{row['tag']} {fmt!r} {value!r}: {row['text']!r} {shown!r}")
    assert not failures, "\n".join(failures)
