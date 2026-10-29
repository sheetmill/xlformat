"""Excel's default horizontal alignment against ``tests/captures/alignment.layout.tsv``.

The rows come from Excel's PDF export (``capture/alignment.py``, CAPTURING.md section 2 item 31):
each case in a column 40 file units wide, its alignment read from where the printed text sits
between the gridlines. The text is compared at that width with spaces removed (the PDF draws no
pad spaces). Rows without text (``;;;`` and the like) and rows whose text fills the column (a
``*`` fill, a ``#`` run) have no alignment to read.
"""

from __future__ import annotations

import csv
from ast import literal_eval
from pathlib import Path

import pytest

from xlformat import file_column_metrics, format_cell

LAYOUT = Path(__file__).parent / "captures" / "alignment.layout.tsv"
COLUMN_WIDTH = file_column_metrics(40).width  # capture/alignment.py: COLUMN_WIDTH


def _rows() -> list[dict[str, str]]:
    with LAYOUT.open(encoding="utf-8", newline="") as f:
        lines = [line for line in f if not line.startswith("#")]
    return list(csv.DictReader(lines, dialect="excel-tab"))


@pytest.mark.excel
def test_alignment_matches_the_pdf_export() -> None:
    rows = _rows()
    assert rows, "no alignment rows"
    failures = []
    counts = {"aligned": 0, "fills the column": 0, "no text": 0, "errors (step 8)": 0}
    for row in rows:
        if row["value_type"] == "error":
            counts["errors (step 8)"] += 1
            continue
        cell = format_cell(
            literal_eval(row["value_repr"]), literal_eval(row["format"]), column_width=COLUMN_WIDTH
        )
        if cell.text.replace(" ", "") != row["text"].replace(" ", ""):
            failures.append(f"{row['format']} {row['value_repr']}: text {cell.text!r}")
            continue
        if not row["align"]:
            counts["no text"] += 1
        elif cell.overflow or cell.fill is not None:
            counts["fills the column"] += 1
        else:
            counts["aligned"] += 1
            if cell.align != row["align"]:
                failures.append(
                    f"{row['format']} {row['value_type']} {row['value_repr']}: expected "
                    f"{row['align']}, got {cell.align}"
                )
    print(f"\nalignment: {len(rows)} rows, " + ", ".join(f"{n} {k}" for k, n in counts.items()))
    assert not failures, f"{len(failures)} rows differ:\n" + "\n".join(failures)
