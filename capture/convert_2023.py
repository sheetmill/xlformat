"""Convert the 2023 clipboard capture into the current capture schema, gzipped.

Source: ``batch_test_cases.tsv`` from the 2023 ``pyxlformat`` project (path given on the command
line; it lives outside this repository). Output: ``tests/captures/capture2023.tsv.gz``, a
full-tier-only suite tagged ``capture2023``.
"""

from __future__ import annotations

import csv
import gzip
import sys
from pathlib import Path

from capture.schema import COLUMNS, Row, write_rows

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests" / "captures" / "capture2023.tsv.gz"


def main(source: str) -> None:
    """Read the 2023 TSV and write the gzipped suite."""
    rows: list[Row] = []
    with Path(source).open(newline="") as f:
        for expected, value_info, format_repr in csv.reader(f, dialect="excel-tab"):
            value_type, value_repr = value_info.split(" ", 1)
            rows.append(
                Row(
                    suite="capture2023",
                    tag="legacy",
                    value_type=value_type,
                    value_repr=value_repr,
                    format=format_repr,
                    width="",
                    actual_width="",
                    date1904="0",
                    expected_text=expected,
                    expected_color="",
                    tier="full",
                )
            )
    header = {
        "suite": "capture2023",
        "excel_version": "Excel for Mac, 2023 build (not recorded)",
        "standard_font": "not recorded",
        "workbook_font": "Calibri 11 (openpyxl default)",
        "locale": "en_US",
        "captured": "2023 (Jan to Sep)",
        "method": "clipboard: cells copied in Excel, pbpaste; width-less",
        "columns": " ".join(COLUMNS),
    }
    with gzip.open(OUT, "wt", newline="", encoding="utf-8") as g:
        write_rows(g, header, rows)
    print(f"wrote {OUT} rows={len(rows)} bytes={OUT.stat().st_size}")


if __name__ == "__main__":
    main(sys.argv[1])
