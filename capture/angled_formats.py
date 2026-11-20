"""Padding and fills in angled text: the rows Excel fits.

    uv run python -m capture.angled_formats build   # the workbook, no Excel
    uv run python -m capture.angled_formats run     # open it, read the row heights, close it

One workbook, Normal font Calibri 11, one case per row in column A at the default width: a value
under a format with padding (``_)``), a ``*`` fill or neither, turned 15, 30, 45, 60, 75 or 90
degrees or -45 or -90 (stored as 90 plus the angle's size, ECMA-376), Calibri 11,
rows fitted by Excel on opening. A fitted row of angled text grows with the text's width
(``W·|sin θ| + S·|cos θ| + 7``), so rows of the same value under formats with and without padding
or a fill differ by what Excel adds to the turned text: a fill expanded to the column's width
would add tens of points, padding a few. The script reads each row's height through
``capture/row_heights.applescript``, closes the workbook without saving and restores alerts. No
clipboard. Writes ``tests/captures/angled_formats.layout.tsv``.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

from capture.build import WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript_or_exit

SCRIPT = Path(__file__).with_name("row_heights.applescript")
PATH = (WORKBOOK_DIR / "angled_formats.xlsx").resolve()
OUTPUT = Path(__file__).resolve().parents[1] / "tests" / "captures" / "angled_formats.layout.tsv"
ANGLES = (15, 30, 45, 60, 75, 90, -45, -90)
# (value, format): each group holds a plain control first
CASES: tuple[tuple[object, str], ...] = (
    ("Total", "@"),
    ("Total", "_(@_)"),
    ("Total", "@* "),
    ("Total", "* @"),
    ("Total", "@*-"),
    (1234.5, "0.0"),
    (1234.5, "0.0_)"),
    (1234.5, "_(0.0_)"),
    (1234.5, "0.0* "),
    (1234.5, "* 0.0"),
    (1234.5, "0.0*-"),
    (1234.5, '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'),
    (12345678901, "0"),
    (12345678901, "0_)"),
    (12345678901, "* 0"),
    (12345678901, "0*-"),
)


def rows() -> list[tuple[int, object, str]]:
    return [(angle, value, fmt) for angle in ANGLES for value, fmt in CASES]


def build() -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment

    wb = Workbook()
    ws = wb.active
    for r, (angle, value, fmt) in enumerate(rows(), 1):
        cell = ws.cell(row=r, column=1, value=value)
        if isinstance(value, str):
            cell.data_type = "s"
        cell.number_format = fmt
        cell.alignment = Alignment(text_rotation=angle if angle >= 0 else 90 - angle)
    PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(PATH)
    return PATH


def run() -> None:
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    build()
    n = len(rows())
    proc = osascript_or_exit(str(SCRIPT), str(PATH), str(n), PATH.name, seconds=120.0)
    if proc.returncode != 0:
        sys.exit(proc.stderr.strip()[:200])
    fields = proc.stdout.strip().split("\x1f")
    if fields[0] != "standard" or len(fields) != n + 2:
        sys.exit(f"expected {n} heights, got {len(fields) - 2}")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: angled_formats (capture.angled_formats); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; Normal font Calibri 11, cells Calibri 11"
            f" in column A at the default width; rows fitted by Excel on opening; standard height"
            f" {fields[1]}\n"
        )
        out.write("angle\tvalue\tformat\trow_height\n")
        for (angle, value, fmt), height in zip(rows(), fields[2:], strict=True):
            out.write(f"{angle}\t{value!r}\t{fmt}\t{height}\n")
    print(f"wrote {OUTPUT}: {n} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        print(build())
    elif command == "run":
        run()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
