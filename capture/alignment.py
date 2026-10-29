"""Capture Excel's default horizontal alignment of values under number formats.

Usage: ``uv run --group capture python -m capture.alignment``. Neither AppleScript's
``horizontal alignment`` (which reports ``general``) nor ``string value`` tells where Excel puts a
value, so this reads Excel's PDF export (the method of sheetpeek's layout capture): one case per
row in a wide column with gridlines printed, at 100 % scale, so one PDF point is one sheet point;
pdfminer.six gives each row's text run, and its place between the two gridlines gives the
alignment (left text starts 3 points after the left gridline, right text ends 2 points before
the right one). Writes ``tests/captures/alignment.layout.tsv``: tag, value type and repr, format,
the text the PDF holds and the alignment (``left``, ``center``, ``right``, or empty for no
text). CAPTURING.md section 2 item 31.

The clipboard is not used. The workbook is written inside Excel's container and closed without
saving; alerts are off while the script runs.
"""

from __future__ import annotations

import csv
import datetime as dt
import itertools
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles.numbers import BUILTIN_FORMATS_REVERSE

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.cases import Case, CellError

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "alignment.layout.tsv"
WORKBOOK = WORKBOOK_DIR / "alignment.xlsx"
EXPORT = WORKBOOK_DIR / "alignment.pdf"
SCRIPT = Path(__file__).with_name("pdf_export.applescript")
COLUMN_WIDTH = 40  # file units: about 245 points at Calibri 11, wider than every text here
EDGE = 6.0  # points: a run starting (ending) this close to the left (right) gridline

_VALUES: tuple[tuple[str, object], ...] = (
    ("int", 12345),
    ("float", 1.5),
    ("negative", -1),
    ("zero", 0),
    ("date-serial", 36529),
    ("bool", True),
    ("text", "abc"),
)
_FORMATS: tuple[str, ...] = (
    "General", "0", "0.00", "@", '@" units"', '"x"@', "0;0;0;@", ";;;@", "[Red]@",
    "General;General;General;@", "@@", "m/d/yy", '"abc"', ";;;", "0;-0;;@", "_(@_)", "*-@",
    "@*-", "[>100]0", "[Red]General", '0;-0;0;"t"', "0;-0;0;@*-",
)  # fmt: skip
_ERRORS: tuple[tuple[str, str], ...] = (
    ("#DIV/0!", "=1/0"),
    ("#N/A", "=NA()"),
    ("#NUM!", "=SQRT(-1)"),
    ("#VALUE!", '="a"+1'),
    ("#NAME?", "=XYZZY()"),
    ("#REF!", "=#REF!"),
    ("#NULL!", "=SUM(A1:A2 B1:B2)"),
)
_ERROR_FORMATS: tuple[str, ...] = ("General", "0", "@", "[Red]0", ";;;", "0;-0;0;@")


def cases() -> list[Case]:
    """Every value under every format, then every error under a few."""
    out = [Case("alignment", tag, value, fmt) for fmt in _FORMATS for tag, value in _VALUES]
    out += [
        Case("alignment", f"error {code}", CellError(code), fmt, formula=formula)
        for fmt in _ERROR_FORMATS
        for code, formula in _ERRORS
    ]
    return out


def build() -> None:
    """Write the workbook: one case per row in column A, gridlines printed, 100 % scale."""
    BUILTIN_FORMATS_REVERSE.clear()  # keep format strings verbatim
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = "alignment"
    for row, case in enumerate(cases(), 1):
        cell = ws.cell(row=row, column=1)
        if case.formula is not None:
            cell.value = case.formula
        else:
            cell.value = case.value
            if isinstance(case.value, str):
                cell.data_type = "s"
        cell.number_format = case.format
    ws.column_dimensions["A"].width = COLUMN_WIDTH
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 100
    ws.print_options.gridLines = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.25
    ws.page_margins.header = ws.page_margins.footer = 0
    ws.print_area = f"A1:A{len(cases())}"
    WORKBOOK_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(WORKBOOK)


def export() -> str:
    """Have Excel save the workbook as PDF; return Excel's version."""
    EXPORT.unlink(missing_ok=True)
    result = subprocess.run(
        ["osascript", str(SCRIPT), str(WORKBOOK), str(EXPORT)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EXPORT, BUILD_DIR / EXPORT.name)
    return result.stdout.strip()


@dataclass
class Line:
    """One sheet row as printed: its text and where it sits between the gridlines."""

    text: str
    align: str


def read(path: Path) -> list[Line]:
    """Per sheet row, in order across pages: the printed text and its alignment."""
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LAParams, LTChar, LTRect, LTTextContainer

    out: list[Line] = []
    for page in extract_pages(path, laparams=LAParams(char_margin=2000, line_margin=0.01)):
        rects = [r for r in page if isinstance(r, LTRect)]
        horizontal = sorted(
            {round(r.y0, 2) for r in rects if r.height <= 1.01 and r.width > 1.01}, reverse=True
        )
        vertical = sorted({round(r.x0, 2) for r in rects if r.width <= 1.01 and r.height > 1.01})
        left, right = vertical[0] + 1, vertical[-1]  # the inner edges of the two gridlines
        chars = [
            c
            for box in page
            if isinstance(box, LTTextContainer)
            for line in box
            for c in line
            if isinstance(c, LTChar)
        ]
        for high, low in itertools.pairwise(horizontal):
            mine = sorted((c for c in chars if low < (c.y0 + c.y1) / 2 < high), key=lambda c: c.x0)
            if not mine:
                out.append(Line("", ""))
                continue
            text = "".join(c.get_text() for c in mine)
            x0, x1 = mine[0].x0, mine[-1].x1
            middle = (left + right) / 2
            if x0 - left < EDGE:
                align = "left"
            elif right - x1 < EDGE:
                align = "right"
            elif abs((x0 + x1) / 2 - middle) < 2:
                align = "center"
            else:
                align = "?"
            out.append(Line(text, align))
    return out


def write(version: str) -> None:
    """Read the export and write the layout TSV with its header."""
    lines = read(BUILD_DIR / EXPORT.name)
    named = cases()
    if len(lines) != len(named):
        sys.exit(f"found {len(lines)} rows in the PDF, expected {len(named)}")
    with OUTPUT.open("w", encoding="utf-8", newline="") as out:
        out.write(f"# suite: alignment (capture.alignment); Excel {version} for Mac, en_US\n")
        out.write("# workbook_font: Calibri 11\n")
        out.write(f"# captured: {dt.date.today().isoformat()}; method: PDF export, pdfminer.six\n")
        out.write("# align: where the printed text sits between the gridlines; empty: no text\n")
        writer = csv.writer(out, dialect="excel-tab", lineterminator="\n")
        writer.writerow(["tag", "value_type", "value_repr", "format", "text", "align"])
        for case, line in zip(named, lines, strict=True):
            writer.writerow(
                [case.tag, case.value_type, case.value_repr, repr(case.format), line.text,
                 line.align]
            )  # fmt: skip
    print(f"wrote {OUTPUT}: {len(named)} rows")


def main() -> None:
    """Build, export through Excel, read, write."""
    build()
    print(f"built {WORKBOOK}; exporting through Excel")
    write(export())


if __name__ == "__main__":
    main()
