"""Where Excel puts text vertically in a cell, from its PDF export (a probe).

    uv run --group capture python -m capture.vertical_pdf build   # the workbooks, no Excel
    uv run --group capture python -m capture.vertical_pdf run     # export each and read them
    uv run --group capture python -m capture.vertical_pdf read    # read saved exports again

One workbook per Normal font (Calibri 11, Aptos Narrow 11, Arial 10), one case per row in a wide
column, gridlines printed at 100 % scale, so one PDF point is one sheet point. Each case is the
text "Hxg" in one face: the 27 regular faces (Calibri, Aptos Narrow and Arial at the nine
measured sizes, ``SIZES``, 9 to 24 points) aligned top, center and bottom in rows of a stated
height (the face's font height F, F + 10 and 2F); each face in a row Excel fits itself; two and
three wrapped lines; and the four styles at 11 and 16 points. pdfminer.six gives every
character's baseline (the text origin, not the font's box) and each row's gridlines; the probe
writes, per row, the printed row's top and height and each line's baseline below the top, in
points, to ``capture/build/probes/vertical_pdf.tsv``. No clipboard; alerts are off while Excel
exports.
"""

from __future__ import annotations

import datetime as dt
import itertools
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.cases import CellFont
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript_or_exit

SCRIPT = Path(__file__).with_name("pdf_export.applescript")
OUTPUT = BUILD_DIR / "probes" / "vertical_pdf.tsv"
NORMALS = (CellFont("Calibri", 11), CellFont("Aptos Narrow", 11), CellFont("Arial", 10))
FAMILIES = ("Calibri", "Aptos Narrow", "Arial")
SIZES = (9, 10, 11, 12, 14, 16, 18, 20, 24)
# the font height F of each face (CAPTURING.md section 2 item 44)
FONT_HEIGHT = {
    "Calibri": {9: 12, 10: 14, 11: 15, 12: 16, 14: 19, 16: 21, 18: 24, 20: 26, 24: 31},
    "Aptos Narrow": {9: 13, 10: 14, 11: 15, 12: 16, 14: 19, 16: 22, 18: 24, 20: 27, 24: 32},
    "Arial": {9: 12, 10: 13, 11: 14, 12: 16, 14: 18, 16: 20, 18: 23, 20: 25, 24: 30},
}
TEXT = "Hxg"
COLUMN_WIDTH = 30  # file units, wider than every text here


@dataclass(frozen=True, slots=True)
class Case:
    """One row: tag, face, vertical alignment, stated row height (None: Excel fits), lines."""

    tag: str
    font: CellFont
    vertical: str = "bottom"
    height: float | None = None
    lines: int = 1


def cases() -> list[Case]:
    out: list[Case] = []
    for family in FAMILIES:
        for size in SIZES:
            f = CellFont(family, size)
            fh = FONT_HEIGHT[family][size]
            out.append(Case(f"natural {f.label}", f))
            for vertical in ("top", "center", "bottom"):
                for label, h in (("F", fh), ("F+10", fh + 10), ("2F", 2 * fh)):
                    out.append(Case(f"{vertical} {label} {f.label}", f, vertical, float(h)))
            out.extend(
                Case(f"wrapped {lines} {f.label}", f, "bottom", None, lines) for lines in (2, 3)
            )
    for family in FAMILIES:
        for size in (11, 16):
            for bold, italic in ((True, False), (False, True), (True, True)):
                f = CellFont(family, size, bold, italic)
                out.append(Case(f"natural {f.label}", f))
                fh = FONT_HEIGHT[family][size] + 10
                out.append(Case(f"bottom F+10 {f.label}", f, "bottom", float(fh)))
    # appended later, so that the rows above keep their places: the faces without glyph tables
    # at the table sizes, and Calibri and Aptos Narrow at sizes between them
    extra = [(family, size) for family in EXTRA_FAMILIES for size in SIZES]
    extra += [(family, size) for family in ("Calibri", "Aptos Narrow") for size in EXTRA_SIZES]
    for family, size in extra:
        f = CellFont(family, size)
        fh = _height(family, size)
        out.append(Case(f"natural {f.label}", f))
        for vertical in ("top", "center", "bottom"):
            for label, h in (("F", fh), ("F+10", fh + 10), ("2F", 2 * fh)):
                out.append(Case(f"{vertical} {label} {f.label}", f, vertical, float(h)))
        out.extend(Case(f"wrapped {lines} {f.label}", f, "bottom", None, lines) for lines in (2, 3))
    return out


EXTRA_FAMILIES = ("Times New Roman", "Courier New", "Georgia")
EXTRA_SIZES = (13, 15, 17, 22)


def _height(family: str, size: int) -> int:
    """The face's height by the row model (CAPTURING.md section 2 item 57)."""
    from xlformat import row_metrics

    return int(row_metrics((family, float(size))).font_height)


def _paths(normal: CellFont) -> tuple[Path, Path]:
    stem = "vertical_" + normal.label.replace(" ", "_").lower()
    return (WORKBOOK_DIR / f"{stem}.xlsx").resolve(), (WORKBOOK_DIR / f"{stem}.pdf").resolve()


def build(normal: CellFont) -> Path:
    """Write the workbook for one Normal font into Excel's container."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.indexed_list import IndexedList

    wb = Workbook()
    base = Font(
        name=normal.name, sz=normal.size, family=2,
        scheme="minor" if normal.name == "Calibri" else None,
    )  # fmt: skip
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    ws.title = "vertical"
    for r, case in enumerate(cases(), 1):
        cell = ws.cell(row=r, column=1, value="\n".join([TEXT] * case.lines))
        cell.data_type = "s"
        f = case.font
        cell.font = Font(name=f.name, sz=f.size, b=f.bold, i=f.italic, family=2)
        cell.alignment = Alignment(vertical=case.vertical, wrap_text=case.lines > 1)
        if case.height is not None:
            ws.row_dimensions[r].height = case.height
    ws.column_dimensions["A"].width = COLUMN_WIDTH
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 100
    ws.print_options.gridLines = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.25
    ws.page_margins.header = ws.page_margins.footer = 0
    ws.print_area = f"A1:A{len(cases())}"
    path, _ = _paths(normal)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


@dataclass(frozen=True, slots=True)
class Printed:
    """One printed row: its top (PDF y), its height, and each line's baseline below the top."""

    top: float
    height: float
    baselines: tuple[float, ...]
    text: str


def read(pdf: Path) -> list[Printed]:
    """Per sheet row, in order across pages."""
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LAParams, LTChar, LTRect, LTTextContainer

    out: list[Printed] = []
    for page in extract_pages(pdf, laparams=LAParams(char_margin=2000, line_margin=0.01)):
        rects = [r for r in page if isinstance(r, LTRect)]
        horizontal = sorted(
            {round(r.y0, 3) for r in rects if r.height <= 1.01 and r.width > 1.01}, reverse=True
        )
        chars = [
            c
            for box in page
            if isinstance(box, LTTextContainer)
            for line in box
            for c in line
            if isinstance(c, LTChar)
        ]
        for high, low in itertools.pairwise(horizontal):
            mine = [c for c in chars if low < c.matrix[5] < high]
            origins = sorted({round(c.matrix[5], 3) for c in mine}, reverse=True)
            text = "".join(c.get_text() for c in sorted(mine, key=lambda c: (-c.matrix[5], c.x0)))
            out.append(
                Printed(
                    top=high,
                    height=round(high - low, 3),
                    baselines=tuple(round(high - y, 3) for y in origins),
                    text=text,
                )
            )
    return out


def export(normal: CellFont) -> None:
    path, pdf = _paths(normal)
    build(normal)
    pdf.unlink(missing_ok=True)
    result = osascript_or_exit(str(SCRIPT), str(path), str(pdf), seconds=300.0)
    if result.returncode != 0:
        sys.exit(f"{normal.label}: {result.stderr.strip()[:200]}")
    shutil.copy2(pdf, BUILD_DIR / pdf.name)
    print(f"{normal.label}: exported {pdf.name}")


def write() -> None:
    """Read every saved export and write the probe's TSV."""
    rows = []
    for normal in NORMALS:
        printed = read(BUILD_DIR / _paths(normal)[1].name)
        if len(printed) != len(cases()):
            sys.exit(f"{normal.label}: {len(printed)} printed rows, expected {len(cases())}")
        for case, p in zip(cases(), printed, strict=True):
            if p.text.replace("\n", "") != TEXT * case.lines:
                sys.exit(f"{normal.label} {case.tag}: printed {p.text!r}")
            stated = "" if case.height is None else f"{case.height:g}"
            baselines = " ".join(f"{b:g}" for b in p.baselines)
            rows.append(f"{normal.label}\t{case.tag}\t{stated}\t{p.height:g}\t{baselines}")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: vertical_pdf (capture.vertical_pdf); Excel for Mac, en_US; a probe\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; method: PDF export at 100 %, gridlines "
            "printed; pdfminer.six text origins (baselines) and gridlines, in points\n"
        )
        out.write("normal_font\ttag\tstated_height\tprinted_height\tbaselines_below_top\n")
        out.write("\n".join(rows) + "\n")
    print(f"wrote {OUTPUT}: {len(rows)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        for normal in NORMALS:
            print(build(normal))
    elif command == "run":
        if not answers():
            sys.exit("Excel does not answer: nothing run")
        for normal in NORMALS:
            export(normal)
        write()
    elif command == "read":
        write()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
