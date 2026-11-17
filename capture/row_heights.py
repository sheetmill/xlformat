"""Row heights in Excel for Mac: per Normal font, per cell font, per line of wrapped text.

    uv run python -m capture.row_heights build          # write the workbooks, no Excel
    uv run python -m capture.row_heights run [LABEL]    # capture all, or one Normal font
    uv run python -m capture.row_heights write          # write the layout file (OUTPUT)

One workbook per Normal font: Calibri and Aptos Narrow at every measured size (``SIZES``, 9 to
24 points) in four styles, 72 workbooks. Each holds the same rows (``ROWS``): a number in the
Normal font, a row without an element, a number in each measured cell font, an empty cell
styled with a large or small font, two cell fonts in one row, and wrapped text of one to six
lines (line feeds, wrap on) in the Normal font and in each regular cell font. The sheet keeps
openpyxl's ``<sheetFormatPr>``, without ``customHeight``, as files written by openpyxl have it.
Excel fits every row to its cells when it opens the file; the script reads each row's height
and the sheet's standard height in points, closes the workbook without saving and restores
alerts. No clipboard. ``run`` keeps what it read in ``capture/build/row_heights.json`` so that
one Normal font can be captured again alone; ``write`` gathers them into the layout file.
"""

from __future__ import annotations

import datetime as dt
import gzip
import json
import sys
from dataclasses import dataclass
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.cases import CellFont
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript_or_exit

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "row_height.layout.tsv.gz"
SCRIPT = Path(__file__).with_name("row_heights.applescript")
STATE = BUILD_DIR / "row_heights.json"  # what `run` read so far (capture/build: not committed)
DIRECTORY = WORKBOOK_DIR / "row_heights"
FAMILIES = ("Calibri", "Aptos Narrow")
SIZES = (9, 10, 11, 12, 14, 16, 18, 20, 24)
STYLES = ((False, False), (True, False), (False, True), (True, True))
MAX_LINES = 6


def fonts() -> tuple[CellFont, ...]:
    """Every measured face: both families, nine sizes, four styles (72)."""
    return tuple(
        CellFont(family, size, bold, italic)
        for family in FAMILIES
        for size in SIZES
        for bold, italic in STYLES
    )


@dataclass(frozen=True, slots=True)
class Cell:
    """One cell of a row: column (1 = A), value (None: empty but styled), font, wrap."""

    column: int
    value: object
    font: CellFont | None = None
    wrap: bool = False


@dataclass(frozen=True, slots=True)
class RowCase:
    """One row of every workbook; ``cells`` empty means the row has no element."""

    tag: str
    cells: tuple[Cell, ...]


def _lines(n: int) -> str:
    return "\n".join(["x"] * n)  # one short glyph a line: no line wraps at any width


def rows() -> tuple[RowCase, ...]:
    """The rows every workbook holds, in order (row 1 first)."""
    out = [
        RowCase("number", (Cell(1, 1234),)),
        RowCase("no element", ()),
        RowCase("number again", (Cell(1, 5678),)),
    ]
    out.extend(RowCase(f"cell {f.label}", (Cell(1, 1234, f),)) for f in fonts())
    for family in FAMILIES:
        for size in (9, 16):
            f = CellFont(family, size)
            out.append(RowCase(f"empty styled {f.label}", (Cell(1, None, f),)))
    for family in FAMILIES:
        for size in SIZES:
            f = CellFont(family, size)
            out.append(RowCase(f"normal and {f.label}", (Cell(1, 1234), Cell(2, 1234, f))))
    small, large = CellFont("Calibri", 9), CellFont("Aptos Narrow", 16)
    out.append(RowCase(f"{small.label} and {large.label}", (Cell(1, 1, small), Cell(2, 1, large))))
    out.extend(
        RowCase(f"wrapped {n} lines", (Cell(1, _lines(n), wrap=True),))
        for n in range(1, MAX_LINES + 1)
    )
    for family in FAMILIES:
        for size in SIZES:
            f = CellFont(family, size)
            out.extend(
                RowCase(f"wrapped {n} lines {f.label}", (Cell(1, _lines(n), f, True),))
                for n in range(1, MAX_LINES + 1)
            )
    # wrapped text beside an unwrapped cell in another font, in the same row
    for wrap_font, lines, other in (
        (None, 1, CellFont("Calibri", 14)),
        (None, 2, CellFont("Calibri", 16)),
        (None, 3, CellFont("Aptos Narrow", 16)),
        (CellFont("Calibri", 9), 1, CellFont("Calibri", 16)),
        (CellFont("Calibri", 9), 2, CellFont("Calibri", 16)),
        (CellFont("Calibri", 9), 3, CellFont("Calibri", 14)),
        (CellFont("Calibri", 14), 2, CellFont("Calibri", 16)),
        (CellFont("Aptos Narrow", 9), 2, CellFont("Aptos Narrow", 16)),
    ):
        label = "normal" if wrap_font is None else wrap_font.label
        out.append(
            RowCase(
                f"wrapped {lines} lines {label} beside {other.label}",
                (Cell(1, _lines(lines), wrap_font, True), Cell(2, 1234, other)),
            )
        )
    out.append(RowCase("number at the end", (Cell(1, 1234),)))
    return tuple(out)


def file_name(normal: CellFont) -> str:
    return "rh_" + normal.label.replace(" ", "_").lower() + ".xlsx"


def build(normal: CellFont) -> Path:
    """Write the workbook for one Normal font into Excel's container; return its path."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.indexed_list import IndexedList

    def font(f: CellFont, scheme: str | None = None) -> Font:
        return Font(name=f.name, sz=f.size, b=f.bold, i=f.italic, family=2, scheme=scheme)

    wb = Workbook()
    # the Normal style's font: the first font of styles.xml and the named style's font; only
    # Calibri, openpyxl's theme minor font, keeps the scheme (as capture/build.py does)
    base = font(normal, "minor" if normal.name == "Calibri" else None)
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    ws.title = "rows"
    for r, case in enumerate(rows(), 1):
        for c in case.cells:
            cell = ws.cell(row=r, column=c.column, value=c.value)
            if isinstance(c.value, str):
                cell.data_type = "s"
            if c.font is not None:
                cell.font = font(c.font)
            if c.wrap:
                cell.alignment = Alignment(wrap_text=True)
    DIRECTORY.mkdir(parents=True, exist_ok=True)
    path = (DIRECTORY / file_name(normal)).resolve()
    wb.save(path)
    return path


def _clean() -> bool:
    proc = osascript_or_exit(
        "-e", 'tell application "Microsoft Excel" to get {count of workbooks, display alerts}',
        seconds=20.0,
    )  # fmt: skip
    return proc.returncode == 0 and proc.stdout.strip() == "0, true"


def run(only: str | None) -> None:
    """Capture every Normal font (or the one labeled ``only``) and keep the heights read."""
    if not answers() or not _clean():
        sys.exit("Excel must answer, hold no workbook and have alerts on: nothing run")
    normals = [f for f in fonts() if only is None or f.label == only]
    if not normals:
        sys.exit(f"no Normal font labeled {only!r}")
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    n = len(rows())
    for normal in normals:
        path = build(normal)
        proc = osascript_or_exit(str(SCRIPT), str(path), str(n), path.name, seconds=180.0)
        if proc.returncode != 0:
            sys.exit(f"{normal.label}: {proc.stderr.strip()[:200]}")
        fields = proc.stdout.strip().split("\x1f")
        if fields[0] != "standard" or len(fields) != n + 2:
            sys.exit(f"{normal.label}: expected {n} heights, got {len(fields) - 2}")
        state[normal.label] = {"standard": fields[1], "heights": fields[2:]}
        STATE.write_text(json.dumps(state, indent=1))
        print(f"{normal.label}: standard {fields[1]}, rows {fields[2]} .. {fields[-1]}")
    if not _clean():
        sys.exit("Excel is not clean after the run: look before anything else")


def write() -> None:
    """Gather the heights read into the layout file, one line per (Normal font, row)."""
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    missing = [f.label for f in fonts() if f.label not in state]
    if missing:
        sys.exit(f"not captured yet: {missing}")
    cases = rows()
    with gzip.open(OUTPUT, "wt", encoding="utf-8", newline="") as out:
        out.write("# suite: row_height (capture.row_heights); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            "# method: one workbook per Normal font, opened through AppleScript with alerts off; "
            "row height of each row and the sheet's standard height read in points; "
            "<sheetFormatPr> as openpyxl writes it (no customHeight)\n"
        )
        out.write(
            f"# captured: {dt.date.today().isoformat()}; wrapped lines are 'x' each, "
            "joined by line feeds in column A at the default width\n"
        )
        out.write("normal_font\trow\ttag\theight\n")
        for f in fonts():
            entry = state[f.label]
            out.write(f"{f.label}\t0\tstandard height\t{entry['standard']}\n")
            for r, (case, height) in enumerate(zip(cases, entry["heights"], strict=True), 1):
                out.write(f"{f.label}\t{r}\t{case.tag}\t{height}\n")
    print(f"wrote {OUTPUT}: {len(fonts())} Normal fonts, {len(cases)} rows each")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if not argv or argv[0] not in ("build", "run", "write"):
        sys.exit(__doc__)
    if argv[0] == "build":
        for f in fonts():
            build(f)
        print(f"built {len(fonts())} workbooks of {len(rows())} rows in {DIRECTORY}")
    elif argv[0] == "run":
        run(" ".join(argv[1:]) or None)
    else:
        write()


if __name__ == "__main__":
    main(sys.argv[1:])
