"""Build the workbook for one suite with openpyxl.

Sheet ``cases``: column A holds the 1-based case index; each capture column (B, C, ...) holds
the value with the case's number format applied, one column per entry in ``Suite.widths``.
Column widths are set later by the capture script, in Excel's own units, except in a suite with
``width_source="file"``, whose widths are written into the file as ``<col width>``. Sheet
``measure`` (optional) holds each distinct (value, format, cell font) in its own cell of row 1,
for autofit measurement. A case's ``font`` is set on its cells; others inherit the Normal style.
A JSON manifest next to the workbook tells the capture script what to read.
"""

from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles.numbers import BUILTIN_FORMATS_REVERSE
from openpyxl.utils import get_column_letter
from openpyxl.utils.indexed_list import IndexedList

from capture.cases import CellFont, Suite, distinct_measure_cases

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build"  # manifests, generated scripts, logs
# Workbooks go into Excel's own sandbox container: opening a file from anywhere else makes
# macOS show a "Grant Access" dialog for every file (CAPTURING.md §2, permission rules).
_CONTAINER = Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data"
WORKBOOK_DIR = _CONTAINER / "Documents" / "xlformat"


@dataclass(frozen=True, slots=True)
class ColumnSpec:
    """One capture column: its letter, Excel width (None = width-less) and capture method."""

    letter: str
    width: float | None
    # "clipboard" (width-less), "string_value" (set to the width by the script), "file_width"
    # (written into the file) or "file_default" (no <col> entry); all but the first read
    # string value
    method: str


@dataclass(frozen=True, slots=True)
class Manifest:
    """What the capture script needs to know about a built workbook."""

    suite: str
    workbook: str
    n_rows: int
    columns: list[ColumnSpec]
    measure_cells: list[str]
    date1904: bool
    colors: bool = False  # read the display-format font color of every cell in column B


def _set_value(
    ws,  # noqa: ANN001
    row: int,
    col: int,
    value: object,
    number_format: str,
    formula: str | None = None,
    written: str | None = None,
) -> None:
    cell = ws.cell(row=row, column=col)
    if written is None and formula is None and isinstance(value, float) and math.isfinite(value):
        # openpyxl writes floats with "%.16g", which denotes a different double for about half
        # of all floats (CAPTURING.md section 2 item 16); repr is the shortest exact text.
        written = repr(value)
    if written is not None:
        if float(written) != value:
            msg = f"written text {written!r} does not denote {value!r}"
            raise ValueError(msg)
        cell.value = written
        cell.data_type = "n"  # the writer emits a str as-is, so Excel parses exactly this text
    elif formula is not None:
        if not formula.startswith("="):
            msg = f"formula must start with '=': {formula!r}"
            raise ValueError(msg)
        cell.value = formula  # openpyxl stores a leading '=' string as a formula (data_type f)
    else:
        cell.value = value
        if isinstance(value, str):
            cell.data_type = "s"  # never let a leading '=' become a formula
    cell.number_format = number_format


def _font(font: CellFont) -> Font:
    return Font(name=font.name, sz=font.size, b=font.bold, i=font.italic, family=2)


def _method(suite: Suite, width: float | None) -> str:
    if suite.width_source == "file":
        return "file_default" if width is None else "file_width"
    return "clipboard" if width is None else "string_value"


def build(suite: Suite) -> Manifest:
    """Write ``capture/build/<suite>.xlsx`` and its manifest; return the manifest."""
    BUILTIN_FORMATS_REVERSE.clear()  # keep format strings verbatim, no built-in ids
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # formula cells have no cached value; Excel computes
    if suite.font_size is not None or suite.font_name != "Calibri":
        # the Normal style's font: the first font of styles.xml and the named style's font.
        # openpyxl's theme has Calibri as its minor font, so only Calibri keeps the scheme.
        scheme = "minor" if suite.font_name == "Calibri" else None
        size = 11 if suite.font_size is None else suite.font_size
        normal = Font(name=suite.font_name, sz=size, family=2, scheme=scheme)
        wb._fonts = IndexedList([normal])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = normal
    ws = wb.active
    ws.title = "cases"
    for attr, value in suite.sheet_format:
        setattr(ws.sheet_format, attr, value)
    columns: list[ColumnSpec] = []
    for k, width in enumerate(suite.widths):
        letter = get_column_letter(k + 2)
        columns.append(ColumnSpec(letter, width, _method(suite, width)))
        if suite.width_source == "file" and width is not None:
            ws.column_dimensions[letter].width = width
    for row, case in enumerate(suite.cases, 1):
        ws.cell(row=row, column=1, value=row)
        for k in range(len(suite.widths)):
            _set_value(ws, row, k + 2, case.value, case.format, case.formula, case.written)
            if case.font is not None:
                ws.cell(row=row, column=k + 2).font = _font(case.font)
    measure_cells: list[str] = []
    if suite.measure:
        ms = wb.create_sheet("measure")
        for k, case in enumerate(distinct_measure_cases(suite.cases), 1):
            _set_value(ms, 1, k, case.value, case.format)
            if case.font is not None:
                ms.cell(row=1, column=k).font = _font(case.font)
            measure_cells.append(f"{get_column_letter(k)}1")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    WORKBOOK_DIR.mkdir(parents=True, exist_ok=True)
    path = (WORKBOOK_DIR / f"{suite.name}.xlsx").resolve()
    wb.save(path)
    manifest = Manifest(
        suite=suite.name,
        workbook=str(path),
        n_rows=len(suite.cases),
        columns=columns,
        measure_cells=measure_cells,
        date1904=suite.date1904,
        colors=suite.colors,
    )
    (BUILD_DIR / f"{suite.name}.manifest.json").write_text(json.dumps(asdict(manifest), indent=1))
    return manifest
