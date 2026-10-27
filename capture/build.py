"""Build the workbook for one suite with openpyxl.

Sheet ``cases``: column A holds the 1-based case index; each capture column (B, C, ...) holds
the value with the case's number format applied, one column per entry in ``Suite.widths``.
Column widths are set later by the capture script, in Excel's own units. Sheet ``measure``
(optional) holds each distinct (value, format) in its own cell of row 1, for autofit measurement.
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

from capture.cases import Suite, distinct_value_formats

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
    method: str  # "clipboard" (width-less) or "string_value" (at the given width)


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


def build(suite: Suite) -> Manifest:
    """Write ``capture/build/<suite>.xlsx`` and its manifest; return the manifest."""
    BUILTIN_FORMATS_REVERSE.clear()  # keep format strings verbatim, no built-in ids
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # formula cells have no cached value; Excel computes
    if suite.font_size is not None:
        # the Normal style's font: the first font of styles.xml and the named style's font
        normal = Font(name="Calibri", sz=suite.font_size, family=2, scheme="minor")
        wb._fonts = IndexedList([normal])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = normal
    ws = wb.active
    ws.title = "cases"
    columns: list[ColumnSpec] = []
    for k, width in enumerate(suite.widths):
        letter = get_column_letter(k + 2)
        columns.append(ColumnSpec(letter, width, "clipboard" if width is None else "string_value"))
    for row, case in enumerate(suite.cases, 1):
        ws.cell(row=row, column=1, value=row)
        for k in range(len(suite.widths)):
            _set_value(ws, row, k + 2, case.value, case.format, case.formula, case.written)
    measure_cells: list[str] = []
    if suite.measure:
        ms = wb.create_sheet("measure")
        for k, (value, number_format) in enumerate(distinct_value_formats(suite.cases), 1):
            _set_value(ms, 1, k, value, number_format)
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
