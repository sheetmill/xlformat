"""The capture TSV schema shared by the converter, the capture runner and the test loader.

A suite file is UTF-8 text, optionally gzipped. Lines starting with ``#`` are header metadata
(``# key: value``); the first non-comment line names the columns; every other line is one row.
``format`` is ``repr()`` of the format string so that quotes, backslashes and trailing spaces
survive. ``value_repr`` is ``repr()`` of the value and ``value_type`` its type name
(``int``, ``float``, ``bool``, ``str``, ``NoneType``). ``width`` is the Excel column width
requested, in dialog units, or empty for the width-less rendering; ``actual_width`` is the width
Excel reported after setting it (Excel snaps to whole pixels), empty for width-less rows.
``date1904`` is ``0`` or ``1``.
``expected_color`` is empty when color was not captured. ``tier`` is ``fast`` or ``full``.
"""

from __future__ import annotations

import csv
import gzip
from ast import literal_eval
from collections.abc import Iterable, Iterator
from dataclasses import dataclass, fields
from pathlib import Path
from typing import TextIO

COLUMNS = (
    "suite",
    "tag",
    "value_type",
    "value_repr",
    "format",
    "width",
    "actual_width",
    "date1904",
    "expected_text",
    "expected_color",
    "tier",
)


@dataclass(frozen=True, slots=True)
class Row:
    """One captured row, all fields as the strings they are stored as."""

    suite: str
    tag: str
    value_type: str
    value_repr: str
    format: str
    width: str
    actual_width: str
    date1904: str
    expected_text: str
    expected_color: str
    tier: str

    @property
    def value(self) -> object:
        """The Python value: ``None`` for ``NoneType``, else ``literal_eval`` of the repr."""
        if self.value_type == "NoneType":
            return None
        return literal_eval(self.value_repr)

    @property
    def format_string(self) -> str:
        """The format string itself (the stored column is its repr)."""
        result = literal_eval(self.format)
        assert isinstance(result, str)
        return result

    @property
    def width_value(self) -> float | None:
        """Column width as a number, or ``None``."""
        return float(self.width) if self.width else None


assert tuple(f.name for f in fields(Row)) == COLUMNS


def write_rows(out: TextIO, header: dict[str, str], rows: Iterable[Row]) -> None:
    """Write header comments, the column line and the rows."""
    for key, value in header.items():
        out.write(f"# {key}: {value}\n")
    writer = csv.writer(out, dialect="excel-tab", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
    writer.writerow(COLUMNS)
    for row in rows:
        writer.writerow([getattr(row, name) for name in COLUMNS])


def open_text(path: Path) -> TextIO:
    """Open a plain or gzipped suite file for reading."""
    if path.suffix == ".gz":
        return gzip.open(path, "rt", newline="", encoding="utf-8")
    return path.open(newline="", encoding="utf-8")


def read_rows(path: Path) -> tuple[dict[str, str], list[Row]]:
    """Read a suite file; returns its header metadata and rows."""
    header: dict[str, str] = {}
    with open_text(path) as f:
        lines: Iterator[str] = iter(f)
        first = ""
        for line in lines:
            if line.startswith("#"):
                key, _, value = line[1:].partition(":")
                header[key.strip()] = value.strip()
                continue
            first = line
            break
        columns = tuple(first.rstrip("\n").split("\t"))
        if columns != COLUMNS:
            msg = f"{path}: unexpected columns {columns}"
            raise ValueError(msg)
        rows = [Row(*rec) for rec in csv.reader(lines, dialect="excel-tab")]
    return header, rows
