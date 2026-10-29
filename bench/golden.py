"""Render every captured row in every mode and write the outputs, for byte-for-byte comparison.

Usage: ``uv run python -m bench.golden [OUT]``. For each row of every suite under
``tests/captures`` (and the two layout files' inputs): ``format_cell`` width-less, at the
character width when the row has a whole width, and in Excel mode with the row's column width,
cell font and Normal font, as ``tests/test_captures.py`` renders it. Every field of each
``FormattedCell`` (or the exception) goes on one line. Prints the SHA-256 of the whole output;
two runs that print the same digest rendered every row identically. Writes OUT (default
``capture/build/golden.txt``) for a diff when they do not.
"""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

from capture.schema import read_rows

from xlformat import Font, file_column_metrics, format_cell

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))

from test_captures import SUITE_FILES, _cell_font, _excel_value  # noqa: E402


def _cell(value: object, fmt: str, **kwargs: object) -> str:
    try:
        c = format_cell(value, fmt, **kwargs)  # type: ignore[arg-type]
    except Exception as exc:
        return f"<{type(exc).__name__}: {exc}>"
    return repr((c.text, c.full_text, c.color, c.align, c.fill, c.overflow))


def lines() -> list[str]:
    """One line per row and mode."""
    out: list[str] = []
    for path in SUITE_FILES:
        header, rows = read_rows(path)
        m = re.match(r"(.+?) (\d+(?:\.\d+)?)\b", header.get("workbook_font", ""))
        normal = Font(m[1], float(m[2])) if m else Font("Calibri", 11.0)
        file_widths = header.get("width_source") == "file"
        sheet_format = dict(kv.split("=") for kv in header.get("sheet_format", "").split())
        for row in rows:
            value = _excel_value(row, header)
            fmt = row.format_string
            d = row.date1904 == "1"
            key = f"{path.name}|{row.tag}|{row.value_repr}|{row.format}|{row.width}"
            out.append(f"{key}|none|{_cell(value, fmt, date1904=d)}")
            width = row.width_value
            if file_widths:
                width = file_column_metrics(
                    width,
                    base_col_width=float(sheet_format.get("baseColWidth", 8)),
                    default_col_width=(
                        float(sheet_format["defaultColWidth"])
                        if "defaultColWidth" in sheet_format
                        else None
                    ),
                    font=normal,
                ).width
            if width is not None:
                font = _cell_font(row.tag) or normal
                out.append(
                    f"{key}|excel|"
                    + _cell(
                        value, fmt, column_width=width, font=font, normal_font=normal, date1904=d
                    )
                )
                if width == int(width) and width >= 1:
                    out.append(f"{key}|chars|{_cell(value, fmt, width=int(width), date1904=d)}")
    return out


def main() -> None:
    """Render, write, print the digest."""
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "capture" / "build" / "golden.txt"
    text = "\n".join(lines()) + "\n"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(text, encoding="utf-8")
    digest = hashlib.sha256(text.encode()).hexdigest()
    print(f"{text.count(chr(10)):,} renderings, sha256 {digest}  ({out})")


if __name__ == "__main__":
    main()
