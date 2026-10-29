"""Run every capture suite under tests/captures against the formatter.

One test per suite file; all mismatching rows are collected and reported together. Width-less
rows are compared exactly; width rows are compared in Excel mode (the suite's column width and
font, exact match, ``#`` counts included) and in character mode, where a ``#`` run matches any
``#`` run and rows at a fractional width and rows Excel fitted by glyph widths rather than by
character count are skipped under their own labels.
Every count is printed per suite and recorded as a test property; zero counts are left out of
the printed line. Color rows Excel's DisplayFormat API misreports are skipped and counted.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from capture.schema import Row, read_rows

from xlformat import DEFAULT_PALETTE, format_cell, format_value

CAPTURE_DIR = Path(__file__).parent / "captures"
SUITE_FILES = sorted(
    p
    for p in CAPTURE_DIR.iterdir()
    if p.name.endswith((".tsv", ".tsv.gz")) and ".measure." not in p.name
)

# Rows where Excel's DisplayFormat font color, the capture's source for `expected_color`,
# reports automatic black although the sheet paints the section color (verified by eye on
# 2026-10-24: B9 zero under [Red]0 is red, B132 'abc' under [Cyan]@ is cyan, B147 the '#' run
# under [Red]m/d/yy is red). Three patterns: a value of exactly zero rendered by section 1 or 2,
# text or a boolean through a colored text section, and a '#' run. The package is right and the
# capture is wrong for these rows, so they are skipped, not matched (CAPTURING.md section 2
# item 21).
_ZERO = "zero in section 1 or 2 reads as automatic black"
_TEXT = "text through a colored text section reads as automatic black"
_HASH = "a # run reads as automatic black"
DISPLAY_FORMAT_COLOR_MISREPORTS: dict[tuple[str, str], str] = {
    ("'[White]0'", "0"): _ZERO,
    ("'[Red]0'", "0"): _ZERO,
    ("'[Green]0'", "0"): _ZERO,
    ("'[Blue]0'", "0"): _ZERO,
    ("'[Yellow]0'", "0"): _ZERO,
    ("'[Magenta]0'", "0"): _ZERO,
    ("'[Cyan]0'", "0"): _ZERO,
    ("'[Blue]0;[Red]-0;[Green]0;[Magenta]@'", "'abc'"): _TEXT,
    ("'[Red][<0]0;[Blue]0'", "0"): _ZERO,
    ("'[Green][=0]\"zero\";[Red]0'", "0"): _ZERO,
    ("'0;-0;0;[Magenta]@'", "'abc'"): _TEXT,
    ("'0;-0;0;[Magenta]@'", "True"): _TEXT,
    ("'[Cyan]@'", "'abc'"): _TEXT,
    ("'[Red]m/d/yy'", "-1234.5"): _HASH,
    ("'[Red][h]:mm'", "-1234.5"): _HASH,
}


def _excel_value(row: Row, header: dict[str, str]) -> object:
    """The value Excel loaded for the row.

    Suites captured before 2026-10-24 (no ``numbers`` header) were written by openpyxl with
    ``"%.16g"`` texts, which denote a different double for about half of all floats; the row's
    expected text belongs to that double (CAPTURING.md section 2 item 16). Formula rows always
    denote the Python double.
    """
    value = row.value
    if (
        header.get("numbers") != "repr"
        and row.value_type == "float"
        and not row.tag.startswith("formula:")
    ):
        assert isinstance(value, float)
        return float(f"{value:.16g}")
    return value


def _character_width(width: float | None) -> int | str | None:
    """The character-mode width for a row: an int, None for width-less rows, or "skip".

    Character mode takes a whole number of characters; a row captured at a fractional Excel
    width (8.43) has no character-mode meaning and is skipped.
    """
    if width is None:
        return None
    return int(width) if width == int(width) else "skip"


def _glyph_slack(expected: str, width: int | None) -> bool:
    """Excel fitted more characters than the width: commas, points, spaces and minus signs are
    narrower than digits. Character mode cannot reproduce that by contract; such rows are
    Excel-mode evidence and are skipped here (ARCHITECTURE.md section 8)."""
    return width is not None and len(expected) > width and set(expected) != {"#"}


def _glyph_tight(expected: str, got: str, width: int | None) -> bool:
    """The mirror case: Excel overflowed a text no longer than the width because a glyph such
    as '%' is wider than a digit. Only counts when our text has a non-digit character."""
    return (
        width is not None
        and set(expected) == {"#"}
        and set(got) != {"#"}
        and len(got) <= width
        and any(not c.isdigit() for c in got)
    )


def _texts_match(got: str, expected: str, width: int | None) -> bool:
    """Equal texts; at a character-mode width also two ``#`` runs of any length.

    Excel's count is glyph-measured (six at width 5), the character-mode count is the width by
    contract (CAPTURING.md section 2 items 3 and 23); 4,907 of the 5,396 ``#`` rows compared at
    a character width differ in count for that reason. Width-less rows (255 ``#``) and every
    Excel-mode row are compared exactly.
    """
    if width is None:
        return got == expected
    return got == expected or (set(got) == {"#"} and set(expected) == {"#"})


@pytest.mark.excel
@pytest.mark.parametrize("suite_file", SUITE_FILES, ids=lambda p: p.name.split(".")[0])
def test_suite(suite_file: Path, tier: str, record_property) -> None:  # type: ignore[no-untyped-def]
    header, rows = read_rows(suite_file)
    assert header.get("suite"), f"{suite_file}: missing suite header"
    if tier == "fast":
        rows = [r for r in rows if r.tier == "fast"]
    if not rows:
        pytest.skip(f"{suite_file.name}: no rows in tier {tier}")
    failures: list[tuple[Row, str]] = []
    excel_failures: list[tuple[Row, str]] = []
    color_failures: list[tuple[Row, str]] = []
    counts = {
        "compared": 0,  # character mode
        "skipped (fractional width)": 0,
        "skipped (glyph slack)": 0,
        "Excel mode compared": 0,
        "colors compared": 0,
        "color misreports skipped": 0,
    }
    font_name = header.get("workbook_font", "Calibri 11").split()
    font = (font_name[0], float(font_name[1])) if len(font_name) == 2 else ("Calibri", 11.0)
    for row in rows:
        fmt = row.format_string
        value = _excel_value(row, header)
        date1904 = row.date1904 == "1"
        if row.width_value is not None:
            # Excel mode: the column width as requested (Excel snaps it), the workbook's font;
            # every width row must match exactly, '#' counts included.
            counts["Excel mode compared"] += 1
            try:
                got_excel = format_value(
                    value, fmt, column_width=row.width_value, font=font, date1904=date1904
                )
            except Exception as exc:
                got_excel = f"<{type(exc).__name__}: {exc}>"
            if got_excel != row.expected_text:
                excel_failures.append((row, got_excel))
        width = _character_width(row.width_value)
        if width == "skip":
            counts["skipped (fractional width)"] += 1
            continue
        if _glyph_slack(row.expected_text, width):
            counts["skipped (glyph slack)"] += 1
            continue
        try:
            got = format_value(value, fmt, width=width, date1904=date1904)
        except Exception as exc:
            got = f"<{type(exc).__name__}: {exc}>"
        if _glyph_tight(row.expected_text, got, width):
            counts["skipped (glyph slack)"] += 1
            continue
        counts["compared"] += 1
        if not _texts_match(got, row.expected_text, width):
            failures.append((row, got))
        if row.expected_color and (row.format, row.value_repr) in DISPLAY_FORMAT_COLOR_MISREPORTS:
            counts["color misreports skipped"] += 1
        elif row.expected_color:
            # Excel reports the display format's font color; a cell without a section color
            # reads as the automatic black, 000000, which [Black] also gives.
            counts["colors compared"] += 1
            cell = format_cell(value, fmt, date1904=date1904)
            ours = DEFAULT_PALETTE[cell.color] if cell.color else "000000"
            if ours != row.expected_color:
                color_failures.append((row, ours))
    counts["failed"] = len(failures)
    counts["Excel mode failed"] = len(excel_failures)
    counts["colors failed"] = len(color_failures)
    record_property("rows", len(rows))
    for label, count in counts.items():
        record_property(label.replace(" ", "_").replace("(", "").replace(")", ""), count)
    shown = ", ".join(f"{count} {label}" for label, count in counts.items() if count)
    print(f"\n{suite_file.name} [{tier}]: {len(rows)} rows, {shown}")
    if excel_failures:
        lines = [
            f"{r.format} {r.value_type} {r.value_repr} column_width={r.width} "
            f"expected={r.expected_text!r} got={got!r}"
            for r, got in excel_failures[:40]
        ]
        pytest.fail(f"{len(excel_failures)} Excel-mode rows differ:\n" + "\n".join(lines))
    if color_failures:
        lines = [
            f"{r.format} {r.value_type} {r.value_repr} expected_color={r.expected_color} got={got}"
            for r, got in color_failures[:40]
        ]
        pytest.fail(f"{len(color_failures)} colors differ:\n" + "\n".join(lines))
    if failures:
        lines = [
            f"{r.format} {r.value_type} {r.value_repr} width={r.width or 'none'} "
            f"expected={r.expected_text!r} got={got!r}"
            for r, got in failures[:40]
        ]
        pytest.fail(
            f"{len(failures)} of {counts['compared']} compared rows differ:\n" + "\n".join(lines)
        )
