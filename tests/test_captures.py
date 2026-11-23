"""Run every capture suite under tests/captures against the formatter.

One test per suite file; all mismatching rows are collected and reported together. Width-less
rows are compared exactly; width rows are compared in Excel mode (the suite's column width and
font, exact match, ``#`` counts included) and in character mode, where a ``#`` run matches any
``#`` run and rows at a fractional width and rows Excel fitted by glyph widths rather than by
character count are skipped under their own labels.
Every count is printed per suite and recorded as a test property; zero counts are left out of
the printed line. The glyph-slack count of each suite and tier must equal the one listed in
``tests/captures/glyph_slack.json`` (``--record-glyph-slack`` writes it). Color rows Excel's
DisplayFormat API misreports are skipped and counted.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest
from capture.schema import Row, read_rows

from xlformat import (
    DEFAULT_PALETTE,
    ExcelError,
    Font,
    NumberFormat,
    column_metrics_from_points,
    file_column_metrics,
    format_cell,
    format_value,
    text_function,
)

CAPTURE_DIR = Path(__file__).parent / "captures"
SUITE_FILES = sorted(
    p
    # Excel for Mac's captures, and in windows/ Excel for Windows' (their header says so)
    for p in (*CAPTURE_DIR.iterdir(), *(CAPTURE_DIR / "windows").iterdir())
    # .measure. files hold autofit widths, .layout. files PDF positions; neither is a suite
    if p.name.endswith((".tsv", ".tsv.gz"))
    and ".measure." not in p.name
    and ".layout." not in p.name
)


# rows character mode skips as glyph slack (Excel fitted more or fewer characters than the width),
# per suite and tier: a known deviation of character mode, counted on every run; a change in a
# count fails until the file is written again with --record-glyph-slack
GLYPH_SLACK_FILE = CAPTURE_DIR / "glyph_slack.json"


def _suite_id(path: Path) -> str:
    name = path.name.split(".")[0]
    return f"windows/{name}" if path.parent.name == "windows" else name


# Rows where Excel's DisplayFormat font color, the capture's source for `expected_color`,
# reports automatic black although the sheet paints the section color (seen on screen:
# B9 zero under [Red]0 is red, B132 'abc' under [Cyan]@ is cyan, B147 the '#' run
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
    # the screen shows text under a one-section colored General in that color (rules.layout.tsv,
    # "'Total' [Blue]General" blue; CAPTURING.md section 2 item 60); a bare color section parses
    # as that General, unverified on screen
    ("'[Red]General'", "'abc'"): _TEXT,
    ("'[Red]'", "'abc'"): _TEXT,
    ("'0;-0;0;[Magenta]@'", "True"): _TEXT,
    ("'[Cyan]@'", "'abc'"): _TEXT,
    ("'[Red]m/d/yy'", "-1234.5"): _HASH,
    ("'[Red][h]:mm'", "-1234.5"): _HASH,
}
# the listed rows whose color no screen capture has shown; the package colors them by the rule
# the screen shows for '[Blue]General' (rules.layout.tsv)
NOT_SEEN_ON_SCREEN = {("'[Red]General'", "'abc'"), ("'[Red]'", "'abc'")}

# File widths Excel converts otherwise than ECMA-376's formula, as (width text, digit width); rows
# at them are skipped, and a listed width that matches fails. None is known (CAPTURING.md section 2
# item 53: the width once listed was 10.7109375, recorded with six digits).
FILE_WIDTH_DEVIATIONS: frozenset[tuple[str, float]] = frozenset()


def _written_width(row: Row, header: dict[str, str]) -> float | None:
    """The file width the workbook held. Suites captured before the ``widths: exact`` header
    recorded it with six significant digits, so a 1/256 step such as 10.7109375 reads as
    "10.7109"; the step whose short text the row holds is the width written."""
    width = row.width_value
    if width is None or header.get("widths") == "exact":
        return width
    step = round(width * 256) / 256
    return step if f"{step:g}" == row.width else width


def _excel_value(row: Row, header: dict[str, str]) -> object:
    """The value Excel loaded for the row.

    Suites captured before 2026-10-24 (no ``numbers`` header) were written by openpyxl with
    ``"%.16g"`` texts, which denote a different double for about half of all floats; the row's
    expected text belongs to that double (CAPTURING.md section 2 item 16). Formula rows always
    denote the Python double.
    """
    value = row.value
    if value == "" and not row.tag.startswith(("formula:", "text:")):
        return None  # openpyxl writes the empty string as a blank cell; these rows hold real ones
    if row.value_type == "error":
        assert isinstance(value, str)
        return ExcelError(value)
    if (
        header.get("numbers") != "repr"
        and row.value_type == "float"
        and not row.tag.startswith("formula:")
    ):
        assert isinstance(value, float)
        return float(f"{value:.16g}")
    return value


_CELL_FONT = re.compile(r" \[(.+?) (\d+(?:\.\d+)?)((?: bold)?)((?: italic)?)\]$")


def _cell_font(tag: str) -> Font | None:
    """The cell's own font from a tag ending in its label, ``[Calibri 14 bold]``; ``None`` for
    ``[normal]`` and untagged rows, which inherit the workbook's Normal font."""
    match = _CELL_FONT.search(tag)
    if not match:
        return None
    return Font(match[1], float(match[2]), bold=bool(match[3]), italic=bool(match[4]))


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


def _clipboard_year(row: Row, width: int | str | None, value: object, date1904: bool) -> bool:
    """Width-less, a Buddhist-year code: the clipboard gives the plain year (a ``03`` where the
    sheet, ``string value`` and ``TEXT`` show ``46``), so the row is compared with the plain year
    code. The package follows the sheet (QUIRKS.md, CAPTURING.md section 2 item 35)."""
    fmt = row.format_string
    if width is not None or "b" not in fmt.lower() or not NumberFormat.parse(fmt).is_date:
        return False
    if not any(c.text[:1] == "b" for s in NumberFormat.parse(fmt).sections for c in s.tokens):
        return False
    plain = re.sub("[bB]+(?![12])", lambda m: "yy" if len(m[0]) <= 2 else "yyyy", fmt)
    return format_value(value, plain, date1904=date1904) == row.expected_text


def _glyph_tight_fill(expected: str, got: str, fmt: str, width: int | None) -> bool:
    """Excel filled fewer characters than the width because the text's glyphs are wider than a
    digit: ``'@*-'`` with ``FALSE`` in Arial 10 on Windows at width 6 has no room for the dash
    that character mode adds. Only counts when the extra characters are the format's fill."""
    if width is None or len(got) <= len(expected) or len(got) > width:
        return False
    extra = Counter(got) - Counter(expected)
    return not Counter(expected) - Counter(got) and all(f"*{c}" in fmt for c in extra)


def _glyph_slack_exponent(expected: str, got: str, width: int | None) -> bool:
    """A reduced scientific General among literals: ``( 1.2346E-99 )`` is 14 characters and
    Excel charges the exponent a third digit, but its literals cost less than a digit each, so
    it keeps a form that character mode, charging 15, shortens."""
    return (
        width is not None
        and "E" in expected
        and len(got) < len(expected) <= width
        and any(c in " ()" for c in expected)
    )


def _glyph_slack_literals(expected: str, got: str, width: int | None) -> bool:
    """Excel fitted a text no longer than the width that character mode overflows, because the
    literals around a General cost less than a digit each ('( ' is 0.83 units, two characters):
    ``( General )`` with -1.5E-07 at width 5 is ``( 0 )`` in Excel, where character mode has no
    room left for the General. Only counts when the text has a glyph narrower than a digit."""
    return (
        width is not None
        and set(got) == {"#"}
        and set(expected) != {"#"}
        and len(expected) <= width
        and any(c in " ()-.,:/" for c in expected)
    )


def _check_glyph_slack(suite: str, tier: str, count: int, record: bool) -> None:
    listed: dict[str, dict[str, int]] = (
        json.loads(GLYPH_SLACK_FILE.read_text(encoding="utf-8"))
        if GLYPH_SLACK_FILE.exists()
        else {}
    )
    if record:
        listed.setdefault(suite, {})[tier] = count
        listed = {k: {t: n for t, n in v.items() if n} for k, v in sorted(listed.items())}
        listed = {k: v for k, v in listed.items() if v}
        GLYPH_SLACK_FILE.write_text(json.dumps(listed, indent=1) + "\n", encoding="utf-8")
        return
    want = listed.get(suite, {}).get(tier, 0)
    if count != want:
        pytest.fail(
            f"{suite} [{tier}]: {count} glyph-slack rows skipped, {want} listed in "
            f"{GLYPH_SLACK_FILE.name}; write it again with --record-glyph-slack if intended"
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
@pytest.mark.parametrize("suite_file", SUITE_FILES, ids=_suite_id)
def test_suite(suite_file: Path, tier: str, record_glyph_slack: bool, record_property) -> None:  # type: ignore[no-untyped-def]
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
        "clipboard plain year": 0,
        "skipped (wide glyphs)": 0,
        "skipped (file width)": 0,
        "skipped (cell font)": 0,
        "Excel mode compared": 0,
        "file widths checked": 0,
        "known deviation (file width)": 0,
        "colors compared": 0,
        "color misreports skipped": 0,
    }
    # "Calibri 11", "Aptos Narrow 11", "Calibri 11 (openpyxl default)"
    match = re.match(r"(.+?) (\d+(?:\.\d+)?)\b", header.get("workbook_font", ""))
    normal = Font(match[1], float(match[2])) if match else Font("Calibri", 11.0)
    platform = header.get("platform", "mac")  # Excel for Windows' captures name it
    # widths written into the file (<col width>, or none: <sheetFormatPr>), not set by the script
    file_widths = header.get("width_source") == "file"
    sheet_format = dict(kv.split("=") for kv in header.get("sheet_format", "").split())
    show_zeros = header.get("show_zeros") != "0"  # a sheet that hides zeros says so
    for row in rows:
        fmt = row.format_string
        value = _excel_value(row, header)
        date1904 = row.date1904 == "1"
        if row.tag.startswith("text:"):  # the cell holds TEXT of the value under the format
            counts["compared"] += 1
            got = text_function(value, fmt, date1904=date1904)
            if got != row.expected_text:
                failures.append((row, got))
            continue
        column_width = row.width_value
        if row.tag.startswith("merged:"):  # the width is the block's total points
            counts["Excel mode compared"] += 1
            cw = column_metrics_from_points(row.width_value or 0, normal, platform=platform).width
            got = format_value(
                value, fmt, column_width=cw, font=normal, normal_font=normal, platform=platform
            )
            if got != row.expected_text:
                excel_failures.append((row, got))
            continue
        if file_widths:
            metrics = file_column_metrics(
                _written_width(row, header),
                base_col_width=float(sheet_format.get("baseColWidth", 8)),
                default_col_width=(
                    float(sheet_format["defaultColWidth"])
                    if "defaultColWidth" in sheet_format
                    else None
                ),
                normal_font=normal,
                platform=platform,
            )
            column_width = metrics.width
            # the conversion itself: the dialog number Excel showed, to two decimals half up
            # (1.125 shows as 1.13)
            shown = Decimal(repr(metrics.width)).quantize(Decimal("0.01"), ROUND_HALF_UP)
            if (row.width, metrics.digit_width) in FILE_WIDTH_DEVIATIONS:
                # a listed width that Excel now converts as the package does is a stale entry
                if shown == Decimal(row.actual_width):
                    excel_failures.append((row, "<listed in FILE_WIDTH_DEVIATIONS, now matches>"))
                counts["known deviation (file width)"] += 1
                continue
            counts["file widths checked"] += 1
            if shown != Decimal(row.actual_width):
                excel_failures.append((row, f"<dialog width {metrics.width:.2f}>"))
        if column_width is not None:
            # Excel mode: the column width as requested (Excel snaps it), the workbook's font;
            # every width row must match exactly, '#' counts included.
            counts["Excel mode compared"] += 1
            cell_font = _cell_font(row.tag)
            try:
                got_excel = format_value(
                    value,
                    fmt,
                    column_width=column_width,
                    font=cell_font or normal,
                    normal_font=normal,
                    date1904=date1904,
                    show_zeros=show_zeros,
                    platform=platform,
                )
            except Exception as exc:
                got_excel = f"<{type(exc).__name__}: {exc}>"
            if got_excel != row.expected_text:
                excel_failures.append((row, got_excel))
        if file_widths:
            counts["skipped (file width)"] += 1  # a file width has no character count
            continue
        if header["suite"].startswith("advance_tenths"):
            counts["skipped (wide glyphs)"] += 1  # glyphs wider than a digit, by design
            continue
        if _cell_font(row.tag) is not None:
            counts["skipped (cell font)"] += 1  # characters of another font than the column's
            continue
        width = _character_width(row.width_value)
        if width == "skip":
            counts["skipped (fractional width)"] += 1
            continue
        if _clipboard_year(row, width, value, date1904):
            counts["clipboard plain year"] += 1
            continue
        if _glyph_slack(row.expected_text, width):
            counts["skipped (glyph slack)"] += 1
            continue
        try:
            got = format_value(value, fmt, width=width, date1904=date1904, show_zeros=show_zeros)
        except Exception as exc:
            got = f"<{type(exc).__name__}: {exc}>"
        if (
            _glyph_tight(row.expected_text, got, width)
            or _glyph_tight_fill(row.expected_text, got, fmt, width)
            or _glyph_slack_literals(row.expected_text, got, width)
            or _glyph_slack_exponent(row.expected_text, got, width)
        ):
            counts["skipped (glyph slack)"] += 1
            continue
        counts["compared"] += 1
        if not _texts_match(got, row.expected_text, width):
            failures.append((row, got))
        if row.expected_color and (row.format, row.value_repr) in DISPLAY_FORMAT_COLOR_MISREPORTS:
            # strict: a listed row Excel now reports as the package colors it is a stale entry
            counts["color misreports skipped"] += 1
            cell = format_cell(value, fmt, date1904=date1904)
            ours = DEFAULT_PALETTE[cell.color] if cell.color else "000000"
            if ours == row.expected_color:
                color_failures.append((row, f"{ours} (listed as misreported, now matches)"))
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
    _check_glyph_slack(
        _suite_id(suite_file), tier, counts["skipped (glyph slack)"], record_glyph_slack
    )


def test_color_misreports_are_all_used() -> None:
    """Every listed color misreport names a captured row: an unused entry is stale."""
    used: set[tuple[str, str]] = set()
    for path in SUITE_FILES:
        header, rows = read_rows(path)
        if "colors=" not in header.get("method", ""):
            continue
        used.update((r.format, r.value_repr) for r in rows if r.expected_color)
    unused = sorted(set(DISPLAY_FORMAT_COLOR_MISREPORTS) - used)
    assert not unused, f"color misreports no captured row has: {unused}"
