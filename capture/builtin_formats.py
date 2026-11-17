"""What Excel for Mac shows for each built-in number format id (a probe).

    uv run python -m capture.builtin_formats build   # the workbooks, no Excel
    uv run python -m capture.builtin_formats run     # open each, read, write the TSV
    uv run python -m capture.builtin_formats run LABEL   # the date, time and currency ids only

A cell can name a number format by id alone (``<xf numFmtId="14">`` with no ``<numFmt>``): Excel
writes id 14 for every date typed as a short date, and the format string it stands for is the
reader's to know. Each row here is one id from 0 to 163 (164 is the first custom id) under seven
values that tell the formats apart: a date and time (2024-01-01 13:30), 1234.5678, -1234.5678,
0.25, 0, 1.5 and a text. openpyxl writes only the ids it knows, so every row gets a placeholder
custom format, and the saved file's ``styles.xml`` is rewritten so that the row's style names the
built-in id and the placeholder is gone. Excel reads back each row's ``number format`` and
``number format local`` and each value cell's displayed text at a column width of 40.

An id Excel does not take can make it refuse the whole file, so the ids go in groups, and a group
Excel refuses is split until each refused id stands alone. No clipboard. Writes
``tests/captures/builtin_formats.layout.tsv``, or with a label ``capture/build/probes/``; refused
ids are rows with ``refused`` in place of the format.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
import zipfile
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript

SCRIPT = Path(__file__).with_name("builtin_formats.applescript")
OUTPUT = BUILD_DIR.parent.parent / "tests" / "captures" / "builtin_formats.layout.tsv"
DIRECTORY = WORKBOOK_DIR / "builtin_formats"
IDS = tuple(range(164))
VALUES: tuple[tuple[str, object], ...] = (
    ("datetime", 45292.5625),
    ("number", 1234.5678),
    ("negative", -1234.5678),
    ("quarter", 0.25),
    ("zero", 0),
    ("day and a half", 1.5),
    ("text", "abc"),
)
GROUPS = (range(0, 50), range(50, 82), range(82, 123), range(123, 164))
US = "\x1f"


def _placeholder(ident: int) -> str:
    return f'"__id_{ident}__"0'


def build(ids: list[int], name: str) -> Path:
    """One row per id: its label in A, the values in B onward under that built-in id."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "builtin"
    for r, ident in enumerate(ids, 1):
        ws.cell(row=r, column=1, value=f"id {ident}").data_type = "s"
        for c, (_, value) in enumerate(VALUES, 2):
            cell = ws.cell(row=r, column=c, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            cell.number_format = _placeholder(ident)
    DIRECTORY.mkdir(parents=True, exist_ok=True)
    path = (DIRECTORY / name).resolve()
    wb.save(path)
    _point_at_builtins(path)
    return path


def _point_at_builtins(path: Path) -> None:
    """Rewrite styles.xml: each placeholder's style names its built-in id; no custom formats."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    styles = parts["xl/styles.xml"].decode()
    placeholders = r'<numFmt numFmtId="(\d+)" formatCode="&quot;__id_(\d+)__&quot;0"\s*/>'
    mapping = dict(re.findall(placeholders, styles))
    if not mapping:
        sys.exit(f"{path.name}: no placeholder formats found in styles.xml")
    styles = re.sub(r"<numFmts\b.*?</numFmts>", "", styles, flags=re.S)
    styles = re.sub(
        r'numFmtId="(\d+)"', lambda m: f'numFmtId="{mapping.get(m.group(1), m.group(1))}"', styles
    )
    if "__id_" in styles:
        sys.exit(f"{path.name}: a placeholder survived the rewrite")
    parts["xl/styles.xml"] = styles.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def _read(ids: list[int], tag: str) -> dict[int, list[str]] | None:
    """Open the ids' workbook; their rows, or None when Excel refuses the file."""
    name = f"builtin_{tag}.xlsx"
    path = build(ids, name)
    proc = osascript(
        str(SCRIPT), str(path), name, str(len(ids)), str(len(VALUES)), seconds=60.0 + 2 * len(ids)
    )
    if proc.returncode != 0:
        print(f"  {tag}: refused ({proc.stderr.strip()[-80:]})")
        return None
    lines = proc.stdout.rstrip("\n").split("\n")
    if len(lines) != len(ids):
        sys.exit(f"{tag}: {len(lines)} rows read, expected {len(ids)}")
    return {ident: line.split(US) for ident, line in zip(ids, lines, strict=True)}


def _capture(ids: list[int], tag: str, out: dict[int, list[str] | None]) -> None:
    """Read a group; split a refused one until each refused id stands alone."""
    rows = _read(ids, tag)
    if rows is not None:
        out.update(rows)
    elif len(ids) == 1:
        out[ids[0]] = None
    else:
        half = len(ids) // 2
        _capture(ids[:half], f"{tag}a", out)
        _capture(ids[half:], f"{tag}b", out)


LOCALE_IDS = (*range(5, 9), *range(14, 23), *range(37, 48))
"""The ids the locale runs check: dates and times (14 to 22, 45 to 47) and the currency and
accounting ids (5 to 8, 37 to 44); on the Mac only 14 to 16 and 22 follow a setting (CAPTURING.md
section 2 item 49)."""


def run(label: str = "") -> None:
    """Capture every id, or with ``label`` the ``LOCALE_IDS`` under the current macOS settings."""
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    found: dict[int, list[str] | None] = {}
    groups = [LOCALE_IDS] if label else [tuple(g) for g in GROUPS]
    for k, group in enumerate(groups):
        _capture(list(group), f"{label}g{k}", found)
        print(f"group of {len(group)} ids: done")
    out = BUILD_DIR / "probes" / f"builtin_formats_{label}.tsv" if label else OUTPUT
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        f.write(
            "# suite: builtin_formats (capture.builtin_formats); Excel for Mac, en_US; a probe\n"
        )
        f.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        f.write(
            f"# captured: {dt.date.today().isoformat()}; method: <xf numFmtId> with no <numFmt>, "
            "string value at column width 40, number format and number format local of the cell; "
            "values: " + ", ".join(f"{label}={value!r}" for label, value in VALUES) + "\n"
        )
        f.write("id\tnumber_format\tnumber_format_local\t" + "\t".join(v for v, _ in VALUES) + "\n")
        for ident in sorted(found):
            row = found.get(ident)
            fields = ["refused"] * (2 + len(VALUES)) if row is None else row
            f.write("\t".join([str(ident), *(repr(x) for x in fields)]) + "\n")
    refused = [i for i in sorted(found) if found[i] is None]
    print(f"wrote {out}: {len(found)} ids, refused {refused}")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        for k, group in enumerate(GROUPS):
            print(build(list(group), f"builtin_g{k}.xlsx"))
    elif command == "run":
        run(argv[1] if len(argv) > 1 else "")
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
