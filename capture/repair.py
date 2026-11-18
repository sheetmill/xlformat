"""What Excel's repair does to a cell whose number format Excel refuses in a file.

    uv run python -m capture.repair open all|single     # build the case, open it as a person does
    uv run python -m capture.repair read all|single [--report TEXT]   # after the dialog is answered
    uv run python -m capture.repair write              # write tests/captures/repair.layout.tsv

The one capture that needs a person at Excel. Excel refuses a file holding such a format
(``Parameter error (-50)`` with alerts off, CAPTURING.md section 2 item 9), and with alerts on it
offers to repair it. Case ``all`` holds every refused format known, one per row, between rows of
valid custom formats as controls; case ``single`` holds one refused format, to see whether the
repair treats it the same alone.

``open`` checks that Excel answers, holds no workbook and has alerts on, writes the case into
Excel's container and opens it through Launch Services, as a person opening the file does, then
exits: nothing is sent to Excel while its dialog may be open (a request sent while a person
answers a dialog has crashed Excel). Answer the repair dialog ("Yes") and close the repair
report, then run ``read``, which finds the workbook by the names Excel reports (a
repaired workbook keeps its file's name; its window's caption says "Repaired"), reads every case
cell's displayed text and number format, closes it without saving and checks that Excel holds no
workbook and has alerts on. If no workbook matches it prints that and leaves Excel as it is. The
repair report's text, which no script can read, may be given with ``--report``; only its records
are kept, paths removed. ``write`` gathers both cases into the TSV (CAPTURING.md section 2
item 33). No clipboard.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles.numbers import BUILTIN_FORMATS_REVERSE

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript_or_exit, preflight

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "repair.layout.tsv"
SCRIPT = Path(__file__).with_name("repair.applescript")
STATE = BUILD_DIR / "repair_cases.json"  # the cases read so far (capture/build: not committed)
US = "\x1f"
VALUES: tuple[object, ...] = (1234.5678, -0.5, 0, "abc")

REFUSED: tuple[str, ...] = (
    "0.00e+00", "0.00e-00", "##0.0e+0", "0.00E+00;0.00e+00", "0;0;0;0;0", "[Red", '0.00"',
    "[Red][Blue]0", "[>1][<5]0", "@;@", "0;@;0", "yyyy 0", "0 yyyy", "0.00E+", "# ?/",
    "[hh]:mm:[ss]", "General General",
    "[Color57]0", "[Color0]0", "[]0", "[abc]0", "ss.0000", "yyyy @", "0 @", "@ 0", "# ?/?/?",
    "[>1]0;[>2]0;[>3]0", "[>1]0;[>2]0;[>3]0;@", "0;0;0;[>1]@",
)  # fmt: skip
CONTROLS: tuple[str, ...] = ("0.000", "[Red]0.00", '"x"0', "#,##0.0", "m/d/yy")


def rows_all() -> list[tuple[str, str]]:
    """(tag, format) per row: a control, then each refused format followed by a control."""
    out = [("control", CONTROLS[0])]
    for k, fmt in enumerate(REFUSED):
        out.append(("refused", fmt))
        out.append(("control", CONTROLS[(k + 1) % len(CONTROLS)]))
    return out


def rows_single() -> list[tuple[str, str]]:
    return [("control", CONTROLS[0]), ("refused", "0.00e+00"), ("control", CONTROLS[1])]


CASES = {"all": ("repair_all.xlsx", rows_all), "single": ("repair_single.xlsx", rows_single)}


def build(name: str, rows: list[tuple[str, str]]) -> Path:
    BUILTIN_FORMATS_REVERSE.clear()
    wb = Workbook()
    ws = wb.active
    ws.title = "repair"
    for r, (tag, fmt) in enumerate(rows, 1):
        ws.cell(row=r, column=1, value=f"{tag} {fmt}").data_type = "s"
        for c, value in enumerate(VALUES, 2):
            cell = ws.cell(row=r, column=c, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            cell.number_format = fmt
    WORKBOOK_DIR.mkdir(parents=True, exist_ok=True)
    path = (WORKBOOK_DIR / name).resolve()
    wb.save(path)
    return path


def _ask(script: str) -> str:
    """One short question to Excel (never while a person's dialog may be open)."""
    proc = osascript_or_exit("-e", f'tell application "Microsoft Excel" to {script}', seconds=20.0)
    if proc.returncode != 0:
        sys.exit(f"Excel did not answer {script!r}: {proc.stderr.strip()[:160]}")
    return proc.stdout.strip()


def _clean() -> str:
    """Whether Excel holds no workbook and has alerts on, as a line to print."""
    count, alerts = _ask("get {count of workbooks, display alerts}").split(", ")
    state = "clean" if count == "0" and alerts == "true" else "NOT clean"
    return f"Excel is {state}: {count} workbooks, display alerts {alerts}"


def open_case(case: str) -> None:
    """Check Excel, build the case and open it through Launch Services; then exit."""
    preflight()
    if not answers():
        sys.exit("Excel does not answer a question about its workbooks: nothing opened")
    if _ask("get count of workbooks") != "0":
        names = _ask("get name of every workbook")
        sys.exit(f"Excel holds workbooks ({names}): close them first; nothing opened")
    if _ask("get display alerts") != "true":
        _ask("set display alerts to true")
    name, rows = CASES[case]
    path = build(name, rows())
    subprocess.run(["open", "-a", "Microsoft Excel", str(path)], check=True, timeout=30)
    print(
        f"{case}: opened {name} through Launch Services. Excel asks whether to recover the "
        "workbook's contents: answer Yes, close the repair report (copy its text if wanted), "
        f"then run `read {case}`."
    )


def records(report: str) -> str:
    """The repair report's records (what was repaired or removed), without paths or names."""
    found = re.findall(
        r"<(repairedRecord|removedRecord|removedPart|repairedPart|removedFeature)>(.*?)</\1>",
        report,
        re.S,
    )
    text = " | ".join(f"{kind}: {body.strip()}" for kind, body in found) if found else report
    return " ".join(re.sub(r"/Users/[^'\s<]*", "<path>", text).split())


def read_case(case: str, report: str) -> None:
    """Find the repaired workbook by name, read and close it, and record the case."""
    name, rows = CASES[case]
    answer = _ask("get name of every workbook")  # "missing value" when none is open
    open_names = [n.strip() for n in answer.split(",") if n.strip() and answer != "missing value"]
    stem = name.removesuffix(".xlsx")
    matches = [n for n in open_names if n == name or n.startswith(stem)]
    if len(matches) != 1:
        sys.exit(
            f"{case}: no single workbook named like {name}; Excel has {open_names}. "
            "Excel is left as it is."
        )
    (workbook,) = matches
    proc = osascript_or_exit(
        str(SCRIPT), workbook, str(len(rows())), str(len(VALUES)), seconds=120.0
    )
    if proc.returncode != 0:
        sys.exit(f"reading {workbook} failed: {proc.stderr.strip()[:200]}")
    lines = proc.stdout.rstrip("\n").split("\n")
    windows = lines[0].split(US)[1:]
    cells = [line.split(US)[1:] for line in lines[1:]]
    if len(cells) != len(rows()):
        sys.exit(f"{case}: expected {len(rows())} rows, read {len(cells)}")
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    state[case] = {
        "workbook": workbook,
        "windows": windows,
        "report": records(report),
        "cells": cells,
    }
    STATE.write_text(json.dumps(state, indent=1))
    print(f"{case}: read {workbook} (windows {windows}), {len(cells)} rows. {_clean()}")


def write() -> None:
    """Gather the cases read into ``repair.layout.tsv``."""
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    missing = [case for case in CASES if case not in state]
    if missing:
        sys.exit(f"not read yet: {missing}")
    out: list[list[str]] = []
    for case, (name, rows) in CASES.items():
        cells = state[case]["cells"]
        for (tag, fmt), fields in zip(rows(), cells, strict=True):
            for k, value in enumerate(VALUES):
                out.append([name, tag, repr(fmt), repr(value), fields[2 * k], fields[2 * k + 1]])
    with OUTPUT.open("w", encoding="utf-8", newline="") as f:
        f.write("# suite: repair (capture.repair); Excel for Mac, en_US\n")
        f.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        f.write(
            "# method: opened through Launch Services with alerts on, Excel's repair dialog "
            "answered Yes by a person, then each cell's string value and number format read "
            "through AppleScript and the workbook closed without saving\n"
        )
        for case, (name, _) in CASES.items():
            entry = state[case]
            windows = "; ".join(entry["windows"]) or "none read"
            f.write(f"# {name}: Excel's workbook {entry['workbook']}, window {windows}\n")
            if entry["report"]:
                f.write(f"# {name} repair report: {entry['report']}\n")
        f.write(f"# captured: {dt.date.today().isoformat()}; workbook_font: Calibri 11\n")
        writer = csv.writer(f, dialect="excel-tab", lineterminator="\n")
        writer.writerow(["workbook", "tag", "format", "value_repr", "text", "format_after"])
        writer.writerows(out)
    STATE.unlink()
    print(f"wrote {OUTPUT}: {len(out)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if not argv or argv[0] not in ("open", "read", "write"):
        sys.exit(__doc__)
    command, rest = argv[0], argv[1:]
    if command == "write":
        write()
        return
    report = ""
    if "--report" in rest:
        i = rest.index("--report")
        rest, report = rest[:i], " ".join(rest[i + 1 :])
    if len(rest) != 1 or rest[0] not in CASES:
        sys.exit(f"name a case: {' or '.join(CASES)}")
    if command == "open":
        open_case(rest[0])
    else:
        read_case(rest[0], report)


if __name__ == "__main__":
    main(sys.argv[1:])
