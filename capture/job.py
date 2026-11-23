"""Capture suites as Excel capture jobs, and the suite files their results give.

    uv run python -m capture.job write GROUP JOB_ID OUT_DIR [PURPOSE_NOTE]
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/e3a5207c251929a8a19d713777b70ace3506fb95/docs/capture-jobs.md
A group (``GROUPS``) is a list of suites run as one job, one workbook each. ``write`` makes the job
folder: the manifest (``job.json``) and each suite's workbook in ``files/``, built as
``capture/build.py`` builds it for Excel for Mac, with its zip entries' times and document
properties fixed, so that writing it again gives the same bytes and the SHA-256 the result
records. Each workbook's steps do in Excel what ``capture/capture_template.applescript`` does on
the Mac: a width-less column is copied as text (``copy_text``); a column at a width is set to it
(``set_size``, Excel's character units) and its cells' text read (``read``, the cell's displayed
text, one cell at a time), with the column's width read back in units and points; a column
whose width the file sets is read as it opens; a measure sheet is autofitted and each column's
width read; colors are the display format's font color of each width-less cell.

``read`` checks the manifest's and the inputs' SHA-256 against the result's and writes, per suite,
``tests/captures/windows/<suite>.tsv`` (gzipped above 50 KB) in the schema of the Mac's suite
files, and ``<suite>.measure.tsv`` for a measure suite. The header names the group's manifest,
copied beside the files byte for byte as ``<group>.job.json``, and holds the result's ``header``
whole, as one line of JSON (``job_header``).
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from capture.build import Manifest, build
from capture.capture import make_rows, write_suite
from capture.cases import SUITES, Suite, distinct_measure_cases
from capture.schema import COLUMNS
from capture.tiers import assign_tiers, rule_description

ROOT = Path(__file__).resolve().parents[1]
WINDOWS = ROOT / "tests" / "captures" / "windows"
FIXED_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FORMAT_URL = (
    "https://github.com/sheetmill/sheetmill/blob/e3a5207c251929a8a19d713777b70ace3506fb95/"
    "docs/capture-jobs.md"
)


def _family(name: str) -> str:
    return next((f for f in ("aptos_narrow", "arial") if f in name), "calibri")


_TABLES = tuple(name for name in SUITES if re.match(r"(?:measure|space_fill)(?:\d|_)", name))
GROUPS: dict[str, tuple[str, ...]] = {
    # the glyph tables: autofit widths of every measured face, and the space by fill counts
    **{
        f"tables_{family}": tuple(n for n in _TABLES if _family(n) == family)
        for family in ("calibri", "aptos_narrow", "arial")
    },
    # the rules: columns without a width, file widths at 1/256 steps, fills past 255
    # characters, file widths at every size, cell fonts against the Normal font, error values
    "rules": tuple(
        name
        for name in SUITES
        if re.fullmatch(
            r"column_default_\w+|column_boundary\d+|long_fills|column_file\w*"
            r"|cell_font_styles|cell_font_large|cell_font_arial|error_width",
            name,
        )
    ),
    # validation: rows the tables were not fitted on, in two jobs of about ten minutes
    "validation_numbers": (
        "numeric_core_width",
        "numeric_core_width_aptos_narrow11",
        "numeric_core_width_arial10",
        "cell_font_numeric",
    ),
    "validation_general": tuple(
        name
        for name in SUITES
        if re.fullmatch(
            r"general_points\w+|general_literals_points|general_zero_sign|general_cell_font\w*"
            r"|cell_font|column_narrow\d+|general_text_narrow|general_sign_narrow|width_fill"
            r"|width_general\w*|huge_values|letters|letters_upper|elapsed_decimals",
            name,
        )
    ),
}


def _cell(sheet: str, letter: str, row: int) -> str:
    return f"{sheet}!{letter}{row}"


def _steps(suite: Suite, manifest: Manifest) -> list[dict[str, Any]]:
    """The steps of one suite's workbook: what the Mac's capture script does, as operations."""
    if manifest.date1904:
        msg = f"{suite.name}: a 1904 workbook is not written as a job"
        raise ValueError(msg)
    if suite.merged:
        msg = f"{suite.name}: merged columns are not written as a job"
        raise ValueError(msg)
    n = manifest.n_rows
    steps: list[dict[str, Any]] = [{"op": "open"}]
    for spec in manifest.columns:
        col = spec.letter
        cells = f"cases!{col}1:{col}{n}"
        key = f"{col.lower()}"
        if spec.method == "clipboard":
            steps.append({"op": "copy_text", "id": f"clip_{key}", "range": cells})
            continue
        if spec.method == "string_value":
            steps.append({"op": "set_size", "columns": f"cases!{col}:{col}", "width": spec.width})
        steps.append({"op": "read", "id": f"text_{key}", "range": cells, "props": ["text"]})
        steps.append({
            "op": "read", "id": f"size_{key}", "range": _cell("cases", col, 1),
            "props": ["column_width", "width"],
        })  # fmt: skip
    if manifest.colors:
        steps.append({
            "op": "read_format", "id": "colors", "range": f"cases!B1:B{n}",
            "props": ["font.color"], "display": True,
        })  # fmt: skip
    if manifest.measure_cells:
        first = manifest.measure_cells[0].rstrip("0123456789")
        last = manifest.measure_cells[-1].rstrip("0123456789")
        steps.append({"op": "autofit", "id": "autofit", "columns": f"measure!{first}:{last}"})
        steps.append({
            "op": "read", "id": "measure", "range": f"measure!{first}1:{last}1",
            "props": ["column_width", "width"],
        })  # fmt: skip
    steps.append({"op": "close", "save": False})
    return steps


def _fixed_copy(raw: Path, path: Path) -> None:
    """``raw`` rewritten to ``path`` with fixed zip times and document dates."""
    with zipfile.ZipFile(raw) as source, zipfile.ZipFile(path, "w") as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "docProps/core.xml":
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)",
                    rf"\g<1>{FIXED_TIME}\2",
                    data.decode("utf-8"),
                )
                data = text.encode("utf-8")
            entry = zipfile.ZipInfo(info.filename, date_time=ZIP_TIME)
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o600 << 16
            out.writestr(entry, data)


def _built(name: str, files: Path) -> tuple[Suite, Manifest]:
    suite = SUITES[name]()
    with tempfile.TemporaryDirectory() as tmp:
        manifest = build(suite, Path(tmp))
        _fixed_copy(Path(manifest.workbook), files / f"{name}.xlsx")
    return suite, manifest


def manifest_for(group: str, job_id: str, steps: list[dict[str, Any]], note: str) -> dict[str, Any]:
    purpose = f"xlformat capture group {group}: displayed text, widths and autofit per suite"
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "xlformat",
        "purpose": purpose + (f"; {note}" if note else ""),
        "limits": {"job_seconds": 3600, "step_seconds": 600},
        "audit": False,
        "workbooks": steps,
    }


def write(group: str, job_id: str, out: Path, note: str = "") -> Path:
    """The job folder ``out/<job_id>``."""
    folder = out / job_id
    files = folder / "files"
    files.mkdir(parents=True)
    workbooks = []
    for name in GROUPS[group]:
        suite, manifest = _built(name, files)
        workbooks.append({"file": f"{name}.xlsx", "steps": _steps(suite, manifest)})
    doc = manifest_for(group, job_id, workbooks, note)
    (folder / "job.json").write_text(json.dumps(doc, indent=1) + "\n", encoding="utf-8")
    return folder


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _hex(bgr: object) -> str:
    """Excel for Windows' ``Font.Color``, a BGR long, as the Mac's suites write it: ``RRGGBB``."""
    if not isinstance(bgr, (int, float)):
        msg = f"font color {bgr!r}"
        raise ValueError(msg)
    v = int(bgr)
    return f"{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


def _number(v: object) -> str:
    assert isinstance(v, (int, float)), v
    return str(float(v))


def _clipboard(text: str | None, n: int) -> list[str]:
    cells = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if len(cells) == n + 1 and cells[-1] == "":
        cells.pop()
    return cells


def _results(wb: dict[str, Any]) -> dict[str, Any]:
    if wb.get("status") != "ok":
        msg = f"{wb.get('file')}: workbook status {wb.get('status')!r}"
        raise ValueError(msg)
    out = {}
    for step in wb["steps"]:
        if step.get("status") != "ok":
            msg = f"{wb.get('file')}: step {step.get('op')} {step.get('id')}: {step.get('error')}"
            raise ValueError(msg)
        if "id" in step:
            out[step["id"]] = step.get("result")
    return out


def _suite_files(
    name: str, group: str, results: dict[str, Any], header: dict[str, Any]
) -> list[Path]:
    suite = SUITES[name]()
    with tempfile.TemporaryDirectory() as tmp:
        manifest = build(suite, Path(tmp))
    n = manifest.n_rows
    columns: dict[str, list[str]] = {}
    meta: dict[str, str] = {}
    for spec in manifest.columns:
        key = spec.letter.lower()
        if spec.method == "clipboard":
            columns[spec.letter] = _clipboard(results[f"clip_{key}"]["text"], n)
        else:
            columns[spec.letter] = [row[0] for row in results[f"text_{key}"]["text"]]
            size = results[f"size_{key}"]
            meta[f"actual_width_{spec.letter}"] = _number(size["column_width"][0][0])
            meta[f"points_{spec.letter}"] = _number(size["width"][0][0])
        if len(columns[spec.letter]) != n:
            msg = f"{name} column {spec.letter}: {len(columns[spec.letter])} cells, not {n}"
            raise ValueError(msg)
    colors = None
    if manifest.colors:
        colors = [_hex(row[0]["font.color"]) for row in results["colors"]["cells"]]
    rows = assign_tiers(make_rows(suite.cases, manifest, columns, meta, colors), suite.fast_cap)
    env = header.get("environment", {})
    excel = env.get("excel", {})
    head = {
        "suite": name,
        "platform": "windows",
        "excel_build": str(excel.get("build", "")),
        "job": f"{group}.job.json",
        "job_id": str(header.get("job_id", "")),
        "workbook_font": f"{suite.font_name} {suite.font_size or 11}",
        "captured": dt.date.today().isoformat(),
        "generator": f"capture.cases.{name}, capture/job.py",
        "method": "; ".join(
            f"width {'none' if c.width is None else c.width}={c.method}" for c in manifest.columns
        )
        + ("; colors=display format font color per cell" if manifest.colors else ""),
        "numbers": "repr",
        "widths": "exact",
        "column_points": " ".join(
            f"{c.letter}={meta[f'points_{c.letter}']}"
            for c in manifest.columns
            if f"points_{c.letter}" in meta
        ),
        "tier_rule": rule_description(suite.fast_cap),
        **({"width_source": "file"} if suite.width_source == "file" else {}),
        **(
            {"sheet_format": " ".join(f"{k}={v:g}" for k, v in suite.sheet_format)}
            if suite.sheet_format
            else {}
        ),
        **({"show_zeros": "0"} if not suite.show_zeros else {}),
        "job_header": json.dumps(header, sort_keys=True, ensure_ascii=False),
        "columns": " ".join(COLUMNS),
    }
    written = [write_suite(WINDOWS / f"{name}.tsv", head, rows)]
    if manifest.measure_cells:
        measured = results["measure"]
        widths = measured["column_width"][0]
        points = measured["width"][0]
        entries = distinct_measure_cases(suite.cases)
        fonts = any(c.font is not None for c in entries)
        path = WINDOWS / f"{name}.measure.tsv"
        with path.open("w", encoding="utf-8") as f:
            f.write(
                "# autofit column width (Excel units) per distinct (value, format); font: "
                f"{head['workbook_font']}; job: {group}.job.json, {head['job_id']}\n"
            )
            f.write(
                "value_type\tvalue_repr\tformat\tautofit_width\tautofit_points"
                + ("\tcell_font\n" if fonts else "\n")
            )
            for case, w, pt in zip(entries, widths, points, strict=True):
                font = ("\t" + (case.font.label if case.font else "")) if fonts else ""
                f.write(
                    f"{case.value_type}\t{case.value_repr}\t{case.format!r}\t{_number(w)}"
                    f"\t{_number(pt)}{font}\n"
                )
        written.append(path)
    return written


def read(group: str, job: Path, result: Path) -> list[Path]:
    """The suite files from a job folder and its result folder (``result.json``)."""
    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        sys.exit(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        sys.exit("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            sys.exit(f"the result's input {name} is not this job's")
    WINDOWS.mkdir(parents=True, exist_ok=True)
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    written: list[Path] = []
    for name in GROUPS[group]:
        written += _suite_files(name, group, _results(by_file[f"{name}.xlsx"]), header)
    shutil.copyfile(job / "job.json", WINDOWS / f"{group}.job.json")
    return written


def main(argv: list[str]) -> None:
    if argv[:1] == ["write"] and len(argv) in (4, 5):
        print(write(argv[1], argv[2], Path(argv[3]), argv[4] if len(argv) == 5 else ""))
    elif argv[:1] == ["read"] and len(argv) == 4:
        for path in read(argv[1], Path(argv[2]), Path(argv[3])):
            print(path)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
