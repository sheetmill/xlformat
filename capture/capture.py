"""Capture a suite from Excel for Mac and write ``tests/captures/<suite>.tsv``.

Usage: ``uv run python -m capture <suite> [--diff-2023]``.

Steps: build the workbook (``capture.build``), fill the AppleScript template, run it with
``osascript`` (needs macOS Automation permission for Excel), parse the output, assign tiers,
write the TSV with a metadata header. ``--diff-2023`` compares the captured text with the 2023
capture for identical (value, format) pairs and prints every difference.

The clipboard is saved before and restored after the run (text only).
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import gzip
import io
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

from capture.build import BUILD_DIR, Manifest, build
from capture.cases import SUITES, Case, distinct_measure_cases
from capture.schema import COLUMNS, Row, read_rows, write_rows
from capture.tiers import assign_tiers, rule_description

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "capture" / "capture_template.applescript"
OUT_DIR = ROOT / "tests" / "captures"
PROBE_DIR = BUILD_DIR / "probes"  # probe suites: evidence for a model not yet built
RS, US = "\x1e", "\x1f"


def _applescript_text(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def render_script(manifest: Manifest) -> str:
    """Fill the template for one manifest."""
    cols = ", ".join(
        "{"
        + ", ".join([_applescript_text(c.letter), str(c.width or 0), _applescript_text(c.method)])
        + "}"
        for c in manifest.columns
    )
    cells = ", ".join(_applescript_text(c) for c in manifest.measure_cells)
    return (
        TEMPLATE.read_text()
        .replace("{{PATH}}", manifest.workbook.replace("\\", "\\\\").replace('"', '\\"'))
        .replace("{{NAME}}", Path(manifest.workbook).name)
        .replace("{{NROWS}}", str(manifest.n_rows))
        .replace("{{COLUMNS}}", "{" + cols + "}")
        .replace("{{MEASURE_CELLS}}", "{" + cells + "}")
        .replace("{{DATE1904}}", "true" if manifest.date1904 else "false")
        .replace("{{COLORS}}", "true" if manifest.colors else "false")
    )


def _pb(cmd: str, data: str | None = None) -> str:
    return subprocess.run([cmd], input=data, capture_output=True, text=True, check=False).stdout


def run_osascript(script_path: Path) -> str:
    """Run the script; return its stdout. Raises on a non-zero exit."""
    saved = _pb("pbpaste")
    try:
        proc = subprocess.run(
            ["osascript", str(script_path)], capture_output=True, text=True, check=False
        )
    finally:
        _pb("pbcopy", saved)
    if proc.returncode != 0:
        msg = f"osascript failed ({proc.returncode}):\n{proc.stderr}"
        raise RuntimeError(msg)
    return proc.stdout.rstrip("\n")


def _hex_color(rgb: str) -> str:
    """``"0,0,255"`` as reported by Excel to ``"0000FF"``."""
    r, g, b = (int(part) for part in rgb.split(","))
    return f"{r:02X}{g:02X}{b:02X}"


def parse_output(
    text: str, manifest: Manifest
) -> tuple[dict[str, str], dict[str, list[str]], list[str], list[str]]:
    """Split the script output into metadata, per-column texts, measure widths and colors."""
    meta: dict[str, str] = {}
    columns: dict[str, list[str]] = {}
    measure: list[str] = []
    colors: list[str] = []
    for record in text.split(RS):
        fields = record.split(US)
        kind = fields[0]
        if kind == "column":
            letter, actual_width, *cells = fields[1:]
            meta[f"actual_width_{letter}"] = actual_width
            columns[letter] = cells
        elif kind == "clipboard":
            letter, b64 = fields[1], fields[2]
            raw = base64.b64decode(b64).decode("utf-8")
            cells = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            if len(cells) == manifest.n_rows + 1 and cells[-1] == "":
                cells.pop()
            columns[letter] = cells
        elif kind == "measure":
            measure = fields[1:]
        elif kind == "measure_points":
            meta["measure_points"] = US.join(fields[1:])
        elif kind == "colpoints":
            meta[f"points_{fields[1]}"] = fields[2]
        elif kind == "colors":
            colors = [_hex_color(c) for c in fields[1:]]
            if len(colors) != manifest.n_rows:
                msg = f"colors: expected {manifest.n_rows} cells, got {len(colors)}"
                raise ValueError(msg)
        else:
            meta[kind] = fields[1] if len(fields) > 1 else ""
    for letter, cells in columns.items():
        if len(cells) != manifest.n_rows:
            msg = f"column {letter}: expected {manifest.n_rows} cells, got {len(cells)}"
            raise ValueError(msg)
    return meta, columns, measure, colors


def make_rows(
    cases: tuple[Case, ...],
    manifest: Manifest,
    columns: dict[str, list[str]],
    meta: dict[str, str],
    colors: list[str] | None = None,
) -> list[Row]:
    """One TSV row per (case, capture column); colors go on the width-less rows."""
    rows: list[Row] = []
    for spec in manifest.columns:
        cells = columns[spec.letter]
        actual = meta.get(f"actual_width_{spec.letter}", "")
        row_colors = colors if (colors and spec.width is None) else [""] * len(cells)
        for case, text, color in zip(cases, cells, row_colors, strict=True):
            rows.append(
                Row(
                    suite=case.suite,
                    tag=case.tag,
                    value_type=case.value_type,
                    value_repr=case.value_repr,
                    format=repr(case.format),
                    width="" if spec.width is None else f"{spec.width:g}",
                    actual_width=actual,
                    date1904="1" if case.date1904 else "0",
                    expected_text=text,
                    expected_color=color,
                    tier="full",
                )
            )
    return rows


GZIP_ABOVE = 50_000  # bytes; larger suite files are stored gzipped


def write_suite(path: Path, header: dict[str, str], rows: list[Row]) -> Path:
    """Write a suite file, gzipped when it would exceed ``GZIP_ABOVE`` bytes; return the path."""
    buffer = io.StringIO()
    write_rows(buffer, header, rows)
    text = buffer.getvalue()
    plain = path.with_suffix("") if path.suffix == ".gz" else path
    gz = plain.with_suffix(plain.suffix + ".gz")
    if len(text.encode("utf-8")) > GZIP_ABOVE:
        with gzip.open(gz, "wt", newline="", encoding="utf-8") as g:
            g.write(text)
        plain.unlink(missing_ok=True)
        return gz
    plain.write_text(text, encoding="utf-8", newline="")
    gz.unlink(missing_ok=True)
    return plain


def capture(suite_name: str) -> tuple[Path, dict[str, str]]:
    """Build, capture and write one suite. Returns the TSV path and the metadata."""
    suite = SUITES[suite_name]()
    manifest = build(suite)
    script_path = BUILD_DIR / f"{suite.name}.capture.applescript"
    script_path.write_text(render_script(manifest))
    print(f"built {manifest.workbook} ({manifest.n_rows} rows, {len(manifest.columns)} columns)")
    print(f"running {script_path}")
    started = dt.datetime.now()
    output = run_osascript(script_path)
    print(f"osascript finished in {(dt.datetime.now() - started).total_seconds():.1f} s")
    meta, columns, measure, colors = parse_output(output, manifest)
    rows = assign_tiers(make_rows(suite.cases, manifest, columns, meta, colors), suite.fast_cap)
    header = {
        "suite": suite.name,
        "excel_version": meta.get("excel_version", ""),
        "standard_font": meta.get("standard_font", ""),
        "workbook_font": meta.get("workbook_font", ""),
        "locale": "en_US",
        "captured": dt.date.today().isoformat(),
        "generator": f"capture.cases.{suite_name}",
        "method": "; ".join(
            f"width {'none' if c.width is None else c.width:}={c.method}" for c in manifest.columns
        )
        + ("; colors=display format font color per cell" if manifest.colors else ""),
        "numbers": "repr",  # float texts are repr(value); absent in older suites ("%.16g")
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
        "columns": " ".join(COLUMNS),
    }
    out_dir = PROBE_DIR if suite.probe else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out = write_suite(out_dir / f"{suite.name}.tsv", header, rows)
    if measure:
        measure_path = out_dir / f"{suite.name}.measure.tsv"
        with measure_path.open("w", encoding="utf-8") as f:
            f.write(
                "# autofit column width (Excel units) per distinct (value, format); font: "
                f"{meta.get('workbook_font', '')}\n"
            )
            raw_points = meta.get("measure_points", "")
            points = raw_points.split(US) if raw_points else [""] * len(measure)
            entries = distinct_measure_cases(suite.cases)
            # a cell_font column only when some cell has its own font (glyph_fit reads five)
            fonts = any(c.font is not None for c in entries)
            f.write(
                "value_type\tvalue_repr\tformat\tautofit_width\tautofit_points"
                + ("\tcell_font\n" if fonts else "\n")
            )
            for case, w, pt in zip(entries, measure, points, strict=True):
                font = ("\t" + (case.font.label if case.font else "")) if fonts else ""
                f.write(f"{case.value_type}\t{case.value_repr}\t{case.format!r}\t{w}\t{pt}{font}\n")
        print(f"wrote {measure_path}")
    fast = sum(1 for r in rows if r.tier == "fast")
    print(f"wrote {out}: {len(rows)} rows, {fast} fast")
    return out, meta


def diff_2023(path: Path) -> int:
    """Compare a capture with the 2023 rows for the same (value, format); print differences."""
    _, new_rows = read_rows(path)
    _, old_rows = read_rows(OUT_DIR / "capture2023.tsv.gz")
    old = {(r.value_type, r.value_repr, r.format): r.expected_text for r in old_rows}
    by_width: dict[str, list[tuple[Row, str]]] = defaultdict(list)
    compared: dict[str, int] = defaultdict(int)
    for r in new_rows:
        key = (r.value_type, r.value_repr, r.format)
        if key not in old:
            continue
        compared[r.width] += 1
        if old[key] != r.expected_text:
            by_width[r.width].append((r, old[key]))
    for width, diffs in sorted(by_width.items()):
        print(f"-- width {width or 'none'}: {len(diffs)} differences of {compared[width]} compared")
        for r, old_text in diffs[:400]:
            print(
                f"   {r.format:50s} {r.value_type} {r.value_repr:28s} "
                f"2023={old_text!r} new={r.expected_text!r}"
            )
    total = sum(len(d) for d in by_width.values())
    print(f"total differences: {total} of {sum(compared.values())} compared")
    return total


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    parser = argparse.ArgumentParser(prog="python -m capture", description=__doc__)
    parser.add_argument("suite", choices=sorted(SUITES))
    parser.add_argument("--diff-2023", action="store_true")
    args = parser.parse_args(argv)
    out, _ = capture(args.suite)
    if args.diff_2023:
        diff_2023(out)


if __name__ == "__main__":
    main(sys.argv[1:])
