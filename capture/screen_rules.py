"""Where Excel draws underlines and strikethrough, and what color text takes, from the screen.

    uv run python -m capture.screen_rules build     # the workbook, no Excel
    uv run python -m capture.screen_rules predict   # write the predictions, before a run
    uv run python -m capture.screen_rules run       # open, capture every view at zoom 100 and 200
    uv run python -m capture.screen_rules read      # measure the saved captures again

One workbook, Normal style Calibri 11, column A filled magenta and cyan in turn as in
``capture/screen_vertical.py`` (whose AppleScript, window capture and band matching it uses), a
case a row in column B, left-aligned and bottom-aligned, rows fitted by Excel:

- ``rule``: "HHHH" in each face (``FACES``) at 9, 11, 14, 20 and 36 points without a rule and with
  each underline (single, double, single and double accounting) or a strikethrough. Measured at
  zoom 100 and 200: the baseline (the bottom of the H stems' ink) and each rule's first and last
  pixel row, a rule being an ink row whose longest run spans the text.
- ``color``: "Total" and 12 in Calibri 11 under formats with one to four sections and colors.
  Measured: the mean color of the text's colored ink, or ``black``.

The run needs the display on, Excel in front and nobody at the keyboard; no clipboard. Writes
``tests/captures/rules.layout.tsv``; the images stay in ``capture/build/screen_rules/``.
"""

from __future__ import annotations

import datetime as dt
import math
import sys
from dataclasses import dataclass
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, excel_windows
from capture.font_tables import FILES, tables
from capture.screen_vertical import (
    Image,
    _capture,
    _marker_column,
    _marker_right,
    _request,
    _request_count,
    _window_id,
    is_ink,
    read_png,
    rows_in_view,
)

NAME = "screen_rules.xlsx"
PATH = (WORKBOOK_DIR / NAME).resolve()
SHOTS = BUILD_DIR / "screen_rules"
OUTPUT = Path(__file__).resolve().parents[1] / "tests" / "captures" / "rules.layout.tsv"
PREDICTIONS = BUILD_DIR / "screen_rules.predictions.tsv"
FACES = ("Calibri", "Arial", "Aptos Narrow", "Times New Roman", "Courier New", "Georgia")
SIZES = (9, 11, 14, 20, 36)
RULES = ("none", "single", "double", "singleAccounting", "doubleAccounting", "strike")
ZOOMS = (100, 200)
COLOR_FORMATS = (
    "General",
    "[Blue]General",
    "[Red]0.0",
    "[Blue]@",
    "[Blue]0;[Red]0",
    "[Blue]0;[Red]0;[Magenta]0",
    "[Blue]0;[Red]0;[Magenta]0;@",
    "[Blue]0;[Red]0;[Magenta]0;[Red]@",
    "[Red][<0]0;[Blue]0",
)


@dataclass(frozen=True, slots=True)
class Case:
    kind: str  # "rule" or "color"
    tag: str
    face: str = "Calibri"
    size: float = 11
    rule: str = "none"
    value: object = "HHHH"
    format: str = "General"


def cases() -> list[Case]:
    out = [
        Case("rule", f"{face} {s} {rule}", face, float(s), rule)
        for face in FACES
        for s in SIZES
        for rule in RULES
    ]
    out += [
        Case("color", f"{value!r} {fmt}", value=value, format=fmt)
        for fmt in COLOR_FORMATS
        for value in ("Total", 12)
    ]
    return out


def build() -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    for r, case in enumerate(cases(), 1):
        ws.cell(row=r, column=1).fill = markers[r % 2]
        cell = ws.cell(row=r, column=2, value=case.value)
        if isinstance(case.value, str):
            cell.data_type = "s"
        cell.number_format = case.format
        underline = None if case.rule in ("none", "strike") else case.rule
        cell.font = Font(
            name=case.face, sz=case.size, family=2, u=underline, strike=case.rule == "strike"
        )
        cell.alignment = Alignment(horizontal="left", vertical="bottom")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(PATH)
    return PATH


# --- predictions: the font's tables at the screen size, divided back ----------------------------


def predict_rule(face: str, size: float, rule: str, zoom: int) -> str:
    """Each rule's (top, thickness) in sheet points from the baseline, down positive."""
    t = tables(FILES[face])
    upem, p, k = t["upem"], size * zoom / 100, zoom / 100
    thick = max(1, round(p * t["underline_thickness"] / upem)) / k
    if rule == "none":
        return ""
    if rule == "strike":
        tops = [-round(p * t["strikeout_position"] / upem) / k]
    elif rule in ("single", "double"):
        tops = [math.floor(p * -t["underline_position"] / upem) / k]
    else:
        tops = [round(p * t["win_descent"] / upem) / k]
    if rule.startswith("double"):
        tops.append(tops[0] + 2 * thick)
    return " ".join(f"{top:g}/{thick:g}" for top in tops)


def predict() -> None:
    with PREDICTIONS.open("w", encoding="utf-8") as out:
        out.write(
            f"# predicted {dt.datetime.now().isoformat(timespec='seconds')}: per rule, "
            "top/thickness\n"
        )
        out.write(
            "# in sheet points below the baseline (strike: negative, above); color: sheetpeek's\n"
        )
        out.write("# finding that one-section formats color text, the first section's color\n")
        for zoom in ZOOMS:
            for c in cases():
                if c.kind == "rule":
                    out.write(f"{zoom}\t{c.tag}\t{predict_rule(c.face, c.size, c.rule, zoom)}\n")
                elif zoom == 100:
                    out.write(f"{zoom}\t{c.tag}\t{_predict_color(c)}\n")
    print(f"wrote {PREDICTIONS}")


def _predict_color(c: Case) -> str:
    """A number: the package's color today. Text: the hypothesis under test, the text section's
    color with four sections, else the first section's."""
    if not isinstance(c.value, str):
        from xlformat import format_cell

        color = format_cell(c.value, c.format).color
        return "black" if color is None else color.lower()
    sections = c.format.split(";")
    return _color_name(sections[3] if len(sections) == 4 else sections[0])


def _color_name(section: str) -> str:
    for name in ("Blue", "Red", "Magenta"):
        if section.startswith(f"[{name}]") or f"][{name}]" in section:
            return name.lower()
    return "black"


# --- capture and measurement ---------------------------------------------------------------------


def run() -> None:
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    if _request_count() != "0" or excel_windows():
        sys.exit("Excel must hold no workbook and show no window: nothing run")
    build()
    n = len(cases())
    _request("open", str(PATH), NAME, seconds=120.0)
    try:
        heights = _request("heights", NAME, str(n), seconds=120.0)
        window = _window_id(NAME.removesuffix(".xlsx"))
        for zoom in ZOOMS:
            folder = SHOTS / f"zoom{zoom}"
            folder.mkdir(parents=True, exist_ok=True)
            for old in folder.glob("*.png"):
                old.unlink()
            (folder / "heights.txt").write_text(heights + "\n")
            reported = _request("zoom", NAME, str(zoom))
            if float(reported) != zoom:
                sys.exit(f"zoom {zoom} not set: Excel reports {reported!r}")
            top = 1
            while top <= n:
                if _request("scroll", NAME, str(top)) != str(top):
                    sys.exit(f"row {top} not at the top")
                shot = folder / f"top_{top:03d}.png"
                _capture(window, shot)
                rows = rows_in_view(read_png(shot), top, [float(h) for h in heights.split(",")])
                if not rows or rows[0][0] != top:
                    sys.exit(f"{shot.name}: row {top} not found whole at the top")
                last = rows[-1][0]
                top = last if last > top and last < n else last + 1
            print(f"zoom {zoom}: {len(list(folder.glob('*.png')))} views")
    finally:
        _request("close", NAME)
    write()


def _rule_rows(
    image: Image, x0: int, x1: int, y0: int, y1: int
) -> tuple[int | None, list[tuple[int, int]]]:
    """The H stems' bottom and the rules' (first, last) pixel rows within a band."""
    longest, extent = {}, []
    for y in range(y0, y1 + 1):
        run = best = 0
        cols = []
        for x in range(x0, x1):
            if is_ink(image.rgb(x, y)):
                run += 1
                best = max(best, run)
                cols.append(x)
            else:
                run = 0
        longest[y] = best
        if cols:
            extent.append((min(cols), max(cols)))
    if not extent:
        return None, []
    # the text's extent: columns of rows whose ink is not one long run (a rule)
    widest = max(b - a + 1 for a, b in extent)
    text_rows = [y for y in longest if 0 < longest[y] < 0.5 * widest]
    if not text_rows:
        return None, []
    text_cols = [x for y in text_rows for x in range(x0, x1) if is_ink(image.rgb(x, y))]
    width = max(text_cols) - min(text_cols) + 1
    rule_ys = [y for y in longest if longest[y] >= 0.95 * width]
    runs: list[tuple[int, int]] = []
    for y in rule_ys:
        if runs and y == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], y)
        else:
            runs.append((y, y))
    stems = [y for y in text_rows if y not in rule_ys]
    return max(stems), runs


def _color(image: Image, x0: int, x1: int, y0: int, y1: int) -> str:
    colored, dark = [], 0
    for y in range(y0, y1 + 1):
        for x in range(x0, x1):
            r, g, b = image.rgb(x, y)
            if max(r, g, b) - min(r, g, b) > 80 and 0.299 * r + 0.587 * g + 0.114 * b < 200:
                colored.append((r, g, b))
            elif max(r, g, b) < 90:
                dark += 1
    if len(colored) > dark:
        n = len(colored)
        return "{} {} {}".format(*(round(sum(c[i] for c in colored) / n) for i in range(3)))
    return "black" if dark else "none"


def read_zoom(zoom: int) -> list[str]:
    folder = SHOTS / f"zoom{zoom}"
    heights = [float(h) for h in (folder / "heights.txt").read_text().strip().split(",")]
    all_cases = cases()
    found: dict[int, str] = {}
    at_edge: set[int] = set()
    for shot in sorted(folder.glob("top_*.png")):
        top = int(shot.stem.split("_")[1])
        image = read_png(shot)
        rows = rows_in_view(image, top, heights)
        if not rows:
            continue
        first_row, first_band = rows[0]
        ratios = sorted((b[1] - b[0] + 1) / heights[r - 1] for r, b in rows)
        scale = round(ratios[len(ratios) // 2])
        for k, (row, _) in enumerate(rows):
            edge = k == len(rows) - 1 and row < len(all_cases)
            if row in found and (row not in at_edge or edge):
                continue
            top_px = first_band[0] + round(scale * sum(heights[first_row - 1 : row - 1]))
            bottom_px = top_px + round(scale * heights[row - 1]) - 1
            left = _marker_right(image, (top_px + bottom_px) // 2, _marker_column(image)) + 2
            right = min(image.width, left + 900)
            case = all_cases[row - 1]
            if case.kind == "rule":
                stems, runs = _rule_rows(image, left, right, top_px, bottom_px)
                base = "" if stems is None else str(stems - top_px)
                text = " ".join(f"{a - top_px}-{b - top_px}" for a, b in runs)
                measured = f"{base}\t{text}"
            else:
                measured = f"\t{_color(image, left, right, top_px, bottom_px)}"
            found[row] = f"{zoom}\t{case.tag}\t{heights[row - 1]:g}\t{scale}\t{measured}"
            if edge:
                at_edge.add(row)
            else:
                at_edge.discard(row)
    missing = [r for r in range(1, len(all_cases) + 1) if r not in found]
    if missing:
        sys.exit(f"zoom {zoom}: rows never measured: {missing[:10]}")
    return [found[r] for r in sorted(found)]


def write() -> None:
    lines = [line for zoom in ZOOMS for line in read_zoom(zoom)]
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: rules (capture.screen_rules); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; window captures at zoom 100 and 200 on a "
            "display of scale 2; pixels from the row's top: the H stems' last ink row, each "
            "rule's first and last row; color: the mean of the text's colored ink, or black\n"
        )
        out.write("zoom\ttag\trow_height\tpixels_per_point\tbaseline_px\tmeasured\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {OUTPUT}: {len(lines)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        print(build())
    elif command == "predict":
        predict()
    elif command == "run":
        run()
    elif command == "read":
        write()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
