"""Where Excel draws underlines and strikethrough, and what color text takes, from the screen.

    uv run python -m capture.screen_rules build [suite]    # the workbook, no Excel
    uv run python -m capture.screen_rules predict [suite]  # write the predictions, before a run
    uv run python -m capture.screen_rules run [suite]      # open, capture at zoom 100 and 200
    uv run python -m capture.screen_rules read [suite]     # measure the saved captures again

Two suites, ``rules`` (the default) and ``rules_double``.

One workbook, Normal style Calibri 11, column A filled magenta and cyan in turn as in
``capture/screen_vertical.py`` (whose AppleScript, window capture and band matching it uses), a
case a row in column B, left-aligned and bottom-aligned, rows fitted by Excel:

- ``rule``: "HHHH" in each face (``FACES``) at 9, 11, 14, 20 and 36 points without a rule and with
  each underline (single, double, single and double accounting) or a strikethrough. Measured at
  zoom 100 and 200: the baseline (the bottom of the H stems' ink) and each rule's first and last
  pixel row, a rule being an ink row whose longest run spans the text.
- ``color``: "Total" and 12 in Calibri 11 under formats with one to four sections and colors.
  Measured: the mean color of the text's colored ink, or ``black``.

``rules_double`` (CAPTURING.md section 2 item 62) has only ``rule`` cases: five faces (Courier New
left out) at every size from 9 to 36 points and from 40 to 72 in steps of 4, each with every
underline and a strikethrough (no plain row: the strikethrough row gives the plain row's
baseline), then 20 small rows, so that the sheet does not end on rows taller than a view; column
B is 60 characters wide, so that the text never overflows the cell.

The run needs the display on, Excel in front and nobody at the keyboard; no clipboard. Writes
``tests/captures/<suite>.layout.tsv``; the images stay in ``capture/build/screen_<suite>/``.
"""

from __future__ import annotations

import datetime as dt
import math
import sys
from collections.abc import Iterable
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
    numbered_views,
    read_png,
    read_views,
    rows_in_view,
)

CAPTURES = Path(__file__).resolve().parents[1] / "tests" / "captures"
SUITES = ("rules", "rules_double", "rules_courier", "rules_zoom")
FACES = ("Calibri", "Arial", "Aptos Narrow", "Times New Roman", "Courier New", "Georgia")
SIZES = (9, 11, 14, 20, 36)
RULES = ("none", "single", "double", "singleAccounting", "doubleAccounting", "strike")
ZOOMS = (100, 200)
DOUBLE_FACES = ("Calibri", "Arial", "Aptos Narrow", "Times New Roman", "Georgia")
DOUBLE_SIZES = (*range(9, 37), *range(40, 73, 4))
DOUBLE_RULES = ("single", "double", "singleAccounting", "doubleAccounting", "strike")
HELD_OUT_SEED = 20261121  # rules_double: a third of the combinations, held out of any choice
# rules_courier: Courier New at every size from 6 to 72 and every half size from 36.5 to 71.5 (odd
# screen sizes 73 to 143 at zoom 200), Times New Roman at every half size from 36 to 41, and
# Calibri as a control
COURIER_SIZES = (*range(6, 73), *(s + 0.5 for s in range(36, 72)))
TIMES_SIZES = tuple(36 + k / 2 for k in range(11))
CONTROL_SIZES = (9, 14, 20, 36, 48)
# rules_zoom: screen sizes that are not whole numbers, at zoom 150, which keeps a whole number of
# device pixels to a sheet point (3) as the reader needs, with zoom 100 beside (zoom 50, 1 pixel
# a point, stopped the reader at its first view of large rows)
ZOOM_SIZES = (
    8, 8.5, 9.5, 10.5, 11.5, 12.5, 16.5, 17, 17.5, 20.5, 22.5, 23, 23.5, 34.5, 35.5, 36.5,
    51.5, 52.5, 69.5, 70.5,
)  # fmt: skip


def zooms(suite: str) -> tuple[int, ...]:
    return (100, 150) if suite == "rules_zoom" else ZOOMS


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


def workbook_name(suite: str) -> str:
    return "screen_rules.xlsx" if suite == "rules" else f"screen_{suite}.xlsx"


def workbook_path(suite: str) -> Path:
    return (WORKBOOK_DIR / workbook_name(suite)).resolve()


def shots(suite: str) -> Path:
    return BUILD_DIR / f"screen_{suite}"


def output(suite: str) -> Path:
    return CAPTURES / f"{suite}.layout.tsv"


def predictions(suite: str) -> Path:
    return BUILD_DIR / f"screen_{suite}.predictions.tsv"


def _rule_cases(faces_sizes: list[tuple[str, float]]) -> list[Case]:
    found = [
        Case("rule", f"{face} {s:g} {rule}", face, float(s), rule)
        for face, s in faces_sizes
        for rule in DOUBLE_RULES
    ]
    return found + [Case("rule", f"filler {k} none") for k in range(20)]


def cases(suite: str = "rules") -> list[Case]:
    if suite == "rules_courier":
        return _rule_cases(
            [("Courier New", s) for s in COURIER_SIZES]
            + [("Times New Roman", s) for s in TIMES_SIZES]
            + [("Calibri", s) for s in CONTROL_SIZES]
        )
    if suite == "rules_zoom":
        return _rule_cases([(face, s) for face in ("Courier New", "Calibri") for s in ZOOM_SIZES])
    if suite == "rules_double":
        found = [
            Case("rule", f"{face} {s} {rule}", face, float(s), rule)
            for face in DOUBLE_FACES
            for s in DOUBLE_SIZES
            for rule in DOUBLE_RULES
        ]
        return found + [Case("rule", f"filler {k} none") for k in range(20)]
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


def build(suite: str = "rules") -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    for r, case in enumerate(cases(suite), 1):
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
    # rules_double: wide enough for "HHHH" at 72 points, since an accounting underline spans the
    # cell and is read as a rule only where it is as long as the text
    ws.column_dimensions["B"].width = 30 if suite == "rules" else 60
    path = workbook_path(suite)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


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


def predict(suite: str = "rules") -> None:
    if suite == "rules_double":
        predict_double()
        return
    with predictions(suite).open("w", encoding="utf-8") as out:
        out.write(
            f"# predicted {dt.datetime.now().isoformat(timespec='seconds')}: per rule, "
            "top/thickness\n"
        )
        out.write(
            "# in sheet points below the baseline (strike: negative, above); color: text takes\n"
        )
        out.write("# the text section's color in four sections, else the first section's\n")
        for zoom in ZOOMS:
            for c in cases(suite):
                if c.kind == "rule":
                    out.write(f"{zoom}\t{c.tag}\t{predict_rule(c.face, c.size, c.rule, zoom)}\n")
                elif zoom == 100:
                    out.write(f"{zoom}\t{c.tag}\t{_predict_color(c)}\n")
    print(f"wrote {predictions(suite)}")


def held_out() -> set[tuple[int, str, int]]:
    """rules_double: the (zoom, face, size) combinations held out of any choice made on the rest."""
    import random

    keys = [(z, f, s) for z in ZOOMS for f in DOUBLE_FACES for s in DOUBLE_SIZES]
    return set(random.Random(HELD_OUT_SEED).sample(keys, len(keys) // 3))


def predict_double() -> None:
    """rules_double: each line's top/thickness in sheet points below the baseline (strike:
    negative, above), as ``rule_metrics`` gives it (``current``) and with a double underline's
    first line stepped up (``step``). Written before the run of item 62, when ``rule_metrics``
    put a double's first line at the single underline's top; run now, ``current`` is the
    package's placement, which already steps."""
    from xlformat import rule_metrics

    held = held_out()
    with predictions("rules_double").open("w", encoding="utf-8") as out:
        out.write(
            f"# predicted {dt.datetime.now().isoformat(timespec='seconds')}: per rule, "
            "top/thickness in sheet points below the baseline (strike: above, negative)\n"
            "# current: rule_metrics as it stands\n"
            "# step: a double's first line steps up from the single's top by its thickness until "
            "the double (3 thicknesses) fits above the accounting underline's top; from the "
            "font tables\n"
            "# not a column, scored after the run: the same steps from the captured single "
            "underline's top, double's thickness and accounting underline's top\n"
            "# set: test = held out (seed 20261121), sel = the rest\n"
            "zoom\ttag\tset\tcurrent\tstep\n"
        )
        for zoom in ZOOMS:
            for c in cases("rules_double"):
                if c.tag.startswith("filler"):
                    continue
                m = rule_metrics((c.face, c.size), zoom=zoom)
                th = m.underline_thickness
                if c.rule == "strike":
                    cur = [(-m.strikethrough_top, m.strikethrough_thickness)]
                elif c.rule in ("single", "double"):
                    cur = [(m.underline_top, th)]
                else:
                    cur = [(m.accounting_top, th)]
                if c.rule.startswith("double"):
                    cur.append((cur[0][0] + m.double_gap, th))
                step = cur
                if c.rule == "double":
                    first = m.underline_top
                    while first > 0 and first + 3 * th > m.accounting_top:
                        first -= th
                    step = [(first, th), (first + 2 * th, th)]
                kind = "test" if (zoom, c.face, int(c.size)) in held else "sel"
                cur_s, step_s = (" ".join(f"{a:g}/{b:g}" for a, b in v) for v in (cur, step))
                out.write(f"{zoom}\t{c.tag}\t{kind}\t{cur_s}\t{step_s}\n")
    print(f"wrote {predictions('rules_double')}")


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


def run(suite: str = "rules") -> None:
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    if _request_count() != "0" or excel_windows():
        sys.exit("Excel must hold no workbook and show no window: nothing run")
    path = build(suite)
    name = workbook_name(suite)
    n = len(cases(suite))
    _request("open", str(path), name, seconds=120.0)
    try:
        heights = _request("heights", name, str(n), seconds=120.0)
        window = _window_id(name.removesuffix(".xlsx"))
        for zoom in zooms(suite):
            folder = shots(suite) / f"zoom{zoom}"
            folder.mkdir(parents=True, exist_ok=True)
            for old in folder.glob("*.png"):
                old.unlink()
            (folder / "heights.txt").write_text(heights + "\n")
            reported = _request("zoom", name, str(zoom))
            if float(reported) != zoom:
                sys.exit(f"zoom {zoom} not set: Excel reports {reported!r}")
            top = 1
            while top <= n:
                if _request("scroll", name, str(top)) != str(top):
                    sys.exit(f"row {top} not at the top")
                shot = folder / f"top_{top:03d}.png"
                _capture(window, shot)
                rows = rows_in_view(
                    read_png(shot), top, [float(h) for h in heights.split(",")], _scale(zoom)
                )
                if not rows or rows[0][0] != top:
                    sys.exit(f"{shot.name}: row {top} not found whole at the top")
                last = rows[-1][0]
                top = last if last > top and last < n else last + 1
            print(f"zoom {zoom}: {len(list(folder.glob('*.png')))} views")
    finally:
        _request("close", name)
    write(suite)


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


def _scale(zoom: int) -> float | None:
    """Device pixels a sheet point where the view's median would mislead: at zoom 150 Excel
    rounds each row to whole pixels, small rows up, so the median drifts above 3 and a tall row
    no longer fits it. At zoom 100 and 200 the median serves, as it always has."""
    return None if zoom in (100, 200) else 2 * zoom / 100


def read_zoom(zoom: int, suite: str = "rules") -> list[str]:
    folder = shots(suite) / f"zoom{zoom}"
    heights = [float(h) for h in (folder / "heights.txt").read_text().strip().split(",")]
    return measure_rules(
        numbered_views(folder),
        heights,
        zoom,
        suite,
        match_scale=_scale(zoom),
        own_bands=_scale(zoom) is not None,
    )


def measure_rules(
    views: Iterable[tuple[int, Path]],
    heights: list[float],
    zoom: int,
    suite: str,
    *,
    match_scale: float | None = None,
    pixels_per_point: float | None = None,
    own_bands: bool = False,
) -> list[str]:
    """Every row of a rules suite measured from its views at one zoom (``read_views``): per row
    the zoom, tag, height, device pixels a point, and the baseline and rules in pixels from the
    row's top (or a color row's color)."""
    all_cases = cases(suite)

    def measure_row(
        image: Image, row: int, band: tuple[int, int], top_px: int, bottom_px: int, scale: float
    ) -> str:
        left = _marker_right(image, (top_px + bottom_px) // 2, _marker_column(image)) + 2
        right = min(image.width, left + (900 if suite == "rules" else 1600))
        case = all_cases[row - 1]
        if case.kind == "rule":
            stems, runs = _rule_rows(image, left, right, top_px, bottom_px)
            base = "" if stems is None else str(stems - top_px)
            text = " ".join(f"{a - top_px}-{b - top_px}" for a, b in runs)
            measured = f"{base}\t{text}"
        else:
            measured = f"\t{_color(image, left, right, top_px, bottom_px)}"
        return f"{zoom}\t{case.tag}\t{heights[row - 1]:g}\t{scale:g}\t{measured}"

    found = read_views(
        views,
        heights,
        measure_row,
        label=f"zoom {zoom}",
        match_scale=match_scale,
        pixels_per_point=pixels_per_point,
        own_bands=own_bands,
    )
    return [found[r] for r in sorted(found)]


def display_profile(suite: str) -> str:
    """The color profile of the display the views were captured on, read from the first view:
    colors on a wide-gamut display (Display P3) are converted and read a few units off."""
    import subprocess

    first = min((shots(suite) / f"zoom{zooms(suite)[0]}").glob("top_*.png"))
    out = subprocess.run(
        ["sips", "-g", "profile", str(first)], check=True, capture_output=True, text=True
    ).stdout
    return out.rsplit("profile:", 1)[1].strip()


def write(suite: str = "rules") -> None:
    lines = [line for zoom in zooms(suite) for line in read_zoom(zoom, suite)]
    if suite != "rules":
        lines = [line for line in lines if "\tfiller " not in line]
    command = "capture.screen_rules" + ("" if suite == "rules" else f" {suite}")
    color = "; color: the mean of the text's colored ink, or black" if suite == "rules" else ""
    target = output(suite)
    with target.open("w", encoding="utf-8") as out:
        out.write(f"# suite: {suite} ({command}); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; window captures at zoom "
            f"{' and '.join(map(str, zooms(suite)))} on a "
            "display of scale 2; pixels from the row's top: the H stems' last ink row, each "
            f"rule's first and last row{color}\n"
        )
        out.write(f"# display_profile: {display_profile(suite)}\n")
        out.write("zoom\ttag\trow_height\tpixels_per_point\tbaseline_px\tmeasured\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {target}: {len(lines)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    suite = argv[1] if len(argv) > 1 else "rules"
    if suite not in SUITES or len(argv) > 2:
        sys.exit(__doc__)
    if command == "build":
        print(build(suite))
    elif command == "predict":
        predict(suite)
    elif command == "run":
        run(suite)
    elif command == "read":
        write(suite)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
