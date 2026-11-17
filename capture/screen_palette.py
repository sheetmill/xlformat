"""Which color Excel draws a section color in, under a workbook palette (a probe; screen).

    uv run python -m capture.screen_palette run     # both workbooks; then read
    uv run python -m capture.screen_palette read    # measure the saved captures again

A workbook's ``<indexedColors>`` replaces Excel's 64-entry palette; ``[ColorN]`` is entry N + 7.
Excel's scripting reports the palette's color for the eight names (``[Red]`` ...) whatever the
screen shows (CAPTURING.md section 2 item 21), so the screen decides. Two workbooks: one with a
custom palette of 64 colors far from the default ones, one with the default palette as the
control. Each has a row per name and per ``[Color1]`` to ``[Color56]``: 88888888 in bold 20-point
text under ``[<name>]0`` on a brown fill (so white text shows), with column A's alternating marker
fills marking the rows as in ``capture/screen_vertical.py``. The captures are converted to sRGB
(``sips --matchTo``); a cell's text color is its most frequent pixel away from the fill, matched
to the nearest of the workbook's palette entry, the default palette's entry and the name's
standard color. Writes ``tests/captures/palette.layout.tsv``; needs the display on, Excel in
front and nobody at the keyboard; no clipboard.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from collections import Counter
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, excel_windows
from capture.screen_vertical import (
    Image,
    _capture,
    _marker_column,
    _marker_right,
    _request,
    _window_id,
    read_png,
    rows_in_view,
)
from xlformat import DEFAULT_PALETTE

SHOTS = BUILD_DIR / "screen_palette"
OUTPUT = BUILD_DIR.parent.parent / "tests" / "captures" / "palette.layout.tsv"
NAMES = ("Black", "White", "Red", "Green", "Blue", "Yellow", "Magenta", "Cyan")
TAGS = (*NAMES, *(f"Color{i}" for i in range(1, 57)))
FILL = "5A3C28"
ROW_HEIGHT = 32.0
SRGB = "/System/Library/ColorSync/Profiles/sRGB Profile.icc"
STANDARD = {name: DEFAULT_PALETTE[name] for name in NAMES}


def _rgb(hex6: str) -> tuple[int, int, int]:
    return int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return sum((x - y) ** 2 for x, y in zip(a, b, strict=True)) ** 0.5


def custom_palette() -> tuple[str, ...]:
    """64 colors from a grid of the RGB cube, each at least 40 from every default palette color
    and the fill and 30 from each other, so a measured color names its source."""
    avoid = [_rgb(c) for c in (*DEFAULT_PALETTE.values(), FILL)]
    out: list[tuple[int, int, int]] = []
    steps = (20, 70, 120, 170, 220)
    for r in steps:
        for g in steps:
            for b in steps:
                c = (r, g, b)
                if all(_distance(c, a) >= 40 for a in avoid) and all(
                    _distance(c, o) >= 30 for o in out
                ):
                    out.append(c)
    if len(out) < 64:
        sys.exit(f"only {len(out)} palette colors found")
    return tuple(_hex(c) for c in out[:64])


def _hex(c: tuple[int, int, int]) -> str:
    return "".join(f"{v:02X}" for v in c)


def palette_rgb(tag: str, palette: tuple[str, ...]) -> str:
    """The palette entry a tag names: the eight names are entries 0 to 7, ColorN is N + 7."""
    index = NAMES.index(tag) if tag in NAMES else int(tag[5:]) + 7
    return palette[index]


def _paths(kind: str) -> tuple[str, Path]:
    name = f"screen_palette_{kind}.xlsx"
    return name, (WORKBOOK_DIR / name).resolve()


def build(kind: str) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    if kind == "custom":
        wb._colors = tuple("FF" + c for c in custom_palette())
    ws = wb.active
    ws.title = "palette"
    markers = (PatternFill("solid", fgColor="FF00FF"), PatternFill("solid", fgColor="00FFFF"))
    fill = PatternFill("solid", fgColor=FILL)
    for r, tag in enumerate(TAGS, 1):
        ws.cell(row=r, column=1).fill = markers[r % 2]
        cell = ws.cell(row=r, column=2, value=88888888)
        cell.number_format = f"[{tag}]0"
        cell.font = Font(name="Calibri", sz=20, b=True)
        cell.fill = fill
        ws.row_dimensions[r].height = ROW_HEIGHT
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    path = _paths(kind)[1]
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def run_kind(kind: str) -> None:
    name, path = _paths(kind)
    build(kind)
    folder = SHOTS / kind
    folder.mkdir(parents=True, exist_ok=True)
    for old in folder.glob("*.png"):
        old.unlink()
    _request("open", str(path), name, seconds=120.0)
    try:
        heights = [ROW_HEIGHT] * len(TAGS)
        window = _window_id(name.removesuffix(".xlsx"))
        top = 1
        while top <= len(TAGS):
            if _request("scroll", name, str(top)) != str(top):
                sys.exit(f"{kind}: Excel did not put row {top} at the top")
            shot = folder / f"top_{top:03d}.png"
            _capture(window, shot)
            rows = rows_in_view(read_png(shot), top, heights)
            if not rows or rows[0][0] != top:
                sys.exit(f"{shot.name}: row {top} not found whole at the top")
            top = rows[-1][0] + 1
    finally:
        _request("close", name)
    print(f"{kind}: captured {len(list(folder.glob('top_*.png')))} views")


def _text_color(
    image: Image, srgb: Image, band: tuple[int, int]
) -> tuple[tuple[int, int, int], int]:
    """The most frequent pixel color in the cell away from the fill, and its count."""
    y0, y1 = band
    x0 = _marker_right(image, (y0 + y1) // 2, _marker_column(image)) + 4
    fill = _rgb(FILL)
    # column B ends where its fill does: the last fill pixel on the band's top rows
    near = [
        x
        for y in (y0 + 1, y0 + 2)
        for x in range(x0, min(image.width, x0 + 900))
        if _distance(srgb.rgb(x, y), fill) <= 40
    ]
    if not near:
        return (-1, -1, -1), 0
    x1 = max(near)
    counts: Counter[tuple[int, int, int]] = Counter()
    for y in range(y0 + 2, y1 - 1):
        for x in range(x0, x1 - 2):
            c = srgb.rgb(x, y)
            if _distance(c, fill) > 40:
                counts[c] += 1
    if not counts:
        return (-1, -1, -1), 0
    return counts.most_common(1)[0]


def write() -> None:
    lines = []
    for kind in ("custom", "default"):
        palette = (
            custom_palette()
            if kind == "custom"
            else tuple(DEFAULT_PALETTE[t] for t in TAGS[:8])
            + tuple(DEFAULT_PALETTE[f"Color{i}"] for i in range(1, 57))
        )
        found: dict[int, tuple[tuple[int, int, int], int]] = {}
        shots = (s for s in (SHOTS / kind).glob("top_*.png") if ".srgb" not in s.name)
        for shot in sorted(shots):
            converted = shot.with_name(shot.stem + ".srgb.png")
            if not converted.exists():
                subprocess.run(
                    ["sips", "--matchTo", SRGB, str(shot), "--out", str(converted)],
                    capture_output=True, check=True,
                )  # fmt: skip
            image, srgb = read_png(shot), read_png(converted)
            top = int(shot.stem.split("_")[1])
            for row, band in rows_in_view(image, top, [ROW_HEIGHT] * len(TAGS)):
                found.setdefault(row, _text_color(image, srgb, band))
        for r, tag in enumerate(TAGS, 1):
            color, count = found.get(r, ((-1, -1, -1), 0))
            candidates = {
                "workbook": palette_rgb(tag, palette),
                "default": DEFAULT_PALETTE[tag],
                **({"standard": STANDARD[tag]} if tag in NAMES else {}),
            }
            dist = {k: round(_distance(color, _rgb(v))) for k, v in candidates.items()}
            nearest = min(dist, key=lambda k: dist[k])
            measured = _hex(color) if count else ""
            lines.append(
                f"{kind}\t{tag}\t{measured}\t{count}\t{nearest}\t"
                + " ".join(f"{k}={v}:{dist[k]}" for k, v in candidates.items())
            )
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as f:
        f.write("# suite: palette (capture.screen_palette); Excel for Mac, en_US\n")
        f.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        f.write(f"# captured: {dt.date.today().isoformat()}; window captures in sRGB\n")
        f.write("palette\ttag\tmeasured\tpixels\tnearest\tcandidates (rgb:distance)\n")
        f.write("\n".join(lines) + "\n")
    print(f"wrote {OUTPUT}: {len(lines)} rows")


def main(argv: list[str]) -> None:
    command = argv[0] if argv else ""
    if command == "run":
        if not answers():
            sys.exit("Excel does not answer: nothing run")
        if excel_windows():
            sys.exit("a window of Excel's is on screen: nothing run")
        for kind in ("custom", "default"):
            run_kind(kind)
        write()
    elif command == "read":
        write()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
