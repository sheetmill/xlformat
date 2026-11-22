"""Fit glyph advances from the autofit measurements and write the glyph modules.

Usage: ``uv run python -m capture.glyph_fit [--windows] [--check]``. Reads every
``tests/captures/measure*.measure.tsv`` (with ``--windows``, ``tests/captures/windows/``, Excel for
Windows at 96 dpi, fitted in pixels: points times 96 / 72; CAPTURING.md section 2 item 65)
(CAPTURING.md section 2 items 24, 26 and 30): every printable ASCII character in runs of 1, 2, 5
and 10, digit pairs, month and day names, AM, PM, E+07, ####, mixed strings, autofitted one per
column, with the column width read back in points. The header names the workbook's Normal font
(family and size); a ``cell_font`` column names the styled face the cells carried, else the
cells are the Normal font.

Model: the autofit width in points is ``floor(sum of glyph advances) + PAD``; PAD comes out of
the digit runs (``2 * points(5 digits) - points(10 digits)``): the column's padding plus 2 for an
upright face (7 for a digit width up to 8, 9 for 9 to 12, 11 for 13;
``xlformat._measure.column_padding``),
more for an italic one, whose slant reaches half the excess past its advances on each side (the
overhang, which a cell's budget loses). The advance of a glyph is
``(points of its ten-run - PAD) / 10`` (digits are whole points, most glyphs are, a few wide ones
carry a 0.1 fraction). Autofit cannot resolve the space, so its advance is
chosen by rendering the matching ``space_fill`` suite through ``xlformat`` in Excel mode with each
candidate and keeping the one that matches every row (CAPTURING.md section 2 item 25). Writes
``src/xlformat/_glyphs_<family>.py``, one table per style and size.
"""

from __future__ import annotations

import ast
import json
import math
import re
from pathlib import Path

from capture.schema import read_rows
from xlformat import Font, format_value
from xlformat import _measure as measure_module
from xlformat import _render as render_module
from xlformat._measure import column_padding
from xlformat._platform import Platform

ROOT = Path(__file__).resolve().parents[1]
CAPTURE_DIR = ROOT / "tests" / "captures"
SPACE_CANDIDATES = [tenths / 10 for tenths in range(10, 161)]
_LABEL = re.compile(r"(.+?) (\d+(?:\.\d+)?)((?: bold)?)((?: italic)?)$")
MODULES = {"Calibri": "CALIBRI", "Aptos Narrow": "APTOS_NARROW", "Arial": "ARIAL"}
# the platform being fitted: the Mac (points) or Windows at 96 dpi (pixels); set by main()
PLATFORM = Platform.MAC


def _font(label: str) -> Font:
    match = _LABEL.match(label)
    if not match:
        msg = f"not a font label: {label!r}"
        raise ValueError(msg)
    return Font(match[1], float(match[2]), bool(match[3]), bool(match[4]))


def load(path: Path) -> tuple[Font, Font, dict[str, float]]:
    """The cell face, the Normal font (from the header) and text to autofit points."""
    lines = path.read_text().splitlines()
    match = re.search(r"font: ([^;]+)", lines[0])
    if not match:
        msg = f"{path.name}: header does not name the font"
        raise ValueError(msg)
    normal = _font(match[1].strip())
    faces: set[str] = set()
    out: dict[str, float] = {}
    for line in lines[1:]:
        if line.startswith(("#", "value_type")):
            continue
        fields = line.split("\t")
        device = float(fields[4]) * PLATFORM.scale
        if device != round(device):
            msg = f"{path.name}: {fields[4]} points is not a whole device unit"
            raise ValueError(msg)
        out[ast.literal_eval(fields[1])] = device
        if len(fields) > 5 and fields[5]:
            faces.add(fields[5])
    if len(faces) > 1:
        msg = f"{path.name}: more than one cell font {sorted(faces)}"
        raise ValueError(msg)
    face = _font(faces.pop()) if faces else normal
    return face, normal, out


def padding(measured: dict[str, float]) -> float:
    """Autofit padding in points, from the five- and ten-runs of every digit (must agree)."""
    pads = {2 * measured[d * 5] - measured[d * 10] for d in "0123456789"}
    if len(pads) != 1:
        msg = f"digit runs disagree on the padding: {sorted(pads)}"
        raise ValueError(msg)
    return pads.pop()


def fit(measured: dict[str, float]) -> tuple[dict[str, float], float, int, int, list[str]]:
    """Advances per glyph (space unset), padding, exact matches, off-by-one, worse strings."""
    pad = padding(measured)
    adv = {chr(c): (measured[chr(c) * 10] - pad) / 10 for c in range(33, 127)}
    exact = off = 0
    worse: list[str] = []
    for text, pts in measured.items():
        if " " in text:
            continue  # scored once the space is known
        diff = _predict(text, adv, pad) - pts
        if diff == 0:
            exact += 1
        elif abs(diff) == 1:
            off += 1
        else:
            worse.append(text)
    return adv, pad, exact, off, worse


def _predict(text: str, adv: dict[str, float], pad: float) -> float:
    return math.floor(sum(adv[c] for c in text) + 1e-9) + pad


def _install(face: Font, table: dict[str, float], overhang: float = 0.0) -> None:
    """Put ``table`` and ``overhang`` in the package's font registry for ``face`` (fitting)."""
    key = face.name.lower()
    fonts, overhangs = measure_module._REGISTRY[PLATFORM.name]
    if key not in fonts:
        fonts[key] = (face.name, {})
        overhangs[key] = {}
    fonts[key][1].setdefault(face.style, {})[int(face.size)] = table
    if overhang:
        overhangs[key].setdefault(face.style, {})[int(face.size)] = overhang
    measure_module._advances.cache_clear()
    measure_module._whole_advances.cache_clear()
    measure_module._overhang.cache_clear()
    render_module._excel_width.cache_clear()  # holds a measure built from the old table


def fit_space(
    path: Path, face: Font, normal: Font, adv: dict[str, float], overhang: float = 0.0
) -> tuple[float, int, int, list[float]]:
    """The space advance that renders the most rows of the face's ``space_fill`` suite exactly.

    Returns the advance, the rows matched, the rows tried and every candidate that ties.
    """
    fill = path.with_name(path.name.replace("measure", "space_fill", 1).replace(".measure", ""))
    if not fill.exists():
        msg = f"{fill.name} missing: capture it, autofit cannot resolve the space"
        raise FileNotFoundError(msg)
    _, rows = read_rows(fill)
    width_rows = [r for r in rows if r.width_value is not None]
    scores: dict[float, int] = {}
    for space in SPACE_CANDIDATES:
        _install(face, {**adv, " ": space}, overhang)
        scores[space] = sum(
            format_value(
                r.value,
                r.format_string,
                column_width=r.width_value,
                font=face,
                normal_font=normal,
                platform=PLATFORM,
            )
            == r.expected_text
            for r in width_rows
        )
    best = max(scores.values())
    ties = [space for space, score in scores.items() if score == best]
    return ties[0], best, len(width_rows), ties


def main(argv: list[str] | None = None) -> None:
    """Fit every measured face, report, write one module per family.

    ``--check`` fits the faces whose ``space_fill`` suite exists and writes nothing.
    """
    global PLATFORM
    check = "--check" in (argv or [])
    windows = "--windows" in (argv or [])
    PLATFORM = Platform.WINDOWS if windows else Platform.MAC
    source = CAPTURE_DIR / "windows" if windows else CAPTURE_DIR
    loaded = [(path, *load(path)) for path in source.glob("measure*.measure.tsv")]
    # regular faces first: a styled face's fill rows need its regular Normal font's table
    loaded.sort(key=lambda item: (item[1].style != "regular", item[1].name, item[1].size))
    tables: dict[str, dict[str, dict[int, dict[str, float]]]] = {}
    overhangs: dict[str, dict[str, dict[int, float]]] = {}
    for path, face, normal, measured in loaded:
        fill = path.with_name(path.name.replace("measure", "space_fill", 1).replace(".measure", ""))
        if check and not fill.exists():
            continue
        adv, pad, exact, off, worse = fit(measured)
        # autofit pads a text by the column's padding plus 2; an italic face's overhang beyond
        over = (pad - column_padding(adv["0"]) - 2) / 2
        space, matched, tried, ties = fit_space(path, face, normal, adv, over)
        adv[" "] = space
        _install(face, adv, over)
        # every measured size, zeros included: a size missing from a style's map is scaled from
        # the nearest one listed, so a measured face without overhang must still be listed
        overhangs.setdefault(face.name, {}).setdefault(face.style, {})[int(face.size)] = over
        spaced = [t for t in measured if " " in t and set(t) != {" "}]
        space_exact = sum(_predict(t, adv, pad) == measured[t] for t in spaced)
        tables.setdefault(face.name, {}).setdefault(face.style, {})[int(face.size)] = adv
        print(
            f"{face.name} {face.style} {face.size:g}: padding {pad:g} (overhang {over:g}), "
            f"digit {adv['0']:g}; "
            f"{exact} texts exact, {off} off by one, {len(worse)} worse {worse[:5]}; space "
            f"{space:g} ({matched} of {tried} fill rows, ties {ties}; {space_exact} of "
            f"{len(spaced)} spaced autofit texts exact)"
        )
    if not check:
        for family, styles in tables.items():
            _write(family, styles, overhangs.get(family, {}), windows)


def _q(text: str) -> str:
    """A string literal as ruff formats it: double quotes unless the text holds one."""
    return repr(text) if '"' in text else json.dumps(text)


def _write(
    family: str,
    styles: dict[str, dict[int, dict[str, float]]],
    overhangs: dict[str, dict[int, float]],
    windows: bool = False,
) -> None:
    name = ("WINDOWS_" if windows else "") + MODULES[family]
    slug = ("windows_" if windows else "") + family.lower().replace(" ", "_")
    where = (
        "pixels at 96 dpi, measured in Excel for Windows"
        if windows
        else ("points, measured in Excel for Mac")
    )
    source = "tests/captures/windows/" if windows else "tests/captures/"
    order = [s for s in ("regular", "bold", "italic", "bold italic") if s in styles]
    sizes = sorted({size for style in order for size in styles[style]})
    lines = [
        f'"""{family} glyph advances in {where} (generated).',
        "",
        f"Written by ``capture/glyph_fit.py`` from ``{source}measure*.measure.tsv``.",
        f"Styles: {', '.join(order)}. Sizes: {', '.join(str(n) for n in sizes)}.",
        "Each advance is the autofit width of the glyph's run, less the autofit padding (the",
        *(
            [
                "column's padding plus 2: 7 pixels for a digit width up to 8, 9 for 9 to 12, 11",
                "for 13 to 16, 13 for 17 to 20; more for italic faces); the space comes from the",
                "matching ``space_fill`` suites.",
            ]
            if windows
            else [
                "column's padding plus 2: 7 points for a digit width up to 8, 9 for 9 to 12, 11 for"
                " 13;",
                "more for italic faces); the space comes from the matching ``space_fill`` suites.",
            ]
        ),
        "Advances are hinted per size, so the tables differ beyond scale. The overhang table",
        f"(``{name}_OVERHANG``) holds how far an italic face reaches past its advances on each",
        "side, which a cell's budget loses. Do not edit by hand.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        f"{name}: dict[str, dict[int, dict[str, float]]] = {{",
    ]
    for style in order:
        lines.append(f"    {_q(style)}: {{")
        for size in sorted(styles[style]):
            lines.append(f"        {size}: {{")
            table = styles[style][size]
            lines.extend(f"            {_q(ch)}: {table[ch]:g}," for ch in sorted(table))
            lines.append("        },")
        lines.append("    },")
    lines.append("}")
    lines.append("")
    lines.append(f"{name}_OVERHANG: dict[str, dict[int, float]] = {{")
    for style in [s for s in order if s in overhangs]:
        sizes = ", ".join(
            f"{size}: {overhangs[style][size]:g}" for size in sorted(overhangs[style])
        )
        lines.append(f"    {_q(style)}: {{{sizes}}},")
    lines.append("}")
    lines.append("")
    out = ROOT / "src" / "xlformat" / f"_glyphs_{slug}.py"
    out.write_text("\n".join(lines))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    import sys

    main(sys.argv[1:])
