"""Fit Calibri glyph advances from the autofit measurements and write the glyph module.

Usage: ``uv run python -m capture.glyph_fit``. Reads every ``tests/captures/measure*.measure.tsv``
(CAPTURING.md section 2 items 24 and 26), taking the point size from the file's header: every
printable ASCII character in runs of 1, 2, 5 and 10, digit pairs, month and day names, AM, PM,
E+07, ####, mixed strings, autofitted one per column, with the column width read back in points.

Model: the autofit width in points is ``floor(sum of glyph advances) + PAD``; PAD comes out of
the digit runs (``2 * points(5 digits) - points(10 digits)``, 7 at every size measured) and the
advance of a glyph is ``(points of its ten-run - PAD) / 10`` (digits are whole points, most glyphs
are, a few wide ones carry a 0.1 fraction). Autofit cannot resolve the space, so its advance is
chosen by rendering the ``space_fill<size>`` suite through ``xlformat`` in Excel mode with each
candidate and keeping the one that matches every row (CAPTURING.md section 2 item 25). Writes
``src/xlformat/_glyphs_calibri.py``.
"""

from __future__ import annotations

import ast
import math
import re
from pathlib import Path

from capture.schema import read_rows
from xlformat import _glyphs_calibri, format_value
from xlformat._measure import _advances

ROOT = Path(__file__).resolve().parents[1]
CAPTURE_DIR = ROOT / "tests" / "captures"
SPACE_CANDIDATES = [tenths / 10 for tenths in range(10, 81)]


def load(path: Path) -> tuple[int, dict[str, float]]:
    """Point size (from the header) and text to autofit points."""
    lines = path.read_text().splitlines()
    match = re.search(r"font: Calibri (\d+)", lines[0])
    if not match:
        msg = f"{path.name}: header does not name the font size"
        raise ValueError(msg)
    out: dict[str, float] = {}
    for line in lines[1:]:
        if line.startswith(("#", "value_type")):
            continue
        _, value_repr, _, _, points = line.split("\t")
        out[ast.literal_eval(value_repr)] = float(points)
    return int(match[1]), out


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


def fit_space(size: int, adv: dict[str, float]) -> tuple[float, int, int, list[float]]:
    """The space advance that renders the most ``space_fill<size>`` width rows exactly.

    Returns the advance, the rows matched, the rows tried and every candidate that ties.
    """
    path = CAPTURE_DIR / f"space_fill{size}.tsv"
    if not path.exists():
        msg = f"{path.name} missing: capture it, autofit cannot resolve the space"
        raise FileNotFoundError(msg)
    _, rows = read_rows(path)
    width_rows = [r for r in rows if r.width_value is not None]
    scores: dict[float, int] = {}
    for space in SPACE_CANDIDATES:
        _glyphs_calibri.CALIBRI[size] = {**adv, " ": space}
        _advances.cache_clear()
        scores[space] = sum(
            format_value(
                r.value, r.format_string, column_width=r.width_value, font=("Calibri", size)
            )
            == r.expected_text
            for r in width_rows
        )
    best = max(scores.values())
    ties = [space for space, score in scores.items() if score == best]
    return ties[0], best, len(width_rows), ties


def main() -> None:
    """Fit every measured size, report, write the module."""
    tables: dict[int, dict[str, float]] = {}
    for path in sorted(CAPTURE_DIR.glob("measure*.measure.tsv")):
        size, measured = load(path)
        adv, pad, exact, off, worse = fit(measured)
        space, matched, tried, ties = fit_space(size, adv)
        adv[" "] = space
        spaced = [t for t in measured if " " in t and set(t) != {" "}]
        space_exact = sum(_predict(t, adv, pad) == measured[t] for t in spaced)
        tables[size] = adv
        print(
            f"Calibri {size}: padding {pad:g}, digit {adv['0']:g}; {exact} texts exact, "
            f"{off} off by one, {len(worse)} worse {worse[:5]}; space {space:g} "
            f"({matched} of {tried} fill rows, ties {ties}; {space_exact} of {len(spaced)} "
            f"spaced autofit texts exact)"
        )
    sizes = ", ".join(str(s) for s in sorted(tables))
    lines = [
        '"""Calibri glyph advances in points, measured in Excel for Mac 16.114 (generated).',
        "",
        "Written by ``capture/glyph_fit.py`` from ``tests/captures/measure<size>.measure.tsv``",
        f"(sizes {sizes}): the autofit width of every printable ASCII glyph in runs, less the",
        "autofit padding of 7 points; the space from the ``space_fill<size>`` suites. Advances",
        "are hinted per size, so the tables differ beyond scale (``-`` is 3 at 11 points and 4",
        "at 12). Do not edit by hand.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "CALIBRI: dict[int, dict[str, float]] = {",
    ]
    for size in sorted(tables):
        lines.append(f"    {size}: {{")
        lines.extend(f"        {ch!r}: {tables[size][ch]:g}," for ch in sorted(tables[size]))
        lines.append("    },")
    lines.append("}")
    lines.append("")
    (ROOT / "src" / "xlformat" / "_glyphs_calibri.py").write_text("\n".join(lines))
    print("wrote src/xlformat/_glyphs_calibri.py")


if __name__ == "__main__":
    main()
