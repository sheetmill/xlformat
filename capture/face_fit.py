"""Recover each face's glyph advances from the faces captures; write ``src/xlformat/_faces_mac.py``.

Usage: ``uv run python -m capture.face_fit [--check]`` (``--check`` re-derives the module and
compares it with the committed one). Reads ``tests/captures/faces/`` (``capture/faces.py``;
CAPTURING.md section 2 item 82).

Model (Excel for Mac): a glyph's advance at size S is ``u * S / E`` points, u an integer and E
the face's units per em; an autofit width is the text's width rounded half up plus the padding P
(the same for every text of one face and size). A text holding a non-ASCII glyph sums its
glyphs unrounded, so a recovery text ``T + g * N + T`` (``T`` :data:`capture.faces.TRIGGER`) at
width W gives ``S * (2 * t + N * g)`` in ``[W - P - 0.5, W - P + 0.5]``, and a run ``g * N`` of
a non-ASCII glyph ``S * N * g`` there. At 383, 397 and 409 pt each interval is a fraction of a
font unit wide.

Recovery per face (family and style), for each combination of the paddings its digit runs
allow at the three sizes (twice a five-run less the ten-run is P within one point) and each
usual E (1000, 1024, 2000, 2048, 4096, smallest first): the trigger's runs pin ``t`` to one
multiple of ``1/E``, then every other glyph's intervals, intersected across the sizes, must hold
exactly one multiple of ``1/E``. Up to four glyphs may hold none: glyphs the face lacks, which
Excel draws from another face; their advance is kept unrounded (the middle of the interval) and
flagged. The space of an italic monospace face also comes out half a unit past the grid in
these runs, while the check captures' spaced texts take the advance every other glyph of the
face shares, which it is given. A combination of paddings must also follow the padding rule
``2 * ceil(d / 4) + 3``, d the face's widest digit at the size (rounded half up), plus an even
overhang for an italic face; the solutions left must agree on every advance.

Segoe UI, which Excel for Mac lacks, is captured too but has no solution: its widths are not
proportional to the size (its digit runs give 0.5379 em at 409 pt and 0.5390 em at 397).

The italic overhang at sizes 6 to 192 comes from the ``faces_p*`` captures: half of the padding
(a five-run less five whole digits) in excess of the rule's. Each ten-run must then be ten
digits, the padding and the autofit bonus ``max(0, floor(W / D / 5) - 2)`` (W the run's width
plus one overhang, D the Normal font's digit, Calibri at the workbook's size), which a wide
digit earns against Calibri's.
"""

from __future__ import annotations

import itertools
import json
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from capture.faces import (
    ABSENT,
    CHECK_SIZES,
    FACES_DIR,
    FAMILIES,
    GLYPHS,
    RECOVER_SIZES,
    STYLES,
    TRIGGER,
    _key,
    measure_path,
    read_measure,
)

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "src" / "xlformat" / "_faces_mac.py"
E_CANDIDATES = (1000, 1024, 2000, 2048, 4096)
OFF_GRID_ALLOWED = 4
EPS = 1e-9
DIGITS = "0123456789"
TABLE_FAMILIES = {"Calibri", "Arial", "Aptos Narrow"}  # with tables measured at 9 to 24 pt
TABLE_SIZES = (9, 10, 11, 12, 14, 16, 18, 20, 24)
SUBSTITUTES = {"MS Sans Serif": "Times New Roman"}  # the check captures: equal at every text
_LABEL = re.compile(r"(.+?) (\d+(?:\.\d+)?)((?: bold)?)((?: italic)?)$")

Interval = tuple[float, float]


@dataclass(frozen=True)
class Face:
    """One recovered face: E, each glyph's u (an int, or a float for a glyph the face lacks)."""

    e: int
    advances: dict[str, float]
    fallback: frozenset[str]
    pads: tuple[int, ...]


def parse_label(label: str) -> tuple[str, float, str]:
    match = _LABEL.match(label)
    if not match:
        msg = f"not a font label: {label!r}"
        raise ValueError(msg)
    style = " ".join(s for s in (match[3].strip(), match[4].strip()) if s) or "regular"
    return match[1], float(match[2]), style


def by_style(suite: str) -> dict[str, dict[str, float]]:
    """A capture's widths: style, then text, then points."""
    _, widths = read_measure(measure_path(suite))
    out: dict[str, dict[str, float]] = {}
    for (label, text), points in widths.items():
        out.setdefault(parse_label(label)[2], {})[text] = points
    return out


def whole(u: float, size: float, e: int) -> int:
    return math.floor(u * size / e + 0.5)


def rule_padding(digit: float) -> int:
    """Autofit's padding for an upright face whose widest digit is ``digit`` points."""
    return 2 * math.ceil(digit / 4) + 3


def pad_candidates(widths: dict[str, float]) -> list[int]:
    found: set[int] | None = None
    for d in DIGITS:
        v = int(2 * widths[d * 5] - widths[d * 10])
        found = {v - 1, v, v + 1} if found is None else found & {v - 1, v, v + 1}
    return sorted(found or ())


def _classify(text: str) -> tuple[str, int, bool] | None:
    """(glyph, run length, between triggers) of a recovery text; None for a digit run."""
    if len(set(text)) == 1:
        g = text[0]
        return None if g in DIGITS else (g, len(text), False)
    if text[0] == text[-1] == TRIGGER and len(set(text[1:-1])) == 1:
        return text[1], len(text) - 2, True
    msg = f"not a recovery text: {text!r}"
    raise ValueError(msg)


def _intersect(a: Interval | None, b: Interval) -> Interval:
    return b if a is None else (max(a[0], b[0]), min(a[1], b[1]))


def _on_grid(iv: Interval, e: int) -> list[int]:
    return list(range(math.ceil(iv[0] * e - EPS), math.floor(iv[1] * e + EPS) + 1))


def solve(sizes: dict[int, dict[str, float]], pads: dict[int, int], e: int) -> Face | None:
    """Every glyph's u on E's grid under the given paddings, or None when they do not fit."""
    runs = {size: {text: _classify(text) for text in ws} for size, ws in sizes.items()}
    trigger: Interval | None = None
    for size, ws in sizes.items():
        n = len(next(t for t in ws if set(t) == {TRIGGER}))
        w = ws[TRIGGER * n] - pads[size]
        trigger = _intersect(trigger, ((w - 0.5) / (size * n), (w + 0.5) / (size * n)))
    assert trigger is not None
    t_units = _on_grid(trigger, e)
    if len(t_units) != 1:
        return None
    t = t_units[0] / e
    found: dict[str, Interval] = {}
    for size, ws in sizes.items():
        for text, kind in runs[size].items():
            if kind is None or (kind[0] == TRIGGER and not kind[2]):
                continue
            g, n, between = kind
            w = ws[text] - pads[size]
            lo, hi = (w - 0.5) / size, (w + 0.5) / size
            if between:
                lo, hi = lo - 2 * t, hi - 2 * t
            found[g] = _intersect(found.get(g), (lo / n, hi / n))
    advances: dict[str, float] = {TRIGGER: float(t_units[0])}
    fallback = set()
    for g, iv in found.items():
        if iv[0] > iv[1] + EPS:
            return None  # the sizes disagree: a wrong padding
        units = _on_grid(iv, e)
        if len(units) > 1:
            return None
        if units:
            advances[g] = float(units[0])
        else:
            fallback.add(g)
            advances[g] = round((iv[0] + iv[1]) / 2 * e, 2)
    if len(fallback) > OFF_GRID_ALLOWED:
        return None
    return Face(e, advances, frozenset(fallback), tuple(pads[s] for s in sorted(pads)))


def follows_rule(face: Face, pads: dict[int, int], italic: bool) -> bool:
    for size, pad in pads.items():
        excess = pad - rule_padding(max(whole(face.advances[d], size, face.e) for d in DIGITS))
        if excess != 0 and not (italic and excess > 0 and excess % 2 == 0):
            return False
    return True


def recover(family: str, style: str) -> Face:
    sizes = {s: by_style(f"faces_r{s}_{_key(family)}")[style] for s in RECOVER_SIZES}
    solutions: list[Face] = []
    for combo in itertools.product(*(pad_candidates(sizes[s]) for s in RECOVER_SIZES)):
        pads = dict(zip(RECOVER_SIZES, combo, strict=True))
        for e in E_CANDIDATES:
            face = solve(sizes, pads, e)
            if face is not None:
                if follows_rule(face, pads, "italic" in style):
                    solutions.append(face)
                break  # the smallest E on whose grid the glyphs fall
    ratios = {tuple(sorted((g, u / f.e) for g, u in f.advances.items())) for f in solutions}
    if len(ratios) != 1:
        msg = f"{family} {style}: {len(solutions)} solutions, {len(ratios)} distinct"
        raise ValueError(msg)
    face = min(solutions, key=lambda f: f.e)
    others = {face.advances[chr(c)] for c in range(0x21, 0x7F)}
    if " " in face.fallback and len(others) == 1:
        # a monospace face's space, off the grid only in the trigger runs of its italic: every
        # check row (faces_c10_consolas: its spaced texts) takes the advance the others share
        advances = {**face.advances, " ": others.pop()}
        return Face(face.e, advances, face.fallback - {" "}, face.pads)
    return face


def overhangs(faces: dict[tuple[str, str], Face]) -> dict[tuple[str, str], dict[float, int]]:
    """Each face's overhang per whole padding size; raises where a face breaks the rule.

    A half size autofits as the whole size it rounds to, half up (``faces_phalf``), so it must
    give that size's padding and is not listed. Sizes 1 to 5 (``faces_p1to5``) are not read:
    every width there is a multiple of 5 points, the steps of that workbook's columns.
    """
    out: dict[tuple[str, str], dict[float, int]] = {}
    halves: list[tuple[str, str, float, float]] = []
    problems = []
    calibri = faces[("Calibri", "regular")]
    for path in sorted(FACES_DIR.glob("faces_p*.measure.tsv*")):
        if path.name.startswith("faces_p1to5."):
            continue  # under a Calibri 6 Normal font autofit moved in steps of 5 points
        header, widths = read_measure(path)
        normal = parse_label(header)[1]  # the Normal font: Calibri at that size
        normal_digit = max(1, *(whole(calibri.advances[d], normal, calibri.e) for d in DIGITS))
        for (label, text), five in widths.items():
            if text != "0" * 5:
                continue
            family, size, style = parse_label(label)
            if (family, style) not in faces:
                continue  # Segoe UI: no advances to check its padding against
            face = faces[(family, style)]
            whole_size = math.floor(size + 0.5)
            digit = max(1, *(whole(face.advances[d], whole_size, face.e) for d in DIGITS))
            zero = whole(face.advances["0"], whole_size, face.e)
            # the five-run is the padding's evidence; the ten-run can earn the autofit bonus
            pad = five - 5 * zero
            over = (pad - rule_padding(digit)) / 2
            bonus = max(0, math.floor((10 * zero + over) / normal_digit / 5) - 2)
            if widths[(label, "0" * 10)] != 10 * zero + pad + bonus:
                problems.append(f"{label}: ten-run {widths[(label, '0' * 10)]:g}")
            if whole_size != size:
                halves.append((family, style, whole_size, pad))
                continue
            excess = pad - rule_padding(digit)
            if excess < 0 or excess % 2 or (excess and "italic" not in style):
                problems.append(f"{label}: padding {pad:g}, rule {rule_padding(digit)}")
                continue
            out.setdefault((family, style), {})[size] = int(excess // 2)
    for family, style, size, pad in halves:
        digit = max(
            1,
            *(
                whole(faces[(family, style)].advances[d], size, faces[(family, style)].e)
                for d in DIGITS
            ),
        )
        expected = rule_padding(digit) + 2 * out.get((family, style), {}).get(size, 0)
        if pad != expected:
            problems.append(f"{family} {style} half size to {size}: padding {pad:g}, {expected}")
    if problems:
        raise ValueError("\n".join(problems))
    return out


def _numbers(values: list[str], indent: str) -> str:
    lines, line = [], ""
    for v in values:
        if len(line) + len(v) + 1 > 84:
            lines.append(line)
            line = ""
        line += ("" if not line else " ") + v
    lines.append(line)
    last = len(lines) - 1
    return "\n".join(f'{indent}"{line}{" " if k < last else ""}"' for k, line in enumerate(lines))


def _glyph_lines() -> list[str]:
    """``GLYPHS`` as ASCII string literals of at most 70 characters (escapes for the rest)."""
    lines, line = [], ""
    for g in GLYPHS:
        text = json.dumps(g)[1:-1]
        if len(line) + len(text) > 70:
            lines.append(f'    "{line}"')
            line = ""
        line += text
    return [*lines, f'    "{line}"']


def render(
    faces: dict[tuple[str, str], Face], over: dict[tuple[str, str], dict[float, int]]
) -> str:
    sizes = sorted({s for table in over.values() for s in table})
    out = [
        '"""Glyph advances of the faces Excel for Mac measures, from autofit widths (generated).',
        "",
        "Written by ``capture/face_fit.py`` from ``tests/captures/faces/``; do not edit by hand.",
        "``FACES`` holds per family and style the units per em E and, in the order of ``GLYPHS``,",
        "each glyph's advance u in units of 1/E em, or the name of a style with the same advances.",
        "``FALLBACK`` lists the glyphs whose advance is off the face's grid, kept as recovered,",
        "unrounded: glyphs the face lacks, which Excel draws from another face. ``OVERHANG``",
        "holds each italic face's overhang in points at ``OVERHANG_SIZES``. ``VERIFIED_SIZES``",
        "are the sizes whose every glyph's width was captured; ``SUBSTITUTES`` the names Excel",
        "for Mac has no face for and measures as another.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "GLYPHS = (",
        *_glyph_lines(),
        ")",
        "",
        "FACES: dict[str, dict[str, tuple[int, str] | str]] = {",
    ]
    fallback: list[str] = []
    for family in FAMILIES:
        out.append(f"    {family!r}: {{")
        seen: dict[tuple[int, tuple[float, ...]], str] = {}
        for style in STYLES:
            face = faces[(family, style)]
            key = (face.e, tuple(face.advances[g] for g in GLYPHS))
            if key in seen:
                out.append(f"        {style!r}: {seen[key]!r},")
                continue
            seen[key] = style
            values = [f"{face.advances[g]:g}" for g in GLYPHS]
            out += [f"        {style!r}: (", f"            {face.e},"]
            out += [_numbers(values, " " * 12), "        ),"]
            fallback += [f"    ({family!r}, {style!r}, {g!r})," for g in sorted(face.fallback)]
        out.append("    },")
    out += ["}", "", "FALLBACK: frozenset[tuple[str, str, str]] = frozenset({", *fallback, "})", ""]
    out.append(f"OVERHANG_SIZES = ({', '.join(f'{s:g}' for s in sizes)},)")
    out.append("OVERHANG: dict[tuple[str, str], str] = {")
    for (family, style), table in sorted(over.items(), key=lambda kv: FAMILIES.index(kv[0][0])):
        if "italic" in style:
            out.append(f"    ({family!r}, {style!r}): (")
            out += [_numbers([str(table[s]) for s in sizes], " " * 8), "    ),"]
    out += ["}", "", "VERIFIED_SIZES: dict[str, tuple[int, ...]] = {"]
    for family in FAMILIES:
        tables = set(TABLE_SIZES) if family in TABLE_FAMILIES else set()
        verified = sorted(set(CHECK_SIZES) | tables)
        out.append(f"    {family!r}: ({', '.join(map(str, verified))},),")
    out += ["}", "", f"SUBSTITUTES = {SUBSTITUTES!r}", ""]
    assert set(SUBSTITUTES) <= set(ABSENT)
    return "\n".join(out)


def fit() -> str:
    faces = {(family, style): recover(family, style) for family in FAMILIES for style in STYLES}
    for (family, style), face in faces.items():
        flagged = "".join(sorted(face.fallback))
        print(f"{family:18} {style:12} E={face.e} pads={face.pads} fallback={flagged!r}")
    return render(faces, overhangs(faces))


def main(argv: list[str]) -> None:
    import subprocess

    text = fit()
    scratch = MODULE.with_suffix(".new.py")
    scratch.write_text(text)
    subprocess.run(["ruff", "format", "-q", str(scratch)], check=True)
    formatted = scratch.read_text()
    scratch.unlink()
    if "--check" in argv:
        if not MODULE.exists() or MODULE.read_text() != formatted:
            sys.exit(f"{MODULE.name} differs from what the captures give")
        print(f"{MODULE.name} matches the captures")
        return
    MODULE.write_text(formatted)
    print(f"wrote {MODULE}")


if __name__ == "__main__":
    main(sys.argv[1:])
