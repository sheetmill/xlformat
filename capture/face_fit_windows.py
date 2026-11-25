"""Read whole-pixel advances per ppem from the Windows faces captures; write ``_faces_windows.py``.

Usage: ``uv run python -m capture.face_fit_windows [--check]`` (``--check`` re-derives the module
and compares it with the committed one). Reads ``tests/captures/windows/faces/``
(``capture/faces_windows.py``).

Model (Excel for Windows at 96 dpi): a cell font of size S draws at ppem ``floor(S * 4 / 3 + 0.5)``
and every glyph advances a whole number of pixels. An autofit width is the text's advances W plus
the padding P of its cell font plus a bonus pixel when ``W + overhang`` reaches 12 widths D of the
Normal font's widest digit (every text here stays under 14). The padding of an upright face is
``2 * ceil(d / 4) + 3``, d its widest digit; an italic face's excess over that is twice its
overhang. P is twice the ``0`` five-run's width less the ten-run's, less nothing or the bonus,
whichever makes every run whole; a run of n copies of a glyph is then n advances and a blank
between two ``x`` is ``(W - P - 2 x) / 5``. Every run must give a whole advance, every cell font of
one ppem the same advances, and every mixed check text the sum of its advances plus P and the
bonus.
"""

from __future__ import annotations

import ast
import gzip
import json
import sys
from collections import defaultdict
from pathlib import Path

from capture import faces
from capture import faces_windows as fw

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "src" / "xlformat" / "_faces_windows.py"
GLYPHS = faces.GLYPHS
DIGITS = "0123456789"


def ppem(size: float) -> int:
    """The ppem of a size at 96 dpi, a half rounded up."""
    return int(size * 4 / 3 + 0.5)


def parse_label(label: str) -> tuple[str, float, str]:
    """``Times New Roman 8.5 bold italic`` -> (family, size, style)."""
    words = label.split(" ")
    k = max(i for i, w in enumerate(words) if w.replace(".", "", 1).isdigit())
    return " ".join(words[:k]), float(words[k]), " ".join(words[k + 1 :]) or "regular"


def read(path: Path) -> tuple[str, dict[str, dict[str, int]]]:
    """A measure file's Normal font and its autofit pixels per cell font label and text."""
    with gzip.open(path, "rt", encoding="utf-8") as f:
        lines = f.read().split("\n")
    normal = lines[0].split("font: ", 1)[1].split(";", 1)[0]
    out: dict[str, dict[str, int]] = defaultdict(dict)
    for line in lines[3:]:
        if not line:
            continue
        fields = line.split("\t")
        px = float(fields[4]) * 4 / 3
        if abs(px - round(px)) > 1e-6:
            msg = f"{path.name}: {fields[4]} points is not a whole pixel"
            raise ValueError(msg)
        text = ast.literal_eval(fields[1])
        assert isinstance(text, str)
        out[fields[5]][text] = round(px)
    return normal, out


def rule_padding(digit: int) -> int:
    return 2 * -(-digit // 4) + 3


def bonus(text_px: float, over: float, normal_digit: int) -> int:
    """The autofit bonus in pixels: one from 12 Normal digit widths (texts under 14 only)."""
    return 1 if text_px + over >= 12 * normal_digit else 0


def _text_px(width: int, pad: int, over: float, normal_digit: int) -> int:
    """A text's advances from its autofit width: the width less the padding and the bonus."""
    w = width - pad
    return w - 1 if w - 1 + over >= 12 * normal_digit else w


def advances(
    widths: dict[str, int], normal_digit: int, italic: bool
) -> tuple[int, dict[str, int], list[str]]:
    """The padding and each glyph's advance from one cell font's widths; the misfits."""
    best: tuple[int, int, dict[str, int], list[str]] | None = None
    base = 2 * widths["0" * 5] - widths["0" * 10]
    for pad in (base, base + 1):
        digit = (widths["0" * 5] - pad) // 5
        over = (pad - rule_padding(digit)) / 2 if italic else 0.0
        adv: dict[str, int] = {}
        misfits: list[str] = []
        for g in [*(g for g in GLYPHS if g not in fw.BLANK), *sorted(fw.BLANK)]:
            w = _text_px(widths[fw.run_text(g)], pad, over, normal_digit)
            if g in fw.BLANK:
                w -= 2 * adv[fw.BRACKET]
                n = fw.RUNS["blank"]
            else:
                n = fw.copies(g)
            if w % n:
                misfits.append(f"{g!r}: {w}/{n}")
            adv[g] = w // n if w >= 0 else 0
        ten = _text_px(widths["0" * 10], pad, over, normal_digit)
        if ten != 10 * adv["0"]:
            misfits.append(f"'0' * 10: {ten} px, not ten digits")
        if best is None or len(misfits) < len(best[3]):
            best = (len(misfits), pad, adv, misfits)
    assert best is not None
    return best[1], best[2], best[3]


Key = tuple[str, str, int]  # family, style, ppem


def fit() -> tuple[dict[Key, dict[str, int]], dict[Key, int], list[str]]:
    """Advances and overhangs per (family, style, ppem), and every disagreement found."""
    tables: dict[Key, dict[str, int]] = {}
    overhang: dict[Key, int] = {}
    problems: list[str] = []
    for name in fw.captured()[1]:
        normal, by_label = read(fw.FACES_DIR / f"{name}.measure.tsv.gz")
        normal_widths = by_label[normal]  # the Normal font is also a cell font: its digit
        # the Normal font's widest digit (its ten-runs stay under 12 of its own digit widths)
        zero = (normal_widths["0" * 10] - normal_widths["0" * 5]) // 5
        _, normal_adv, _ = advances(normal_widths, zero, False)
        normal_digit = max(normal_adv[d] for d in DIGITS)
        for label, widths in by_label.items():
            family, size, style = parse_label(label)
            key = (family, style, ppem(size))
            pad, adv, misfits = advances(widths, normal_digit, "italic" in style)
            problems += [f"{label}: {m}" for m in misfits]
            excess = pad - rule_padding(max(adv[d] for d in DIGITS))
            if excess % 2 or excess < 0 or ("italic" not in style and excess):
                problems.append(f"{label}: padding {pad}, the rule's plus {excess}")
            for text in fw.CHECKS:
                px = sum(adv[c] for c in text)
                want = px + pad + bonus(px, excess / 2, normal_digit)
                if widths[text] != want:
                    problems.append(f"{label}: {text!r} {widths[text]} px, advances give {want}")
            if key in tables:
                differ = [g for g in GLYPHS if tables[key][g] != adv[g]]
                if differ or overhang[key] != excess // 2:
                    problems.append(f"{label}: ppem {key[2]} differs from another size: {differ}")
                continue
            tables[key] = adv
            overhang[key] = excess // 2
    return tables, overhang, problems


def _numbers(values: list[str], indent: str) -> list[str]:
    lines, line = [], ""
    for v in values:
        if len(line) + len(v) + 1 > 84:
            lines.append(line)
            line = ""
        line += ("" if not line else " ") + v
    lines.append(line)
    last = len(lines) - 1
    return [f'{indent}"{x}{" " if k < last else ""}"' for k, x in enumerate(lines)]


def render(tables: dict[Key, dict[str, int]], overhang: dict[Key, int]) -> str:
    ppems = sorted({k[2] for k in tables})
    out = [
        '"""Whole-pixel advances of the faces Excel for Windows measures, per ppem (generated).',
        "",
        "Written by ``capture/face_fit_windows.py`` from ``tests/captures/windows/faces/``; do not",
        "edit by hand. At 96 dpi a size S draws at ppem ``floor(S * 4 / 3 + 0.5)``. ``FACES``",
        "holds per family and style, for each ppem of ``PPEMS``, each glyph's advance in pixels in",
        "the order of ``GLYPHS`` (those of ``_faces_mac``), or the name of a style with the same",
        "advances. ``OVERHANG`` holds each italic face's overhang in pixels at ``PPEMS``.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        f"PPEMS = ({', '.join(map(str, ppems))},)",
        "",
        "FACES: dict[str, dict[str, tuple[str, ...] | str]] = {",
    ]
    for family in fw.captured()[0]:
        out.append(f"    {family!r}: {{")
        seen: dict[tuple[int, ...], str] = {}
        for style in fw.STYLES:
            flat = tuple(tables[(family, style, p)][g] for p in ppems for g in GLYPHS)
            if flat in seen:
                out.append(f"        {style!r}: {seen[flat]!r},")
                continue
            seen[flat] = style
            out.append(f"        {style!r}: (")
            for p in ppems:
                values = [str(tables[(family, style, p)][g]) for g in GLYPHS]
                body = _numbers(values, " " * 12)
                body[-1] += ","
                out += [f"            # ppem {p}", *body]
            out.append("        ),")
        out.append("    },")
    out += ["}", "", "OVERHANG: dict[tuple[str, str], str] = {"]
    for family in fw.captured()[0]:
        for style in fw.STYLES:
            if "italic" in style:
                values = [str(overhang[(family, style, p)]) for p in ppems]
                out += [f"    ({family!r}, {style!r}): (", *_numbers(values, " " * 8), "    ),"]
    out += ["}", ""]
    return "\n".join(out)


def main(argv: list[str]) -> None:
    import subprocess

    tables, overhang, problems = fit()
    for line in problems:
        print(line)
    if problems and "--force" not in argv:
        sys.exit(f"{len(problems)} disagreements; nothing written")
    scratch = MODULE.with_suffix(".new.py")
    scratch.write_text(render(tables, overhang))
    subprocess.run(["ruff", "format", "-q", str(scratch)], check=True)
    formatted = scratch.read_text()
    scratch.unlink()
    if "--check" in argv:
        if not MODULE.exists() or MODULE.read_text() != formatted:
            sys.exit(f"{MODULE.name} differs from what the captures give")
        print(f"{MODULE.name} matches the captures")
        return
    MODULE.write_text(formatted)
    print(f"wrote {MODULE} ({json.dumps(len(formatted))} bytes)")


if __name__ == "__main__":
    main(sys.argv[1:])
