"""Predictions for the ``rules_courier`` and ``rules_zoom`` captures, and their scorer.

    uv run python -m capture.courier_predict predict   # write capture/predictions/*.tsv
    uv run python -m capture.courier_predict score     # score a capture against them

Written before the runs (CAPTURING.md section 2 item 63). Lines are (top, thickness) in screen
points below the baseline at the screen size ``P = size * zoom / 100``; a strikethrough's top is
negative (above). The rule predicted for Courier New: every line ``max(1, floor(P * 84 / 2048))``
thick; a double underline's second line at ``u = floor(P * 477 / 2048)`` and its first a gap
``g(P)`` above it (one point lower at ``P`` = 12); a single underline at ``u`` while ``g(P)`` is 2,
on the double's first line above that; an accounting underline at ``A = round(P * 615 / 2048)``
and a double accounting one's second line ``g(P)`` below it. ``GAP`` holds the ``g(P)`` the
predictions assume; at odd ``P`` from 73 to 143 they assume the neighbors' gap where both agree.
For a screen size that is not whole the candidates are: ``round``, ``floor`` and ``ceil`` (every
line as at that whole screen size) and ``exact`` (the formulas at the fractional size, the gap at
the nearest whole one). Calibri follows ``rule_metrics``' rule at the screen size given; Times New
Roman's double is predicted three thicknesses apart where its accounting rows add 4 points (36.5
to 40 points) and two elsewhere.
"""

from __future__ import annotations

import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

from capture.font_tables import FILES, tables
from capture.screen_rules import cases, output, zooms

ROOT = Path(__file__).resolve().parents[1]
PREDICTIONS = ROOT / "capture" / "predictions"
SUITES = ("rules_courier", "rules_zoom")
GAP = {
    **dict.fromkeys(range(6, 18), 2), **dict.fromkeys(range(18, 26), 3),
    **dict.fromkeys(range(26, 35), 4), **dict.fromkeys(range(35, 43), 5),
    **dict.fromkeys(range(43, 49), 6), **dict.fromkeys(range(49, 56), 7),
    **dict.fromkeys(range(56, 65), 8), **dict.fromkeys(range(65, 73), 9),
    74: 10, 76: 10, 78: 11, 80: 11, 82: 11, 84: 11, 86: 12, 88: 12, 90: 12, 92: 12, 94: 12,
    96: 13, 98: 13, 100: 14, 102: 14, 104: 14, 106: 14, 108: 15, 110: 15, 112: 15, 114: 15,
    116: 16, 118: 16, 120: 16, 122: 17, 124: 17, 126: 17, 128: 17, 130: 18, 132: 18, 134: 18,
    136: 18, 138: 19, 140: 19, 142: 19, 144: 19,
}  # fmt: skip
CANDIDATES = ("round", "floor", "ceil", "exact")
Lines = list[tuple[float, float]]


def gap(p: int) -> int | None:
    """The assumed gap at a whole screen size; None where it is not assumed."""
    if p in GAP:
        return GAP[p]
    below, above = GAP.get(p - 1), GAP.get(p + 1)
    return below if below is not None and below == above else None


def courier(p: float, rule: str, g: int | None) -> Lines | None:
    t = tables(FILES["Courier New"])
    em = t["upem"]
    thick = max(1, math.floor(p * t["underline_thickness"] / em))
    if rule == "strike":
        top = math.floor(p * t["strikeout_position"] / em + 0.5)
        return [(-top, max(1, math.floor(p * t["strikeout_size"] / em + 0.5)))]
    acc = math.floor(p * t["win_descent"] / em + 0.5)
    if rule == "singleAccounting":
        return [(acc, thick)]
    if g is None:
        return None
    if rule == "doubleAccounting":
        return [(acc, thick), (acc + g, thick)]
    u = math.floor(p * -t["underline_position"] / em)
    drop = 1 if p == 12 else 0
    if rule == "double":
        return [(u - g + drop, thick), (u + drop, thick)]
    return [(u if g == 2 else u - g, thick)]  # single


def other(face: str, p: float, rule: str) -> Lines:
    """rule_metrics' rule from the font's tables, at the screen size given."""
    t = tables(FILES[face])
    em = t["upem"]
    line = max(1, math.floor(p * t["underline_thickness"] / em))
    if rule == "strike":
        top = math.floor(p * t["strikeout_position"] / em + 0.5)
        return [(-top, max(1, math.floor(p * t["strikeout_size"] / em + 0.5)))]
    acc = math.floor(p * t["win_descent"] / em + 0.5)
    u = math.floor(p * -t["underline_position"] / em)
    if rule == "single":
        return [(u, line)]
    if rule == "singleAccounting":
        return [(acc, line)]
    double = max(1, math.floor(p * min(t["underline_thickness"], t["strikeout_size"]) / em))
    if rule == "doubleAccounting":
        return [(acc, double), (acc + 2 * double, double)]
    first = u
    while first + 3 * double > acc:
        first -= double
    return [(first, double), (first + 2 * double, double)]


def candidates(face: str, size: float, rule: str, zoom: int) -> dict[str, Lines | None]:
    p = size * zoom / 100
    whole = {"round": math.floor(p + 0.5), "floor": math.floor(p), "ceil": math.ceil(p)}
    out: dict[str, Lines | None] = {}
    for name in CANDIDATES:
        q: float = p if name == "exact" else whole[name]
        if face == "Courier New":
            out[name] = courier(q, rule, gap(math.floor(p + 0.5) if name == "exact" else int(q)))
        else:
            out[name] = other(face, q, rule)
    if face == "Times New Roman" and rule == "double":
        lines = out["exact"]
        assert lines is not None
        wide = 36.5 <= size <= 40 and zoom == 100
        top, th = lines[0]
        out = {key: [(top, th), (top + (3 if wide else 2) * th, th)] for key in CANDIDATES}
    return out


def _text(lines: Lines | None) -> str:
    return "" if lines is None else " ".join(f"{a:g}/{b:g}" for a, b in lines)


def predict() -> None:
    PREDICTIONS.mkdir(parents=True, exist_ok=True)
    for suite in SUITES:
        path = PREDICTIONS / f"{suite}.tsv"
        with path.open("w", encoding="utf-8") as out:
            out.write(
                f"# predictions for {suite}, written before its run (capture/courier_predict.py)\n"
            )
            out.write("# per rule, top/thickness in screen points below the baseline\n")
            out.write("zoom\ttag\t" + "\t".join(CANDIDATES) + "\n")
            for zoom in zooms(suite):
                for c in cases(suite):
                    if c.tag.startswith("filler"):
                        continue
                    found = candidates(c.face, c.size, c.rule, zoom)
                    out.write(
                        f"{zoom}\t{c.tag}\t" + "\t".join(_text(found[n]) for n in CANDIDATES) + "\n"
                    )
        print(f"wrote {path}")


def _measured(suite: str) -> dict[tuple[int, str], tuple[Lines, float]]:
    """(zoom, tag) -> (lines, the cell's last ink row), in screen points below the baseline."""
    raw: dict[tuple[int, str, float], dict[str, tuple[int, int, list[tuple[int, int]], int]]]
    raw = defaultdict(dict)
    with output(suite).open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "zoom\t")):
                continue
            zoom, tag, height, ppp, base, measured = line.rstrip("\n").split("\t")
            *face, size, rule = tag.split()
            runs = [(int(a), int(b)) for a, b in (r.split("-") for r in measured.split())]
            runs = [r for r in runs if r[0] > 0]  # from the row's top: the row above's line
            last = round(float(height) * int(ppp)) - 3
            raw[(int(zoom), " ".join(face), float(size))][rule] = (
                int(base) if base else -1, int(ppp), runs, last,
            )  # fmt: skip
    out = {}
    for (zoom, face, size), rules in raw.items():
        plain, acc = rules["strike"][0] + 1, rules["singleAccounting"][0] + 1
        for rule, (_, ppp, runs, last) in rules.items():
            b = acc if "Accounting" in rule else plain
            k = ppp / (zoom / 100)  # device pixels per screen point
            lines = [((a - b) / k, (e - a + 1) / k) for a, e in runs]
            out[(zoom, f"{face} {size:g} {rule}")] = (lines, (last - b + 1) / k)
    return out


def _shown(lines: Lines, bottom: float) -> Lines:
    return [(top, min(th, bottom - top)) for top, th in lines if top < bottom]


def score() -> None:
    for suite in SUITES:
        measured = _measured(suite)
        counts: dict[str, Counter[str]] = defaultdict(Counter)
        misses: dict[str, list[str]] = defaultdict(list)
        with (PREDICTIONS / f"{suite}.tsv").open(encoding="utf-8") as f:
            for line in f:
                if line.startswith(("#", "zoom\t")):
                    continue
                zoom, tag, *preds = line.rstrip("\n").split("\t")
                *face, size, _ = tag.split()
                p = float(size) * int(zoom) / 100
                group = f"{' '.join(face)} " + (
                    "whole P"
                    if p == int(p) and p <= 72
                    else "odd P above 72"
                    if p == int(p)
                    else "P not whole"
                )
                got, bottom = measured[(int(zoom), tag)]
                counts[group]["rows"] += 1
                for name, text in zip(CANDIDATES, preds, strict=True):
                    if not text:
                        continue
                    want = [tuple(float(v) for v in x.split("/")) for x in text.split()]
                    ok = got == _shown(want, bottom)  # type: ignore[arg-type]
                    counts[group][name] += ok
                    if not ok and name == "round":
                        misses[group].append(f"{zoom} {tag}: got {got} want {want}")
        print(suite)
        for group, c in sorted(counts.items()):
            print(f"  {group}: {dict(c)}")
        for group, ms in sorted(misses.items()):
            print(f"  misses of round, {group} ({len(ms)}):")
            for m in ms[:15]:
                print(f"    {m}")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if argv == ["predict"]:
        predict()
    elif argv == ["score"]:
        score()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
