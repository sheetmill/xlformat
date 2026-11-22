"""Predictions for the ``cell_font_large`` capture, and their scorer.

    uv run python -m capture.padding_predict predict   # write capture/predictions/*.tsv.gz
    uv run python -m capture.padding_predict score     # score the capture against them

Written before the run. The question: does a cell font whose column padding ``2 * ceil(MDW / 4)
+ 1`` exceeds the Normal font's lose the difference from the cell's budget? Two candidates, each
rendering every width row in Excel mode as ``tests/test_captures.py`` does (the requested column
width, the cell font, the Normal font Calibri 11):

- ``current``: the package as it stands, the cell's budget the column's whatever the cell font.
- ``padding``: the budget less ``d = max(0, P(cell) - P(normal))`` points, and General's units
  less ``d`` over the cell font's digit, beside the italic overhang. Calibri 18 to 24 and Aptos
  Narrow 18 to 24 (digit 9 to 12, padding 7) lose 2 points; Arial 16 to 20 (digit 9 to 11) lose 2
  and Arial 24 (digit 13, padding 9) loses 4. Calibri 16 and Arial 14 (digit 8, padding 5) and the
  ``[normal]`` rows lose nothing: both candidates agree there, and those rows are controls.

Rows without a width (the default column, read through the clipboard) are not predicted: the
capture test does not compare cell-font rows in character mode.

Decision rule, fixed before the run: on the rows where the candidates differ, the candidate with
no misses is the Mac's rule (``padding`` sets the Mac preset's ``cell_font_padding``; ``current``
keeps it off). If both miss there, neither is adopted: the Mac keeps the current model and the
misses are written up. A miss on a row where the candidates agree is a separate finding about
the model, reported with the score.
"""

from __future__ import annotations

import gzip
import io
import sys
from collections import Counter, defaultdict
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from functools import lru_cache, partial
from pathlib import Path

from capture.cases import CellFont, Suite, cell_font_large
from capture.schema import read_rows
from xlformat import Font, format_value
from xlformat import _render as render
from xlformat._measure import (
    FontSpec,
    Measure,
    column_budget,
    column_padding,
    column_units,
    glyph_measure,
    overhang,
)

ROOT = Path(__file__).resolve().parents[1]
PREDICTIONS = ROOT / "capture" / "predictions"
SUITE = "cell_font_large"
CAPTURES = ROOT / "tests" / "captures"
NORMAL = Font("Calibri", 11)
CANDIDATES = ("current", "padding")


def _padded(
    column_width: float, font: FontSpec | None, normal_font: FontSpec | None
) -> tuple[float, Measure, float]:
    """``_render._excel_width`` with the cell font's extra padding taken off."""
    normal = font if normal_font is None else normal_font
    cell = glyph_measure(font)
    units = column_units(column_width, normal)
    if normal_font is not None:
        units *= glyph_measure(normal)("0") / cell("0")
    extra = max(0.0, column_padding(cell("0")) - column_padding(glyph_measure(normal)("0")))
    lost = overhang(font) + extra
    if lost:
        units -= lost / cell("0")
    return column_budget(column_width, normal) - lost, cell, units


@contextmanager
def _candidate(name: str) -> Iterator[None]:
    original = render._excel_width
    if name == "padding":
        render._excel_width = lru_cache(maxsize=4096)(_padded)
    try:
        yield
    finally:
        render._excel_width = original


def _font(cf: CellFont | None) -> Font:
    return NORMAL if cf is None else Font(cf.name, cf.size, bold=cf.bold, italic=cf.italic)


def _rows(suite: Suite) -> Iterator[tuple[str, float, Callable[[], str]]]:
    for case in suite.cases:
        font = _font(case.font)
        for width in suite.widths:
            if width is not None:
                run = partial(
                    format_value, case.value, case.format,
                    column_width=width, font=font, normal_font=NORMAL,
                )  # fmt: skip
                yield case.tag, width, run


def predict() -> None:
    PREDICTIONS.mkdir(parents=True, exist_ok=True)
    suite = cell_font_large()
    results: dict[str, list[tuple[str, float, str]]] = {}
    for name in CANDIDATES:
        with _candidate(name):
            results[name] = [(tag, w, run()) for tag, w, run in _rows(suite)]
    differ = 0
    path = PREDICTIONS / f"{SUITE}.tsv.gz"
    with gzip.GzipFile(path, "wb", mtime=0) as raw, io.TextIOWrapper(raw, "utf-8") as out:
        out.write(
            f"# predictions for {SUITE}, written before its run (capture/padding_predict.py)\n"
        )
        out.write("tag\twidth\t" + "\t".join(CANDIDATES) + "\n")
        for rows in zip(*(results[n] for n in CANDIDATES), strict=True):
            tag, width, _ = rows[0]
            texts = [r[2] for r in rows]
            differ += len(set(texts)) > 1
            out.write(f"{tag}\t{width!r}\t" + "\t".join(texts) + "\n")
    print(f"wrote {path}: {len(results['current'])} rows, {differ} where the candidates differ")


def score() -> None:
    path = CAPTURES / f"{SUITE}.tsv.gz"
    _, rows = read_rows(path if path.exists() else CAPTURES / f"{SUITE}.tsv")
    got = {(r.tag, float(r.width)): r.expected_text for r in rows if r.width}
    counts: dict[str, Counter[str]] = defaultdict(Counter)
    misses: dict[str, list[str]] = defaultdict(list)
    with gzip.open(PREDICTIONS / f"{SUITE}.tsv.gz", "rt", encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "tag\t")):
                continue
            tag, width, *preds = line.rstrip("\n").split("\t")
            shown = got[(tag, float(width))]
            group = "differ" if len(set(preds)) > 1 else "agree"
            font = tag[tag.rindex("[") + 1 : -1]
            counts[group]["rows"] += 1
            counts[f"{group} {font}"]["rows"] += 1
            for name, text in zip(CANDIDATES, preds, strict=True):
                ok = text == shown
                counts[group][name] += ok
                counts[f"{group} {font}"][name] += ok
                if not ok:
                    misses[f"{name}, {group}"].append(
                        f"{tag} @ {width}: got {shown!r} want {text!r}"
                    )
    for group, c in sorted(counts.items()):
        print(f"{group}: {dict(c)}")
    for key, ms in sorted(misses.items()):
        print(f"misses of {key} ({len(ms)}):")
        for m in ms[:20]:
            print(f"  {m}")


def main(argv: list[str]) -> None:
    {"predict": predict, "score": score}[argv[1]]()


if __name__ == "__main__":
    main(sys.argv)
