"""Predictions for the ``column_half_points_*`` captures, and their scorer.

    uv run python -m capture.half_predict predict   # write capture/predictions/
    uv run python -m capture.half_predict score     # score the captures against them

Written before the run. The question: does Excel for Mac round a column of a half point
(``MDW * w + P`` ending in .5) to even, as the package did, or up, as Excel for Windows does with
a half pixel (CAPTURING.md section 2 item 66)? Each suite holds the cell-font values in its
Normal font at every width from 1 to 20 whose column is a half point. Two candidates:

- ``even``: the package as it stood before the capture, Python's ``round`` (2.5 to 2, 3.5 to 4).
- ``up``: a half rounds up.

Per column, the points Excel reports (``column_points`` in the header) and the dialog width it
reads back; per row, the displayed text. The candidates differ on the columns whose floor is even
(about half; the others are controls) and on the rows those columns change.

Decision rule, fixed before the run: the candidate with no miss on the columns where they differ
is the Mac's rule. If it is ``up``, one rule serves both platforms; if ``even``, Windows' rule stays
its own and QUIRKS.md says so. If both miss, neither is adopted and the misses are written up.

Outcome (CAPTURING.md section 2 item 67): ``up`` fits all 96 columns and 143 rows where the
candidates differ, ``even`` none, so a half point rounds up on the Mac as a half pixel does on
Windows, and the package now rounds it up.
"""

from __future__ import annotations

import math
import sys
from collections import Counter
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from capture.cases import SUITES
from capture.schema import read_rows
from xlformat import Font, format_value
from xlformat import _measure as measure
from xlformat import _render as render
from xlformat._platform import Platform

ROOT = Path(__file__).resolve().parents[1]
PREDICTIONS = ROOT / "capture" / "predictions" / "column_half_points.tsv"
CAPTURES = ROOT / "tests" / "captures"
NAMES = tuple(n for n in SUITES if n.startswith("column_half_points_"))
CANDIDATES = ("even", "up")


@contextmanager
def _candidate(name: str) -> Iterator[None]:
    original = measure.column_points

    def up(
        column_width: float, font: measure.FontSpec | None = None, platform: Platform = Platform.MAC
    ) -> float:
        mdw = measure._mdw(font, platform)
        pad = measure.column_padding(mdw)
        exact = column_width * (mdw + pad) if column_width < 1 else mdw * column_width + pad
        return float(math.floor(exact + 0.5))

    if name == "up":
        measure.column_points = up
    render._excel_width.cache_clear()
    try:
        yield
    finally:
        measure.column_points = original
        render._excel_width.cache_clear()


def predict() -> None:
    PREDICTIONS.parent.mkdir(parents=True, exist_ok=True)
    out: dict[str, list[str]] = {}
    for name in CANDIDATES:
        with _candidate(name):
            lines = []
            for suite_name in NAMES:
                suite = SUITES[suite_name]()
                normal = Font(suite.font_name, float(suite.font_size or 11))
                for width in suite.widths:
                    assert width is not None
                    points = measure.column_points(width, normal)
                    lines.append(f"{suite_name}\tcolumn\t{width!r}\t{points:g}")
                    for case in suite.cases:
                        text = format_value(
                            case.value, case.format, column_width=width, font=normal,
                            normal_font=normal,
                        )  # fmt: skip
                        lines.append(f"{suite_name}\t{case.tag}\t{width!r}\t{text}")
            out[name] = lines
    differ: Counter[str] = Counter()
    with PREDICTIONS.open("w", encoding="utf-8") as f:
        f.write("# predictions for column_half_points_*, written before the run "
                "(capture/half_predict.py)\n")  # fmt: skip
        f.write("suite\ttag\twidth\t" + "\t".join(CANDIDATES) + "\n")
        for even, rounded_up in zip(out["even"], out["up"], strict=True):
            *key, a = even.split("\t")
            b = rounded_up.split("\t")[-1]
            differ["columns" if key[1] == "column" else "rows"] += a != b
            f.write("\t".join(key) + f"\t{a}\t{b}\n")
    print(
        f"wrote {PREDICTIONS}: columns differing {differ['columns']},"
        f" rows differing {differ['rows']}"
    )


def score() -> None:
    got: dict[tuple[str, str, float], str] = {}
    for suite_name in NAMES:
        path = next(p for p in (CAPTURES / f"{suite_name}.tsv", CAPTURES / f"{suite_name}.tsv.gz")
                    if p.exists())  # fmt: skip
        header, rows = read_rows(path)
        letters = dict(kv.split("=") for kv in header["column_points"].split())
        widths = SUITES[suite_name]().widths  # the capture columns, in sheet order
        for letter, width in zip(sorted(letters, key=lambda c: (len(c), c)), widths, strict=True):
            assert width is not None
            got[(suite_name, "column", width)] = f"{float(letters[letter]):g}"
        for r in rows:
            got[(suite_name, r.tag, float(r.width))] = r.expected_text
    counts: dict[str, Counter[str]] = {"column": Counter(), "row": Counter()}
    misses: list[str] = []
    with PREDICTIONS.open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "suite\t")):
                continue
            suite_name, tag, width_text, *preds = line.rstrip("\n").split("\t")
            kind = "column" if tag == "column" else "row"
            group = "differ" if len(set(preds)) > 1 else "agree"
            shown = got[(suite_name, tag, float(width_text))]
            counts[kind][f"{group} rows"] += 1
            for name, text in zip(CANDIDATES, preds, strict=True):
                counts[kind][f"{group} {name}"] += text == shown
                if text != shown and len(misses) < 20:
                    misses.append(
                        f"{name}: {suite_name} {tag} @ {width_text}: {shown!r} not {text!r}"
                    )
    for kind, c in counts.items():
        print(kind, dict(c))
    for m in misses:
        print(" ", m)


def main(argv: list[str]) -> None:
    {"predict": predict, "score": score}[argv[1]]()


if __name__ == "__main__":
    main(sys.argv)
