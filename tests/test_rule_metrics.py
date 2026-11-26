"""``rule_metrics`` against the screen captures of rules (CAPTURING.md section 2 items 60 and 62).

The rows hold "HHHH" in a face, size and zoom under each underline and a strikethrough, as the
pixel rows of each line from the row's top. Baselines come from rows whose rule cannot touch the H
stems: the strikethrough row's for the plain rows, the single accounting row's for the accounting
rows. Rows without a rule and the text-color rows of ``rules.layout.tsv`` are skipped.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

import pytest

import xlformat._measure
from xlformat import RuleMetrics, rule_metrics

CAPTURES = Path(__file__).parent / "captures"
Lines = list[tuple[float, float]]

# rules.layout.tsv rows that differ: the Calibri 9 underline joins the H stems at zoom 100, so
# the reader finds no rule
RULES_UNREAD = {(100, "Calibri 9 single")}


Row = dict[str, tuple[Lines, float]]


def _rows(name: str) -> dict[tuple[int, str, float], Row]:
    """(zoom, face, size) -> rule -> ([(top, thickness)], the cell's last ink row), in sheet
    points below the baseline. A row's last two pixel rows are its gridline."""
    raw: dict[tuple[int, str, float], dict[str, tuple[int, int, list[tuple[int, int]], int]]]
    raw = defaultdict(dict)
    with (CAPTURES / name).open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "zoom\t")):
                continue
            zoom, tag, height, ppp, base, measured = line.rstrip("\n").split("\t")
            if tag.startswith(("'", "12 ")):
                continue
            *face, size, rule = tag.split()
            runs = [(int(a), int(b)) for a, b in (r.split("-") for r in measured.split())]
            runs = [r for r in runs if r[0] > 0]  # from the row's top: the row above's line
            raw[(int(zoom), " ".join(face), float(size))][rule] = (
                int(base) if base else -1, int(ppp), runs, round(float(height) * int(ppp)) - 3,
            )  # fmt: skip
    out: dict[tuple[int, str, float], Row] = {}
    for key, rules in raw.items():
        plain, accounting = rules["strike"][0] + 1, rules["singleAccounting"][0] + 1
        out[key] = {
            rule: ([((a - b) / ppp, (e - a + 1) / ppp) for a, e in runs], (last - b + 1) / ppp)
            for rule, (_, ppp, runs, last) in rules.items()
            if rule != "none"
            for b in [accounting if "Accounting" in rule else plain]
        }
    return out


def _expected(m: RuleMetrics, rule: str) -> Lines:
    if rule == "strike":
        return [(-m.strikethrough_top, m.strikethrough_thickness)]
    if rule == "single":
        return [(m.underline_top, m.underline_thickness)]
    if rule == "singleAccounting":
        return [(m.accounting_top, m.underline_thickness)]
    top = m.double_top if rule == "double" else m.accounting_top
    return [(top, m.double_thickness), (top + m.double_gap, m.double_thickness)]


def _shown(lines: Lines, bottom: float) -> Lines:
    """The lines as a cell shows them: a line reaching past the cell's bottom is cut there (at
    zoom 200 a double accounting underline's second line can run into the row below)."""
    return [(top, min(th, bottom - top)) for top, th in lines if top < bottom]


def _close(lines: Lines) -> list[tuple[float, float]]:
    """Lines to six decimals: at zoom 150 a third of a point is reached by two float routes."""
    return [(round(top, 6), round(th, 6)) for top, th in lines]


def _differing(name: str) -> tuple[int, set[tuple[int, str]]]:
    compared, differ = 0, set()
    for (zoom, face, size), rules in _rows(name).items():
        m = rule_metrics((face, size), zoom=zoom, platform="mac")
        for rule, (got, bottom) in rules.items():
            compared += 1
            want = _expected(m, rule)
            # whole, or cut at the row's bottom, which at zoom 150 the row's height in points
            # gives only roughly (Excel rounds each row to whole pixels)
            if _close(got) not in (_close(want), _close(_shown(want, bottom))):
                differ.add((zoom, f"{face} {size:g} {rule}"))
    return compared, differ


def test_rules_capture() -> None:
    compared, differ = _differing("rules.layout.tsv")
    assert differ == RULES_UNREAD
    assert compared == 300


# rules_double.layout.tsv rows that differ: the Calibri 9 underlines join the stems at zoom 100
DOUBLE_UNREAD = {(100, "Calibri 9 single"), (100, "Calibri 9 double")}

# screen sizes where the rule alone misses a captured double underline (QUIRKS.md): its first
# line drawn higher, and in Arial at some sizes its lines a screen point thicker; Aptos Narrow 9
# from its accounting underline, drawn a point lower than its table gives
RULE_MISSES = {
    "Aptos Narrow": {9},
    "Arial": {
        29, 30, 31, 32, 33, 34, 35, 36, 38, 40, 46, 48, 50, 52, 54, 56, 58, 60, 62, 68, 70, 72,
        80, 88, 96, 104, 112, 120, 128, 136, 144,
    },
    "Calibri": {80, 88, 96, 104, 112, 120, 128, 136, 144},
    "Georgia": {128, 144},
    "Times New Roman": {46, 48, 62, 66, 68, 70, 88, 96, 104, 112, 120, 128, 136, 144},
}  # fmt: skip


def test_rules_double_capture_is_the_table() -> None:
    compared, differ = _differing("rules_double.layout.tsv")
    assert differ == DOUBLE_UNREAD
    assert compared == 1850


def test_rule_alone(monkeypatch: pytest.MonkeyPatch) -> None:
    """The rule from the font tables, the captured table set aside, on every captured double."""
    monkeypatch.setattr(xlformat._measure, "DRAWN", {})
    right, misses = 0, defaultdict(set)
    for (zoom, face, size), rules in _rows("rules_double.layout.tsv").items():
        got = rules["double"][0]
        if len(got) != 2:
            continue
        if got == _expected(rule_metrics((face, size), zoom=zoom, platform="mac"), "double"):
            right += 1
        else:
            misses[face].add(size * zoom / 100)
    assert misses == RULE_MISSES
    assert right == 299  # of 369 captured doubles; 70 rows at the 57 screen sizes listed


def test_rules_courier_capture_is_the_table() -> None:
    compared, differ = _differing("rules_courier.layout.tsv")
    # Courier New 6 at zoom 100: the reader finds no single or double underline
    assert differ == {(100, "Courier New 6 single"), (100, "Courier New 6 double")}
    assert compared == 1190


# rules_zoom.layout.tsv rows that differ. The reader: Courier New 8 has no line read; at zoom 150
# a double accounting underline's second line is cut at the row's bottom, and a strikethrough
# row reads a line of the row above. The rest are the rule's misses at a screen size that is not
# whole, which Excel draws as at the nearest whole one in all but these (QUIRKS.md): a line one
# screen point off, in Courier New at several sizes and in Calibri from screen size 51.
ZOOM_UNREAD = {
    (100, "Courier New 8 single"), (100, "Courier New 8 double"), (150, "Courier New 8 single"),
    (150, "Calibri 22.5 doubleAccounting"), (150, "Calibri 23 doubleAccounting"),
    (150, "Calibri 23.5 doubleAccounting"), (150, "Calibri 51.5 doubleAccounting"),
    (150, "Calibri 69.5 doubleAccounting"), (150, "Calibri 69.5 strike"),
    (150, "Calibri 70.5 strike"),
}  # fmt: skip
ZOOM_MISSES = {
    (150, "Courier New 17 double"), (150, "Courier New 23 single"), (150, "Courier New 23 double"),
    (150, "Courier New 35.5 single"), (150, "Courier New 35.5 double"),
    (150, "Courier New 51.5 single"), (150, "Courier New 51.5 double"),
    (150, "Courier New 69.5 double"), (150, "Courier New 70.5 doubleAccounting"),
    (150, "Calibri 11.5 doubleAccounting"), (150, "Calibri 22.5 double"),
    (150, "Calibri 34.5 single"), (150, "Calibri 34.5 double"), (150, "Calibri 51.5 double"),
    (150, "Calibri 52.5 single"), (150, "Calibri 52.5 double"),
    (150, "Calibri 52.5 singleAccounting"), (150, "Calibri 52.5 doubleAccounting"),
    (150, "Calibri 52.5 strike"), (150, "Calibri 70.5 single"), (150, "Calibri 70.5 double"),
    (150, "Calibri 70.5 singleAccounting"), (150, "Calibri 70.5 doubleAccounting"),
}  # fmt: skip


def test_rules_zoom_capture() -> None:
    """Screen sizes that are not whole: drawn as at the nearest whole one, but for the misses."""
    compared, differ = _differing("rules_zoom.layout.tsv")
    assert differ == ZOOM_UNREAD | ZOOM_MISSES
    assert compared == 400


def test_courier_new_rule_alone(monkeypatch: pytest.MonkeyPatch) -> None:
    """Courier New's rule with the captured table set aside: every captured line it reads."""
    monkeypatch.setattr(xlformat._measure, "DRAWN", {})
    compared, differ = _differing("rules_courier.layout.tsv")
    courier = {d for d in differ if "Courier New" in d[1]}
    assert courier == {(100, "Courier New 6 single"), (100, "Courier New 6 double")}
    assert compared == 1190


def test_courier_new_rule_positions() -> None:
    # the double hangs up from the single's table position; at 11 points the gap is 2
    # thicknesses, so the single stays at that position
    assert rule_metrics(("Courier New", 11), platform="mac") == RuleMetrics(
        2.0, 1.0, 2.0, 3.0, 3.0, 1.0, 0.0, 1.0
    )
    # at 72 points, zoom 200, the gap is wider and the single moves onto the first line
    m = rule_metrics(("Courier New", 72), zoom=200, platform="mac")
    assert (m.underline_top, m.double_top, m.double_gap) == (7.0, 7.0, 9.5)
    with pytest.raises(LookupError, match="measured at screen sizes 6 to 144"):
        rule_metrics(("Courier New", 72), zoom=250, platform="mac")
    with pytest.raises(LookupError, match="screen size 5"):
        rule_metrics(("Courier New", 5), platform="mac")
    assert rule_metrics(("Calibri", 11), platform="mac") == RuleMetrics(
        1.0, 1.0, 2.0, 3.0, 3.0, 1.0, 0.0, 1.0
    )
