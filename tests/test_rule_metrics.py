"""``rule_metrics`` against ``tests/captures/rules.layout.tsv`` (CAPTURING.md section 2 item 60).

The rows hold "HHHH" in six faces at 9, 11, 14, 20 and 36 points under each underline and a
strikethrough, at zoom 100 and 200: the baseline (the H stems' last ink row) and each rule's
first and last pixel row. A double underline's first line is open, so for it only the gap and the
thickness are compared.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from xlformat import RuleMetrics, rule_metrics

LAYOUT = Path(__file__).parent / "captures" / "rules.layout.tsv"
# rows the rule misses, with why: a reader that found no rule (the underline joins the stems),
# one-pixel reads at the top of 36-point rows at zoom 200, Aptos Narrow's accounting rule a point
# lower at 9 and 20, and lines thinner than the table gives at zoom 200
KNOWN_MISSES = {
    (100, "Calibri 9 single"),
    (100, "Aptos Narrow 9 singleAccounting"), (100, "Aptos Narrow 9 doubleAccounting"),
    (100, "Aptos Narrow 20 singleAccounting"), (100, "Aptos Narrow 20 doubleAccounting"),
    (200, "Arial 14 double"), (200, "Arial 14 doubleAccounting"),
    (200, "Arial 36 double"), (200, "Arial 36 doubleAccounting"),
    (200, "Times New Roman 20 single"), (200, "Times New Roman 20 double"),
    (200, "Times New Roman 20 doubleAccounting"), (200, "Times New Roman 36 doubleAccounting"),
    (200, "Times New Roman 36 strike"),
    (200, "Georgia 36 doubleAccounting"), (200, "Georgia 36 strike"),
}  # fmt: skip


def _expected(
    m: RuleMetrics, rule: str, got: list[tuple[float, float]]
) -> list[tuple[float, float]]:
    if rule == "strike":
        return [(-m.strikethrough_top, m.strikethrough_thickness)]
    top = m.underline_top if rule in ("single", "double") else m.accounting_top
    if rule == "double":
        top = got[0][0] if got else top  # the first line is open
    lines = [(top, m.underline_thickness)]
    if rule.startswith("double"):
        lines.append((top + m.double_gap, m.underline_thickness))
    return lines


def test_every_captured_rule() -> None:
    compared, misses, listed = 0, [], 0
    with LAYOUT.open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "zoom\t")):
                continue
            zoom, tag, _, ppp, base, measured = line.rstrip("\n").split("\t")
            if tag.startswith(("'", "12 ")) or tag.endswith(" none") or "Courier New" in tag:
                continue
            *face, size, rule = tag.split()
            m = rule_metrics((" ".join(face), float(size)), zoom=int(zoom))
            px, b = int(ppp), int(base) + 1 if base else 0
            runs = [tuple(int(v) for v in r.split("-")) for r in measured.split()]
            got = [((a - b) / px, (e - a + 1) / px) for a, e in runs]
            ok = got == _expected(m, rule, got)
            if (int(zoom), tag) in KNOWN_MISSES:
                listed += 1
                assert not ok, f"{tag} at zoom {zoom} is listed but now matches"
                continue
            compared += 1
            if not ok:
                misses.append((zoom, tag, got, _expected(m, rule, got)))
    assert not misses, f"{len(misses)} rules differ: {misses[:6]}"
    assert (compared, listed) == (234, 16)


def test_courier_new_has_no_rule_positions() -> None:
    with pytest.raises(LookupError, match="no rule positions for Courier New"):
        rule_metrics(("Courier New", 11))
    assert rule_metrics(("Calibri", 11)) == RuleMetrics(1.0, 1.0, 2.0, 3.0, 3.0, 1.0)
