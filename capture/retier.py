"""Reassign the ``tier`` column of an existing suite TSV from the current rule.

Usage: ``uv run python -m capture.retier <suite> [<suite> ...]`` or ``--all``. No recapture; the
file is rewritten in place with its header's ``tier_rule`` line updated.
"""

from __future__ import annotations

import sys
from pathlib import Path

from capture.cases import DEFAULT_FAST_CAP, FULL_ONLY, SUITES
from capture.schema import read_rows
from capture.tiers import assign_tiers, rule_description

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "tests" / "captures"


def fast_cap_for(suite: str) -> int:
    """The suite's fast-tier cap: from its definition, the full-only table, or the default."""
    if suite in FULL_ONLY:
        return FULL_ONLY[suite]
    if suite in SUITES:
        return SUITES[suite]().fast_cap
    return DEFAULT_FAST_CAP


def retier(path: Path) -> tuple[int, int]:
    """Rewrite one suite file; returns (fast rows, total rows)."""
    header, rows = read_rows(path)
    cap = fast_cap_for(header["suite"])
    rows = assign_tiers(rows, cap)
    header["tier_rule"] = rule_description(cap)
    from capture.capture import write_suite

    path = write_suite(path, header, rows)
    fast = sum(1 for r in rows if r.tier == "fast")
    print(f"{path.name}: {fast} fast of {len(rows)} (cap {cap})")
    return fast, len(rows)


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if argv == ["--all"]:
        paths = sorted(
            p
            for p in OUT_DIR.iterdir()
            if p.name.endswith((".tsv", ".tsv.gz")) and ".measure." not in p.name
        )
    else:
        paths = []
        for name in argv:
            candidates = [OUT_DIR / f"{name}.tsv", OUT_DIR / f"{name}.tsv.gz"]
            paths.extend(c for c in candidates if c.exists())
    total_fast = 0
    for path in paths:
        total_fast += retier(path)[0]
    print(f"fast tier total: {total_fast}")


if __name__ == "__main__":
    main(sys.argv[1:])
