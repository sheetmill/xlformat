"""Diff two suites that share their cases, row by row, and classify the differences.

Usage: ``uv run python -m capture.diff_captures capture2023 capture2026``.

Rows are matched on (value_type, value_repr, format, width). Each difference is classified:
``text16`` when the older suite was written with 16-digit number texts (no ``numbers: repr``
header) and the value's double differs under that text, so the two captures saw different
doubles; ``drift`` otherwise, which is a change between Excel builds (or in the method).
"""

from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

from capture.schema import Row, read_rows

ROOT = Path(__file__).resolve().parents[1]


def _path(suite: str) -> Path:
    plain = ROOT / "tests" / "captures" / f"{suite}.tsv"
    return plain if plain.exists() else plain.with_suffix(".tsv.gz")


def _key(r: Row) -> tuple[str, str, str, str]:
    return (r.value_type, r.value_repr, r.format, r.width)


def _text16_differs(r: Row) -> bool:
    if r.value_type != "float":
        return False
    v = r.value
    assert isinstance(v, float)
    return float(f"{v:.16g}") != v


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    a, b = argv[0], argv[1]
    header_a, rows_a = read_rows(_path(a))
    header_b, rows_b = read_rows(_path(b))
    by_key = {_key(r): r for r in rows_b}
    counts: Counter[str] = Counter()
    examples: dict[str, list[str]] = defaultdict(list)
    by_format: Counter[tuple[str, str]] = Counter()
    for ra in rows_a:
        rb = by_key.get(_key(ra))
        if rb is None:
            counts["unmatched"] += 1
            continue
        counts["compared"] += 1
        if ra.expected_text == rb.expected_text:
            continue
        sixteen = header_a.get("numbers") != "repr" and _text16_differs(ra)
        kind = "text16" if sixteen else "drift"
        counts[kind] += 1
        by_format[(kind, ra.format)] += 1
        if len(examples[kind]) < 40:
            examples[kind].append(
                f"    {ra.format:28} {ra.value_type:5} {ra.value_repr:26} "
                f"{a}={ra.expected_text!r} {b}={rb.expected_text!r}"
            )

    def describe(name: str, header: dict[str, str]) -> str:
        excel = header.get("excel_version", "?")
        return f"{name} (Excel {excel}, numbers {header.get('numbers', '%.16g')})"

    print(f"{describe(a, header_a)} vs {describe(b, header_b)}")
    print(dict(counts))
    for kind in ("text16", "drift"):
        if counts[kind]:
            print(f"\n== {kind}: {counts[kind]} rows, by format:")
            for (k, fmt), n in sorted(by_format.items(), key=lambda kv: -kv[1]):
                if k == kind:
                    print(f"    {n:5}  {fmt}")
            print("  examples:")
            print("\n".join(examples[kind]))


if __name__ == "__main__":
    main(sys.argv[1:])
