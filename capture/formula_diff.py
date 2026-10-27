"""Compare constant rows with formula rows in a captured suite, and both with the package.

Usage: ``uv run python -m capture.formula_diff [suite]`` (default ``clock_formulas``).

Rows are paired on (format, value_repr): the constant row has a descriptive tag, the formula
row a ``formula:`` tag. For each pair the script prints whether Excel showed the same text for
both, and what ``format_value`` gives for the Python double. The summary counts tell whether
the file-constant path and the IEEE-formula path disagree (the parser hypothesis) and which
path the package reproduces.
"""

from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

from capture.schema import Row, read_rows
from xlformat import format_value

ROOT = Path(__file__).resolve().parents[1]


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    suite = argv[0] if argv else "clock_formulas"
    path = ROOT / "tests" / "captures" / f"{suite}.tsv"
    if not path.exists():
        path = path.with_suffix(".tsv.gz")
    _, rows = read_rows(path)
    pairs: dict[tuple[str, str], dict[str, Row]] = defaultdict(dict)
    for r in rows:
        kind = "formula" if r.tag.startswith("formula:") else "constant"
        pairs[(r.format, r.value_repr)][kind] = r
    counts: Counter[str] = Counter()
    groups: dict[str, list[str]] = defaultdict(list)
    for (fmt, value_repr), kinds in pairs.items():
        c, f = kinds.get("constant"), kinds.get("formula")
        if c is None or f is None:
            counts["unpaired"] += 1
            continue
        ours = format_value(c.value, c.format_string)
        same = c.expected_text == f.expected_text
        key = (
            ("const=formula" if same else "const!=formula")
            + (" ours=formula" if ours == f.expected_text else " ours!=formula")
            + (" ours=const" if ours == c.expected_text else " ours!=const")
        )
        counts[key] += 1
        if not same or ours != f.expected_text:
            groups[key].append(
                f"  {fmt:20s} {c.tag:28s} {value_repr:24s} const={c.expected_text!r} "
                f"formula={f.expected_text!r} ours={ours!r}  [{f.tag}]"
            )
    print(f"{len(pairs)} pairs")
    for key, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"{n:5d}  {key}")
    for key, lines in groups.items():
        print(f"\n== {key} ({len(lines)})")
        print("\n".join(lines))


if __name__ == "__main__":
    main(sys.argv[1:])
