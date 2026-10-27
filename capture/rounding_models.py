"""Compare candidate models of Excel's 15-significant-digit rounding against the Excel captures.

Usage: ``uv run python -m capture.rounding_models``. Reads ``tests/captures/rounding_15digit.tsv``
and the float rows of ``capture2023.tsv.gz`` under ``0``, ``0.0``, ``0.00`` and thirty decimals,
applies each model to the exact binary value, formats with the format's decimal count using
half-up rounding, and prints how many rows each model gets wrong. The model with zero failures on
both sets is the one the package implements (``ARCHITECTURE.md`` §5, ``QUIRKS.md``).
"""

from __future__ import annotations

from collections.abc import Callable
from decimal import ROUND_DOWN, ROUND_HALF_DOWN, ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal, getcontext
from pathlib import Path

from capture.schema import Row, read_rows

ROOT = Path(__file__).resolve().parents[1]
CAPTURE_DIR = ROOT / "tests" / "captures"
FORMATS = {"'0'": 0, "'0.0'": 1, "'0.00'": 2, "'0.000000000000000000000000000000;@'": 30}

getcontext().prec = 80


def sig(d: Decimal, n: int, mode: str) -> Decimal:
    """Round ``d`` to ``n`` significant digits with the given rounding mode."""
    if d == 0:
        return d
    return d.quantize(Decimal(1).scaleb(d.adjusted() - n + 1), rounding=mode)


def hybrid(x: Decimal) -> Decimal:
    """Below 1e14: 15 digits, ties up. At or above: 16 digits then 15 with ties down."""
    if x < Decimal("1e14"):
        return sig(x, 15, ROUND_HALF_UP)
    return sig(sig(x, 16, ROUND_HALF_EVEN), 15, ROUND_HALF_DOWN)


MODELS: dict[str, Callable[[Decimal], Decimal]] = {
    "15 digits, ties up": lambda x: sig(x, 15, ROUND_HALF_UP),
    "15 digits, ties down": lambda x: sig(x, 15, ROUND_HALF_DOWN),
    "15 digits, ties even": lambda x: sig(x, 15, ROUND_HALF_EVEN),
    "truncate to 15 digits": lambda x: sig(x, 15, ROUND_DOWN),
    "16 digits even, then 15 ties down": lambda x: sig(
        sig(x, 16, ROUND_HALF_EVEN), 15, ROUND_HALF_DOWN
    ),
    "16 digits up, then 15 ties down": lambda x: sig(
        sig(x, 16, ROUND_HALF_UP), 15, ROUND_HALF_DOWN
    ),
    "17 digits even, then 15 ties down": lambda x: sig(
        sig(x, 17, ROUND_HALF_EVEN), 15, ROUND_HALF_DOWN
    ),
    "hybrid: ties up below 1e14, 16-then-15-down above": hybrid,
}


def render(d: Decimal, places: int, negative: bool) -> str:
    """Format like ``0``, ``0.0``, ...: half-up at ``places`` decimals, minus unless zero."""
    q = d.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    text = f"{q:f}"
    if negative and q != 0:
        text = "-" + text
    return text


def failures(model: Callable[[Decimal], Decimal], rows: list[Row]) -> list[tuple[Row, str]]:
    """Rows where the model's rendering differs from Excel's."""
    out: list[tuple[Row, str]] = []
    for row in rows:
        x = float(row.value_repr)
        got = render(model(Decimal(abs(x))), FORMATS[row.format], x < 0)
        if got != row.expected_text:
            out.append((row, got))
    return out


def main() -> None:
    """Print the failure count of every model on both data sets."""
    sets: dict[str, list[Row]] = {}
    grid = CAPTURE_DIR / "rounding_15digit.tsv"
    if grid.exists():
        sets["rounding_15digit"] = [r for r in read_rows(grid)[1] if r.format in FORMATS]
    sets["capture2023 floats"] = [
        r
        for r in read_rows(CAPTURE_DIR / "capture2023.tsv.gz")[1]
        if r.value_type == "float" and r.format in FORMATS
    ]
    for name, model in MODELS.items():
        parts = []
        for set_name, rows in sets.items():
            fails = failures(model, rows)
            parts.append(f"{set_name}: {len(fails)}/{len(rows)}")
            if fails and name.startswith("hybrid"):
                for row, got in fails[:12]:
                    print(
                        f"    {row.value_repr:28s} {row.format:8s} "
                        f"excel={row.expected_text} model={got}"
                    )
        print(f"{name:52s} " + "   ".join(parts))


if __name__ == "__main__":
    main()
