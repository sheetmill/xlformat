"""Time parsing and rendering, and measure memory, for the speed pass.

Usage: ``uv run python -m bench.timing``. Parse times bypass the cache (the compile step alone);
render times use a parsed format, one call per mode: width-less, character mode (``width``),
Excel mode (``column_width``) and Excel mode with a cell font differing from the Normal font.
Each figure is the best of several repeats, in microseconds per call. Memory: what 1,000
distinct compiled formats retain, and one compiled four-section accounting format.
"""

from __future__ import annotations

import gc
import timeit
import tracemalloc
from collections.abc import Callable

from xlformat import Font, NumberFormat, format_value
from xlformat._parser import _parse_cached

ACCOUNTING = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
PARSE = ("#,##0.00", "General", ACCOUNTING, "m/d/yy", "yyyy-mm-dd hh:mm:ss.000", "# ?/?")
RENDER = (
    ("#,##0.00", 1234.5),
    ("General", 1234567.891),
    (ACCOUNTING, -1234.5),
    ("m/d/yy", 45000),
    ("yyyy-mm-dd hh:mm:ss.000", 45000.123456),
    ("0.00E+00", 1234.5),
    ("# ?/?", 1.3),
)
MODES: dict[str, dict[str, object]] = {
    "none": {},
    "width=10": {"width": 10},
    "column_width=10": {"column_width": 10},
    "cell font": {
        "column_width": 10,
        "font": Font("Calibri", 14, bold=True),
        "normal_font": Font("Calibri", 11),
    },
}


def best(fn: Callable[[], object], number: int) -> float:
    """Microseconds per call, best of five repeats."""
    return min(timeit.repeat(fn, number=number, repeat=5)) / number * 1e6


def parse_times() -> dict[str, float]:
    """Microseconds to compile each format, the cache bypassed."""
    uncached = _parse_cached.__wrapped__
    return {fmt: best(lambda f=fmt: uncached(f), 2000) for fmt in PARSE}


def render_times() -> dict[tuple[str, str], float]:
    """Microseconds to render each (format, value) in each mode."""
    out: dict[tuple[str, str], float] = {}
    for fmt, value in RENDER:
        compiled = NumberFormat.parse(fmt)
        for mode, kwargs in MODES.items():
            out[(fmt, mode)] = best(lambda c=compiled, v=value, k=kwargs: c.format(v, **k), 5000)
    return out


def memory() -> tuple[float, float]:
    """KiB retained by 1,000 distinct compiled formats, and by one accounting format."""
    uncached = _parse_cached.__wrapped__
    gc.collect()
    tracemalloc.start()
    before = tracemalloc.take_snapshot()
    kept = [uncached(f'0.{"0" * (k % 30)}" u{k}"') for k in range(1000)]
    after = tracemalloc.take_snapshot()
    many = sum(s.size_diff for s in after.compare_to(before, "filename")) / 1024
    before = tracemalloc.take_snapshot()
    one = uncached(ACCOUNTING)
    after = tracemalloc.take_snapshot()
    single = sum(s.size_diff for s in after.compare_to(before, "filename")) / 1024
    tracemalloc.stop()
    del kept, one
    return many, single


def main() -> None:
    """Print every figure."""
    format_value(1, "0", column_width=10)  # import and warm the font tables
    print("parse (cache bypassed), µs")
    for fmt, t in parse_times().items():
        print(f"  {t:7.1f}  {fmt}")
    print("render, µs per call:", ", ".join(MODES))
    times = render_times()
    for fmt, _ in RENDER:
        print("  " + "  ".join(f"{times[(fmt, m)]:6.1f}" for m in MODES) + f"  {fmt}")
    many, single = memory()
    print(f"memory: 1,000 compiled formats {many:,.0f} KiB; one accounting format {single:.1f} KiB")


if __name__ == "__main__":
    main()
