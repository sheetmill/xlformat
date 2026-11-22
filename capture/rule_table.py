"""Write ``src/xlformat/_rules_drawn.py`` from the screen captures of rules.

    uv run python -m capture.rule_table

Each face and screen size ``P = size * zoom / 100`` of the capture gives the lines where Excel drew
them, in half screen points (a device pixel on the display of scale 2 the capture was made on): the
single underline's top and thickness, a double underline's first line, thickness and gap, the
accounting underline's top, and the strikethrough's top (above the baseline) and thickness. The
captures are ``rules_double``, ``rules_courier`` and ``rules_zoom`` (CAPTURING.md section 2 items 62
and 63), read in that order; where two reach the same face and screen size the first is kept, and
within a capture zoom 100 (equal screen sizes draw equal lines). ``COURIER_GAP`` holds Courier New's
double underline gap at each whole screen size read. A field is left out (``None``) where its row
does not read as whole lines: one line for a single, a single accounting underline and a
strikethrough, two of one thickness for a double, a line from the row's first pixel row (the row
above's) set aside. Baselines: the strikethrough row's for the plain rows and the single accounting
row's for the accounting rows, which no rule touches.
"""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CAPTURES = tuple(
    ROOT / "tests" / "captures" / f"{suite}.layout.tsv"
    for suite in ("rules_double", "rules_courier", "rules_zoom")
)
TARGET = ROOT / "src" / "xlformat" / "_rules_drawn.py"

Lines = list[tuple[int, int]]


def read(capture: Path) -> dict[tuple[int, str, float], dict[str, tuple[int, int, Lines]]]:
    """(zoom, face, size) -> rule -> (baseline px, pixels per point, [(first, last) px])."""
    rows: dict[tuple[int, str, float], dict[str, tuple[int, int, Lines]]] = defaultdict(dict)
    with capture.open(encoding="utf-8") as f:
        for line in f:
            if line.startswith(("#", "zoom\t")):
                continue
            zoom, tag, _, ppp, base, measured = line.rstrip("\n").split("\t")
            *face, size, rule = tag.split()
            runs = [(int(a), int(b)) for a, b in (r.split("-") for r in measured.split())]
            # a line from the row's first pixel row is the row above's double accounting
            # underline, which crosses into this row at large sizes
            runs = [r for r in runs if r[0] > 0]
            rows[(int(zoom), " ".join(face), float(size))][rule] = (
                int(base) if base else -1,
                int(ppp),
                runs,
            )
    return rows


def drawn(zoom: int, rules: dict[str, tuple[int, int, Lines]]) -> tuple[int | None, ...] | None:
    """The fields of one face and size in half screen points, or None where nothing reads."""
    if rules["strike"][0] < 0 or rules["singleAccounting"][0] < 0:
        return None
    plain, accounting = rules["strike"][0] + 1, rules["singleAccounting"][0] + 1

    def lines(rule: str, base: int) -> Lines:
        _, ppp, runs = rules[rule]
        scale = 2 * zoom / 100 / ppp  # device pixels -> half screen points
        return [(round((a - base) * scale), round((b - a + 1) * scale)) for a, b in runs]

    single, double = lines("single", plain), lines("double", plain)
    acc, strike = lines("singleAccounting", accounting), lines("strike", plain)
    one = len(single) == 1
    two = len(double) == 2 and double[0][1] == double[1][1]
    return (
        single[0][0] if one else None,
        single[0][1] if one else None,
        double[0][0] if two else None,
        double[0][1] if two else None,
        double[1][0] - double[0][0] if two else None,
        acc[0][0] if len(acc) == 1 else None,
        -strike[0][0] if len(strike) == 1 else None,
        strike[0][1] if len(strike) == 1 else None,
    )


def table() -> dict[tuple[str, float], tuple[int | None, ...]]:
    out: dict[tuple[str, float], tuple[int | None, ...]] = {}
    for capture in CAPTURES:
        # zoom 100 first, then by face and size
        for (zoom, face, size), rules in sorted(read(capture).items(), key=_zoom_100_first):
            key = (face.lower(), size * zoom / 100)
            if key in out:
                continue
            fields = drawn(zoom, rules)
            if fields is not None and any(v is not None for v in fields):
                out[key] = fields
    return out


def _zoom_100_first(item: tuple[tuple[int, str, float], object]) -> tuple[bool, int, str, float]:
    zoom, face, size = item[0]
    return (zoom != 100, zoom, face, size)


def courier_gap(entries: dict[tuple[str, float], tuple[int | None, ...]]) -> dict[int, float]:
    """Courier New's double underline gap in screen points at each whole screen size: as drawn,
    and from the accounting rows (a row grows by the gap and the thickness) at a size whose
    double was not read on screen. The two agree at every size both reach."""
    drawn_gap = {
        int(p): fields[4] / 2
        for (face, p), fields in entries.items()
        if face == "courier new" and p == int(p) and fields[4] is not None
    }
    heights = {}
    with (ROOT / "tests" / "captures" / "underline_height.layout.tsv").open(encoding="utf-8") as f:
        for line in f:
            if line.startswith("Courier New ") and line.split("\t")[0].endswith(" 1"):
                tag, height = line.rstrip("\n").split("\t")
                _, _, size, underline, _ = tag.split()
                heights[(int(size), underline)] = float(height)
    for size in sorted({s for s, _ in heights}):
        thickness = max(1, (size * 84) // 2048)
        rows = heights[(size, "singleAccounting")] - heights[(size, "none")] - thickness
        assert drawn_gap.get(size, rows) == rows, (size, drawn_gap[size], rows)
        drawn_gap.setdefault(size, rows)
    return dict(sorted(drawn_gap.items()))


def write() -> None:
    entries = table()
    lines = [
        '"""Where Excel for Mac drew underlines and a strikethrough, by face and screen size.',
        "",
        "Generated by ``capture/rule_table.py`` from the ``rules_double``, ``rules_courier`` and",
        "``rules_zoom`` captures (CAPTURING.md section 2 items 62 and 63); do not edit by hand.",
        '"""',
        "",
        "# (face, screen size) -> (underline top, underline thickness, a double underline's",
        "# first line, its thickness, its gap, accounting underline top, strikethrough top above",
        "# the baseline, strikethrough thickness), in half screen points below the baseline; None",
        "# where not read",
        "DRAWN: dict[tuple[str, float], tuple[int | None, ...]] = {",
    ]
    lines += [f"    ({face!r}, {p!r}): {fields!r}," for (face, p), fields in entries.items()]
    lines.append("}  # fmt: skip")
    gaps = courier_gap(entries)
    lines += [
        "",
        "# whole screen size -> Courier New's double underline gap in screen points, as drawn",
        "COURIER_GAP: dict[int, float] = {",
        *(f"    {p}: {g!r}," for p, g in gaps.items()),
        "}  # fmt: skip",
    ]
    TARGET.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {TARGET}: {len(entries)} face and screen sizes")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if argv:
        sys.exit(__doc__)
    write()


if __name__ == "__main__":
    main(sys.argv[1:])
