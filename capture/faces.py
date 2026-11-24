"""Autofit captures from which the per-face glyph advances are recovered (CAPTURING.md item 82).

Usage: ``uv run python -m capture.faces count|run <group> [suite ...] [--out <dir>]``
(``--out``: a directory under the repository instead of ``tests/captures/faces/``),
``<group>`` one of:

- ``check``: per family and size 10 and 13 (``faces_c<size>_<family>``; Normal font the
  family's regular face at the size, the four styles as cell fonts): every glyph of
  :data:`GLYPHS` ten times, a digit five-run, texts mixing ASCII and other glyphs, and long texts
  across the autofit bonus's steps. Also Segoe UI and MS Sans Serif, faces Excel for Mac lacks.
- ``recover``: per family (and Segoe UI) and size 383, 397 and 409
  (``faces_r<size>_<family>``): every ASCII glyph as ``T + g * N + T`` with the trigger ``T``
  (:data:`TRIGGER`, a non-ASCII glyph, so Excel sums the text unrounded), every other glyph as
  ``g * N``, the trigger as ``T * M`` and every digit's five- and ten-run. ``N`` and ``M`` are
  the longest runs under :data:`LIMIT_DIGITS` digit widths of the Normal font, estimated from
  the ``check`` captures at 13 pt (committed), so that no text reaches the autofit bonus.
- ``padding``: per size (``faces_p<size>``), every face of every family (and Segoe UI) as a
  cell font: ``0``
  five and ten times. Twice the five-run less the ten-run is the autofit padding, which gives
  each italic face's overhang.

- ``switch``: per size from 193 to 320 (``faces_s<size>``; Normal Calibri at the size), in six
  faces, ``0`` and the six ASCII glyphs whose rounding per glyph and rounding of the run's sum
  disagree most, five times each: where Excel stops rounding each ASCII glyph.

Each suite is one workbook; ``run`` writes ``tests/captures/faces/<suite>.measure.tsv``
(gzipped above 50 KB) and drops the row-less suite file. ``capture/face_fit.py`` reads them.
"""

from __future__ import annotations

import gzip
import math
import sys
from collections.abc import Callable, Iterator
from pathlib import Path

from capture import cases as xc

ROOT = Path(__file__).resolve().parents[1]
FACES_DIR = ROOT / "tests" / "captures" / "faces"

FAMILIES = (
    "Calibri",
    "Arial",
    "Aptos Narrow",
    "Aptos",
    "Arial Black",
    "Arial Narrow",
    "Book Antiqua",
    "Bookman Old Style",
    "Calibri Light",
    "Cambria",
    "Candara",
    "Century Gothic",
    "Comic Sans MS",
    "Consolas",
    "Courier",
    "Courier New",
    "Garamond",
    "Georgia",
    "Haettenschweiler",
    "Helvetica",
    "Impact",
    "Lucida Console",
    "Rockwell",
    "Tahoma",
    "Times New Roman",
    "Trebuchet MS",
    "Verdana",
)
ABSENT = ("Segoe UI", "MS Sans Serif")  # names Excel for Mac has no face for
# Segoe UI autofits as none of FAMILIES (the check captures), so its advances are recovered as
# Excel measures them; MS Sans Serif autofits exactly as Times New Roman
MEASURED = (*FAMILIES, "Segoe UI")
STYLES = ("regular", "bold", "italic", "bold italic")
# space, printable ASCII, then the cp1252 code page's other characters (Latin-1 and its additions)
GLYPHS = tuple(chr(c) for c in range(0x20, 0x7F)) + tuple(
    bytes([b]).decode("cp1252")
    for b in range(0x80, 0x100)
    if b not in (0x81, 0x8D, 0x8F, 0x90, 0x9D)
)
TRIGGER = "é"
CHECK_SIZES = (10, 13)
RECOVER_SIZES = (409, 397, 383)
LIMIT_DIGITS = 13  # the autofit bonus starts at 15 digit widths of the Normal font
DIGIT_RUNS = tuple(d * k for d in "0123456789" for k in (5, 10))
MIXED = (
    "Müller",
    "1.234,56 €",
    "naïve résumé",
    "£1,234.50",
    "25°C",
    "± 0.5 µm",
    "Ångström",
    "café au lait",
    "½ off",
    "© 2026 Société",
    "Zürich\u2013Genève",
    "œuvre … fin",
)
PANGRAM = "The quick brown fox jumps over the lazy dog, then naps."
LONG = tuple("0" * k for k in range(14, 32)) + tuple(PANGRAM[:k] for k in range(20, 56, 5))
# whole sizes from 6 to 30 and steps to the 192 pt cap, each its own workbook (Normal: Calibri at
# the size); sizes 1 to 5 and the half sizes share one workbook each (Normal: Calibri 6, 11)
PADDING_SIZES = (*range(6, 31), 32, 36, 40, 48, 56, 64, 72, 96, 128, 160, 192)
PADDING_GROUPS = {"faces_p1to5": (6, (1, 2, 3, 4, 5)), "faces_phalf": (11, (8.5, 9.5, 10.5, 11.5))}


def _key(family: str) -> str:
    return family.lower().replace(" ", "_")


def _font(family: str, size: float, style: str) -> xc.CellFont:
    return xc.CellFont(family, size, "bold" in style, "italic" in style)


def _suite(name: str, normal: str, size: int, texts: dict[str, list[str]]) -> xc.Suite:
    """One workbook: per style key (``family size style``) its texts, each in its own column."""
    cases = []
    for label, ts in texts.items():
        family, cell_size, style = label.split("|")
        font = _font(family, float(cell_size), style)
        for k, t in enumerate(dict.fromkeys(ts)):
            cases.append(xc.Case(name, f"s{k} [{font.label}]", t, "@", font=font))
    return xc.Suite(name, tuple(cases), widths=(), measure=True, font_size=size, font_name=normal)


def check_suite(family: str, size: int) -> xc.Suite:
    texts = [g * 10 for g in GLYPHS] + ["0" * 5, *MIXED, *LONG]
    return _suite(
        f"faces_c{size}_{_key(family)}",
        family,
        size,
        {f"{family}|{size}|{style}": texts for style in STYLES},
    )


def read_measure(path: Path) -> tuple[str, dict[tuple[str, str], float]]:
    """A measure file's header font and its points by (cell font label, text)."""
    opener = gzip.open if path.suffix == ".gz" else open
    with opener(path, "rt", encoding="utf-8") as f:
        lines = f.read().split("\n")
    font = lines[0].split("font: ", 1)[1].strip()
    out: dict[tuple[str, str], float] = {}
    for line in lines[2:]:
        if not line:
            continue
        fields = line.split("\t")
        out[(fields[5] if len(fields) > 5 else "", eval_text(fields[1]))] = float(fields[4])
    return font, out


def eval_text(value_repr: str) -> str:
    import ast

    text = ast.literal_eval(value_repr)
    assert isinstance(text, str)
    return text


def measure_path(suite: str) -> Path:
    plain = FACES_DIR / f"{suite}.measure.tsv"
    return plain if plain.exists() else plain.with_name(plain.name + ".gz")


def estimates(family: str, style: str) -> dict[str, float]:
    """Each glyph's advance per point of size, an upper estimate from the 13 pt check capture.

    The ten-run's width less the five-run's twice-over padding, over 10 and 13, plus 0.06 em for
    the per-glyph rounding (half a point in 13) and an autofit bonus the run may carry.
    """
    _, widths = read_measure(measure_path(f"faces_c13_{_key(family)}"))
    label = _font(family, 13.0, style).label
    ten, five = widths[(label, "0" * 10)], widths[(label, "0" * 5)]
    pad = 2 * five - ten
    return {g: max(0.0, (widths[(label, g * 10)] - pad) / 130) + 0.06 for g in GLYPHS}


def recover_suite(family: str, size: int) -> xc.Suite:
    digit = estimates(family, "regular")["0"] - 0.06  # the Normal font's digit, per point
    budget = LIMIT_DIGITS * digit * size
    texts: dict[str, list[str]] = {}
    for style in STYLES:
        est = estimates(family, style)
        t = est[TRIGGER] * size
        m = max(1, math.floor(budget / t))
        ts = [*DIGIT_RUNS, TRIGGER * m]
        for g in GLYPHS:
            if g == TRIGGER:
                continue
            if g < "\x7f":
                n = max(1, math.floor((budget - 2 * t) / (est[g] * size)))
                ts.append(TRIGGER + g * n + TRIGGER)
            else:
                ts.append(g * max(1, math.floor(budget / (est[g] * size))))
        texts[f"{family}|{size}|{style}"] = ts
    return _suite(f"faces_r{size}_{_key(family)}", family, size, texts)


def padding_suite(name: str, normal_size: int, sizes: tuple[float, ...]) -> xc.Suite:
    texts = {
        f"{family}|{s:g}|{style}": ["0" * 5, "0" * 10]
        for s in sizes
        for family in MEASURED
        for style in STYLES
    }
    return _suite(name, "Calibri", normal_size, texts)


SWITCH_SIZES = tuple(range(193, 321))
SWITCH_FACES = ("Calibri", "Arial", "Times New Roman", "Verdana", "Candara", "Georgia")


def switch_suite(size: int) -> xc.Suite:
    """Five-runs where rounding each ASCII glyph and rounding the run's sum disagree most.

    From the recovered advances (``src/xlformat/_faces_mac.py``): per face the six ASCII glyphs
    whose ``5 * round(x) - round(5 * x)`` is largest in size at ``size`` (x the advance in
    points), plus ``0`` five times.
    """
    from xlformat import _measure

    texts: dict[str, list[str]] = {}
    for family in SWITCH_FACES:
        e, units = _measure._mac_face(family, "regular")
        ascii_glyphs = [chr(c) for c in range(0x21, 0x7F)]

        def gap(g: str, e: int = e, units: dict[str, float] = units) -> float:
            x = units[g] * size / e
            return abs(5 * math.floor(x + 0.5) - math.floor(5 * x + 0.5))

        chosen = sorted(ascii_glyphs, key=lambda g: (-gap(g), g))[:6]
        texts[f"{family}|{size}|regular"] = ["0" * 5, *(g * 5 for g in chosen)]
    return _suite(f"faces_s{size}", "Calibri", size, texts)


def group(name: str) -> Iterator[Callable[[], xc.Suite]]:
    """The suites of a group, built lazily (``recover`` reads the check captures)."""
    if name == "check":
        for size in CHECK_SIZES:
            for family in (*FAMILIES, *ABSENT):
                yield lambda f=family, s=size: check_suite(f, s)
    elif name == "recover":
        for family in MEASURED:
            for size in RECOVER_SIZES:
                yield lambda f=family, s=size: recover_suite(f, s)
    elif name == "padding":
        for size in PADDING_SIZES:
            yield lambda s=size: padding_suite(f"faces_p{s}", s, (s,))
        for suite, (normal, sizes) in PADDING_GROUPS.items():
            yield lambda n=suite, ns=normal, ss=sizes: padding_suite(n, ns, ss)
    elif name == "switch":
        for size in SWITCH_SIZES:
            yield lambda s=size: switch_suite(s)
    else:
        msg = f"unknown group {name!r}: check, recover, padding or switch"
        raise SystemExit(msg)


def run(suites: list[xc.Suite], out_dir: Path = FACES_DIR) -> None:
    from capture import capture as cap
    from capture.edition import excel_edition

    excel_edition()  # refuse before the first workbook rather than after it
    out_dir.mkdir(parents=True, exist_ok=True)
    cap.OUT_DIR = out_dir
    failed = []
    for k, suite in enumerate(suites, 1):
        xc.SUITES[suite.name] = lambda s=suite: s
        print(f"== {k}/{len(suites)} {suite.name}: {len(suite.cases)} columns", flush=True)
        try:
            cap.capture(suite.name)
        except Exception as error:  # keep the workbooks that ran
            print(f"FAILED {suite.name}: {error!r}", flush=True)
            failed.append(suite.name)
            continue
        (out_dir / f"{suite.name}.tsv").unlink(missing_ok=True)  # no clipboard rows
        path = out_dir / f"{suite.name}.measure.tsv"
        if path.stat().st_size > 50_000:
            path.with_name(path.name + ".gz").write_bytes(gzip.compress(path.read_bytes(), mtime=0))
            path.unlink()
    print("failed: " + (" ".join(failed) or "none"), flush=True)
    if failed:
        sys.exit(1)


def main(argv: list[str]) -> None:
    out_dir = FACES_DIR
    if "--out" in argv:  # a directory of its own to fetch: a fetched path is emptied first
        k = argv.index("--out")
        out_dir = ROOT / argv[k + 1]
        argv = argv[:k] + argv[k + 2 :]
    command, name, *only = argv
    suites = [make() for make in group(name)]
    if only:
        suites = [s for s in suites if s.name in only]
    if command == "count":
        columns = sum(len(s.cases) for s in suites)
        fonts = max(len({c.font for c in s.cases}) for s in suites)
        print(f"{len(suites)} workbooks, {columns} columns, at most {fonts} cell fonts a workbook")
    elif command == "run":
        run(suites, out_dir)
    else:
        msg = f"unknown command {command!r}: count or run"
        raise SystemExit(msg)


if __name__ == "__main__":
    main(sys.argv[1:])
