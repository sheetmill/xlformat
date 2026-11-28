"""A face's height at any size, and rows mixing faces: the rows Excel fits.

    uv run python -m capture.font_heights build      # the workbooks, no Excel
    uv run python -m capture.font_heights predict    # write the predictions, before a run
    uv run python -m capture.font_heights run        # capture every face, write the layout file
    uv run python -m capture.font_heights wrapped    # predict, then capture the wrapped rows
    uv run python -m capture.font_heights underlines # predict, then capture the underlined rows
    uv run python -m capture.font_heights accounting # Times New Roman and Georgia, 36 to 41 points
    uv run python -m capture.font_heights more       # the same design for MORE_FACES

One workbook per face (``FACES``) whose Normal style is that face at 1 point, so that no row is
held up by the Normal font. Rows: a number in the face at every size from 1 to 72 points in half
steps; then the face at 9, 11, 14 and 20 points beside Calibri at each table size (9 to 24) in
another column of the same row. Excel fits every row when it opens the file; the script reads
each row's height and the standard height through ``capture/row_heights.applescript``, closes
the workbook without saving and restores alerts. No clipboard. Writes
``tests/captures/font_height.layout.tsv``.

``wrapped``: two workbooks, Normal style Calibri 20 and Calibri 9, each holding every face at 9,
11, 14 and 20 points with one to three wrapped lines ("x" a line); it writes its predictions to
``capture/build/`` first, then ``tests/captures/wrapped_height.layout.tsv``.

``more``: the design of ``run`` for :data:`MORE_FACES`, the families xlformat measures widths for
on the Mac beyond the six of :data:`FACES`, and "arial" in lower case (whether Excel matches a
face's name ignoring case); writes ``tests/captures/font_height_more.layout.tsv``.

``underlines``: one workbook per face, Normal style Calibri at 1 point, holding the face at
every whole size from 6 to 72 points under each underline style (none, single, double, single
and double accounting), one line, and three wrapped lines of "H" without and with a single
accounting underline; predictions first, then ``tests/captures/underline_height.layout.tsv``.
"""

from __future__ import annotations

import datetime as dt
import math
import sys
from pathlib import Path

from capture.build import BUILD_DIR, WORKBOOK_DIR
from capture.edition import edition_line, excel_build
from capture.excel_watch import answers, osascript_or_exit

SCRIPT = Path(__file__).with_name("row_heights.applescript")
DIRECTORY = WORKBOOK_DIR / "font_heights"
OUTPUT = Path(__file__).resolve().parents[1] / "tests" / "captures" / "font_height.layout.tsv"
PREDICTIONS = BUILD_DIR / "font_heights.predictions.tsv"
FACES = ("Calibri", "Arial", "Aptos Narrow", "Times New Roman", "Courier New", "Georgia")
MORE_FACES = (
    "Aptos", "Arial Black", "Arial Narrow", "Book Antiqua", "Bookman Old Style", "Calibri Light",
    "Cambria", "Candara", "Century Gothic", "Comic Sans MS", "Consolas", "Courier", "Garamond",
    "Haettenschweiler", "Helvetica", "Impact", "Lucida Console", "Rockwell", "Tahoma",
    "Trebuchet MS", "Verdana", "arial",
)  # fmt: skip
MORE_OUTPUT = OUTPUT.with_name("font_height_more.layout.tsv")
SIZES = tuple(s / 2 for s in range(2, 145))  # 1 to 72 in half steps
PAIR_SIZES = (9, 11, 14, 20)
CALIBRI_SIZES = (9, 10, 11, 12, 14, 16, 18, 20, 24)
# the slope of each face's height, F(S) = floor(a * ceil(S) + 1.5): the middle of the window of
# slopes the captured rows from 7.5 points allow (CAPTURING.md section 2 item 57)
SLOPES = {
    "Calibri": 1.2689,
    "Arial": 1.2122,
    "Aptos Narrow": 1.2815,
    "Times New Roman": 1.2162,
    "Courier New": 1.3007,
    "Georgia": 1.2190,
}


def rows(face: str) -> list[tuple[str, tuple[tuple[str, float], ...]]]:
    """(tag, ((face, size) per cell)) for each row of the face's workbook, row 1 first."""
    out = [(f"{face} {s:g}", ((face, s),)) for s in SIZES]
    out.extend(
        (f"{face} {s:g} beside Calibri {c:g}", ((face, float(s)), ("Calibri", float(c))))
        for s in PAIR_SIZES
        for c in CALIBRI_SIZES
    )
    return out


def _path(face: str) -> Path:
    lower = "_lower" if face.islower() else ""
    return (DIRECTORY / f"fh_{face.replace(' ', '_').lower()}{lower}.xlsx").resolve()


def build(face: str) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils.indexed_list import IndexedList

    wb = Workbook()
    base = Font(name=face, sz=1, family=2, scheme="minor" if face == "Calibri" else None)
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    for r, (_, cells) in enumerate(rows(face), 1):
        for c, (name, size) in enumerate(cells, 1):
            cell = ws.cell(row=r, column=c, value=1234)
            cell.font = Font(name=name, sz=size, family=2)
    DIRECTORY.mkdir(parents=True, exist_ok=True)
    path = _path(face)
    wb.save(path)
    return path


def height(face: str, size: float) -> int | None:
    """The proposed height of a face at a size (None for a face without a slope)."""
    a = SLOPES.get(face)
    return None if a is None else math.floor(a * math.ceil(size) + 1.5)


def predict_row(cells: tuple[tuple[str, float], ...]) -> int | None:
    """The proposed row: the largest ascent (size + 1) plus the largest B = F - S - 1."""
    heights = [height(name, size) for name, size in cells]
    if any(h is None for h in heights):
        return None
    sizes = [math.ceil(size) for _, size in cells]
    return max(sizes) + 1 + max(h - s - 1 for h, s in zip(heights, sizes, strict=True))  # type: ignore[operator]


def predict() -> None:
    PREDICTIONS.parent.mkdir(parents=True, exist_ok=True)
    with PREDICTIONS.open("w", encoding="utf-8") as out:
        out.write(f"# predicted {dt.datetime.now().isoformat(timespec='seconds')} from SLOPES\n")
        out.write("face\ttag\tpredicted\n")
        for face in FACES:
            for tag, cells in rows(face):
                p = predict_row(cells)
                out.write(f"{face}\t{tag}\t{'' if p is None else p}\n")
    print(f"wrote {PREDICTIONS}")


def run(faces: tuple[str, ...] = FACES, output: Path = OUTPUT, suite: str = "font_height") -> None:
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    lines = []
    standards = []
    for face in faces:
        path = build(face)
        n = len(rows(face))
        proc = osascript_or_exit(str(SCRIPT), str(path), str(n), path.name, seconds=180.0)
        if proc.returncode != 0:
            sys.exit(f"{face}: {proc.stderr.strip()[:200]}")
        fields = proc.stdout.strip().split("\x1f")
        if fields[0] != "standard" or len(fields) != n + 2:
            sys.exit(f"{face}: expected {n} heights, got {len(fields) - 2}")
        standards.append(f"{face} {fields[1]}")
        for (tag, _), h in zip(rows(face), fields[2:], strict=True):
            lines.append(f"{face}\t{tag}\t{h}")
        print(f"{face}: {n} rows")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8") as out:
        out.write(f"# suite: {suite} (capture.font_heights); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; one workbook per face, Normal style the "
            "face at 1 point; a number per cell; rows fitted by Excel on opening; standard "
            f"heights: {', '.join(standards)}\n"
        )
        out.write("face\ttag\trow_height\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {output}: {len(lines)} rows")


WRAPPED_OUTPUT = OUTPUT.with_name("wrapped_height.layout.tsv")
WRAPPED_NORMALS = (20, 9)  # Calibri sizes of the Normal style


def wrapped_rows() -> list[tuple[str, str, float, int]]:
    """(tag, face, size, lines) for each row of a wrapped workbook."""
    return [
        (f"{face} {s} wrapped {n}", face, float(s), n)
        for face in FACES
        for s in PAIR_SIZES
        for n in (1, 2, 3)
    ]


def _wrapped_path(normal: int) -> Path:
    return (DIRECTORY / f"fh_wrapped_calibri{normal}.xlsx").resolve()


def build_wrapped(normal: int) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.indexed_list import IndexedList

    wb = Workbook()
    base = Font(name="Calibri", sz=normal, family=2, scheme="minor")
    wb._fonts = IndexedList([base])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = base
    ws = wb.active
    for r, (_, face, size, n) in enumerate(wrapped_rows(), 1):
        cell = ws.cell(row=r, column=1, value="\n".join(["x"] * n))
        cell.data_type = "s"
        cell.font = Font(name=face, sz=size, family=2)
        cell.alignment = Alignment(wrap_text=True)
    DIRECTORY.mkdir(parents=True, exist_ok=True)
    path = _wrapped_path(normal)
    wb.save(path)
    return path


def predict_wrapped(normal: int, face: str, size: float, n: int) -> int | None:
    """Lines of F + 1, plus the Normal font's line inset over the face's (inset = B - 2), or the
    row the two fonts make unwrapped if taller."""
    f, fn = height(face, size), height("Calibri", normal)
    if f is None or fn is None:
        return None
    b, bn = f - math.ceil(size) - 1, fn - normal - 1
    lift = max(0, (bn - 2) - (b - 2))
    unwrapped = max(math.ceil(size), normal) + 1 + max(b, bn)
    return max(unwrapped, n * (f + 1) + lift)


def wrapped() -> None:
    if not answers():
        sys.exit("Excel does not answer: nothing run")
    predictions = BUILD_DIR / "wrapped_heights.predictions.tsv"
    predictions.parent.mkdir(parents=True, exist_ok=True)  # capture/build is not tracked
    with predictions.open("w", encoding="utf-8") as out:
        out.write(f"# predicted {dt.datetime.now().isoformat(timespec='seconds')}\n")
        for normal in WRAPPED_NORMALS:
            for tag, face, size, n in wrapped_rows():
                p = predict_wrapped(normal, face, size, n)
                out.write(f"Calibri {normal}\t{tag}\t{'' if p is None else p}\n")
    lines = []
    for normal in WRAPPED_NORMALS:
        path = build_wrapped(normal)
        n = len(wrapped_rows())
        proc = osascript_or_exit(str(SCRIPT), str(path), str(n), path.name, seconds=180.0)
        if proc.returncode != 0:
            sys.exit(f"Calibri {normal}: {proc.stderr.strip()[:200]}")
        fields = proc.stdout.strip().split("\x1f")
        if fields[0] != "standard" or len(fields) != n + 2:
            sys.exit(f"Calibri {normal}: expected {n} heights, got {len(fields) - 2}")
        for (tag, *_), h in zip(wrapped_rows(), fields[2:], strict=True):
            lines.append(f"Calibri {normal}\t{tag}\t{h}")
    with WRAPPED_OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: wrapped_height (capture.font_heights wrapped); Excel for Mac, en_US\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; Normal style Calibri 20 or 9; one cell "
            "a row in column A, wrap on, 'x' a line; rows fitted by Excel on opening\n"
        )
        out.write("normal_font\ttag\trow_height\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {WRAPPED_OUTPUT}: {len(lines)} rows")


UNDERLINE_OUTPUT = OUTPUT.with_name("underline_height.layout.tsv")
UNDERLINES = ("none", "single", "double", "singleAccounting", "doubleAccounting")


def underline_rows() -> list[tuple[str, str, float, str, int]]:
    """(tag, face, size, underline, lines) for each row of the underline workbook."""
    return [
        (f"{face} {s} {u} {n}", face, float(s), u, n)
        for face in FACES
        for s in range(6, 73)
        for u in UNDERLINES
        for n in ((1, 3) if u in ("none", "singleAccounting") else (1,))
    ]


def predict_underline(face: str, size: float, underline: str, n: int) -> int | None:
    """The prediction written before item 58's run: an accounting underline adds 3 points a line,
    the others nothing (it held to about 30 points; CAPTURING.md section 2 item 58)."""
    f = height(face, size)
    if f is None:
        return None
    extra = 3 if "Accounting" in underline else 0
    return f + extra if n == 1 else n * (f + 1 + extra)


def underlines() -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.indexed_list import IndexedList

    if not answers():
        sys.exit("Excel does not answer: nothing run")
    predictions = BUILD_DIR / "underline_heights.predictions.tsv"
    predictions.parent.mkdir(parents=True, exist_ok=True)  # capture/build is not tracked
    with predictions.open("w", encoding="utf-8") as out:
        out.write(f"# predicted {dt.datetime.now().isoformat(timespec='seconds')}\n")
        for tag, face, size, u, n in underline_rows():
            p = predict_underline(face, size, u, n)
            out.write(f"{tag}\t{'' if p is None else p}\n")
    # one workbook per face: Excel fits only about the first 1,400 rows of a workbook it opens
    # (later rows read 16, the default height), so no workbook may hold more
    lines = []
    for face in FACES:
        rows = [r for r in underline_rows() if r[1] == face]
        wb = Workbook()
        base = Font(name="Calibri", sz=1, family=2, scheme="minor")
        wb._fonts = IndexedList([base])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = base
        ws = wb.active
        for r, (_, _, size, u, n) in enumerate(rows, 1):
            # one character a line: no line wraps again in the narrow column of a 1-point Normal
            cell = ws.cell(row=r, column=1, value="\n".join(["H"] * n))
            cell.data_type = "s"
            cell.font = Font(name=face, sz=size, family=2, u=None if u == "none" else u)
            if n > 1:
                cell.alignment = Alignment(wrap_text=True)
        DIRECTORY.mkdir(parents=True, exist_ok=True)
        path = (DIRECTORY / f"fh_underlines_{face.replace(' ', '_').lower()}.xlsx").resolve()
        wb.save(path)
        proc = osascript_or_exit(str(SCRIPT), str(path), str(len(rows)), path.name, seconds=180.0)
        if proc.returncode != 0:
            sys.exit(f"{face}: {proc.stderr.strip()[:200]}")
        fields = proc.stdout.strip().split("\x1f")
        if fields[0] != "standard" or len(fields) != len(rows) + 2:
            sys.exit(f"{face}: expected {len(rows)} heights, got {len(fields) - 2}")
        lines.extend(f"{r[0]}\t{h}" for r, h in zip(rows, fields[2:], strict=True))
        print(f"{face}: {len(rows)} rows")
    with UNDERLINE_OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: underline_height (capture.font_heights underlines); Excel for Mac\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; Normal style Calibri 1; 'H' a line in "
            "column A, wrap on for three lines; rows fitted by Excel on opening\n"
        )
        out.write("tag\trow_height\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {UNDERLINE_OUTPUT}: {len(lines)} rows")


ACCOUNTING_OUTPUT = OUTPUT.with_name("accounting_height.layout.tsv")
ACCOUNTING_FACES = ("Times New Roman", "Georgia")


def accounting_rows() -> list[tuple[str, str, float, str, int]]:
    """Times New Roman around its four exceptional sizes, Georgia beside it, every half point."""
    return [
        (f"{face} {s / 2:g} {u} {n}", face, s / 2, u, n)
        for face in ACCOUNTING_FACES
        for s in range(72, 83)
        for u in ("none", "singleAccounting", "doubleAccounting")
        for n in (1, 3)
    ]


def accounting() -> None:
    """Capture ``accounting_rows`` in a workbook per face, as ``underlines`` does."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.indexed_list import IndexedList

    if not answers():
        sys.exit("Excel does not answer: nothing run")
    lines = []
    for face in ACCOUNTING_FACES:
        rows = [r for r in accounting_rows() if r[1] == face]
        wb = Workbook()
        base = Font(name="Calibri", sz=1, family=2, scheme="minor")
        wb._fonts = IndexedList([base])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = base
        ws = wb.active
        for r, (_, _, size, u, n) in enumerate(rows, 1):
            cell = ws.cell(row=r, column=1, value="\n".join(["H"] * n))
            cell.data_type = "s"
            cell.font = Font(name=face, sz=size, family=2, u=None if u == "none" else u)
            if n > 1:
                cell.alignment = Alignment(wrap_text=True)
        DIRECTORY.mkdir(parents=True, exist_ok=True)
        path = (DIRECTORY / f"fh_accounting_{face.replace(' ', '_').lower()}.xlsx").resolve()
        wb.save(path)
        proc = osascript_or_exit(str(SCRIPT), str(path), str(len(rows)), path.name, seconds=180.0)
        if proc.returncode != 0:
            sys.exit(f"{face}: {proc.stderr.strip()[:200]}")
        fields = proc.stdout.strip().split("\x1f")
        if fields[0] != "standard" or len(fields) != len(rows) + 2:
            sys.exit(f"{face}: expected {len(rows)} heights, got {len(fields) - 2}")
        lines.extend(f"{r[0]}\t{h}" for r, h in zip(rows, fields[2:], strict=True))
    with ACCOUNTING_OUTPUT.open("w", encoding="utf-8") as out:
        out.write("# suite: accounting_height (capture.font_heights accounting); Excel for Mac\n")
        out.write(f"# excel_build: {excel_build()}\n{edition_line()}\n")
        out.write(
            f"# captured: {dt.date.today().isoformat()}; Normal style Calibri 1; 'H' a line in "
            "column A, wrap on for three lines; rows fitted by Excel on opening\n"
        )
        out.write("tag\trow_height\n")
        out.write("\n".join(lines) + "\n")
    print(f"wrote {ACCOUNTING_OUTPUT}: {len(lines)} rows")


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    command = argv[0] if argv else ""
    if command == "build":
        for face in FACES:
            print(build(face))
    elif command == "predict":
        predict()
    elif command == "run":
        run()
    elif command == "more":
        out = Path(argv[1]) if len(argv) > 1 else MORE_OUTPUT
        run(MORE_FACES, out, "font_height_more")
    elif command == "wrapped":
        wrapped()
    elif command == "underlines":
        underlines()
    elif command == "accounting":
        accounting()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
