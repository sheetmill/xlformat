"""Guards on the published documents: figures that match the data, current rules, live labels."""

from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path

from capture.schema import read_rows

ROOT = Path(__file__).resolve().parents[1]
CAPTURES = ROOT / "tests" / "captures"
MARKDOWN = [
    ROOT / name
    for name in (
        "README.md",
        "ARCHITECTURE.md",
        "CAPTURING.md",
        "QUIRKS.md",
        "CHANGELOG.md",
        "CONTRIBUTING.md",
    )
] + sorted((ROOT / "docs").glob("*.md"))
CURRENT_ONLY = [
    ROOT / "QUIRKS.md",
    ROOT / "README.md",
    *sorted((ROOT / "docs").glob("*.md")),
    *sorted((ROOT / "src" / "xlformat").glob("*.py")),
]


@lru_cache(maxsize=1)
def _suite_rows() -> dict[str, int]:
    return {
        path.name: len(read_rows(path)[1])
        for path in CAPTURES.iterdir()
        if path.name.endswith((".tsv", ".tsv.gz"))
        and ".measure." not in path.name
        and ".layout." not in path.name
    }


def test_published_suite_totals_match_the_captures() -> None:
    rows = _suite_rows()
    with_rows = sum(1 for count in rows.values() if count)
    total = sum(rows.values())
    page = " ".join((ROOT / "docs" / "testing.md").read_text().split())
    match = re.search(r"In all, (\d+) suite files \((\d+) with rows.*?hold ([\d,]+) rows", page)
    assert match is not None
    assert (int(match[1]), int(match[2])) == (len(rows), with_rows)
    assert int(match[3].replace(",", "")) == total
    architecture = (ROOT / "ARCHITECTURE.md").read_text()
    assert f"`pytest`, {len(rows)} suite files, {total:,} rows" in architecture


def test_the_testing_page_table_accounts_for_every_row() -> None:
    table = 0
    for line in (ROOT / "docs" / "testing.md").read_text().splitlines():
        cells = [cell.strip() for cell in line.split("|")]
        if len(cells) == 5 and re.fullmatch(r"\d[\d,]*(, \d[\d,]*)*", cells[2]):
            table += sum(int(n.replace(",", "")) for n in re.findall(r"\d+(?:,\d{3})*", cells[2]))
    assert table == sum(_suite_rows().values())


def test_user_facing_pages_state_the_current_rule() -> None:
    # superseded values and dated remarks belong in CAPTURING.md and the changelog
    superseded = re.compile(
        r"earlier version|An earlier|were removed on|\(was |\(20\d\d-\d\d-\d\d\)"
    )
    hits = [
        f"{path.relative_to(ROOT)}: {match[0]}"
        for path in CURRENT_ONLY
        for match in superseded.finditer(path.read_text())
    ]
    assert not hits


def test_no_labels_without_a_target() -> None:
    label = re.compile(r"\(point [a-z]\)")
    hits = [str(path.relative_to(ROOT)) for path in MARKDOWN if label.search(path.read_text())]
    assert not hits


def test_the_module_layout_lists_every_module() -> None:
    # ARCHITECTURE.md section 12 names each module of src/xlformat with a line of its own
    text = (ROOT / "ARCHITECTURE.md").read_text(encoding="utf-8")
    section = text[text.index("## 12. Module layout") : text.index("## 13.")]
    listed = set(re.findall(r"^  (_?\w+\.py) ", section, flags=re.MULTILINE))
    modules = {p.name for p in (ROOT / "src" / "xlformat").glob("*.py")}
    assert listed == modules, (sorted(modules - listed), sorted(listed - modules))


def test_the_skipped_color_count_is_the_list_length() -> None:
    from tests.test_captures import DISPLAY_FORMAT_COLOR_MISREPORTS

    words = {"fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19}
    want = len(DISPLAY_FORMAT_COLOR_MISREPORTS)
    pages = ("README.md", "QUIRKS.md", "docs/testing.md", "CAPTURING.md")
    found = 0
    for name in pages:
        text = " ".join((ROOT / name).read_text(encoding="utf-8").split())
        for word in re.findall(r"\b(\w+) color (?:readings|rows)\b", text, flags=re.IGNORECASE):
            assert words.get(word.lower()) == want, (name, word, want)
            found += 1
    assert found == len(pages)


def test_the_color_cases_not_seen_on_screen_are_named_alike() -> None:
    from tests.test_captures import DISPLAY_FORMAT_COLOR_MISREPORTS, NOT_SEEN_ON_SCREEN

    assert set(DISPLAY_FORMAT_COLOR_MISREPORTS) >= NOT_SEEN_ON_SCREEN
    assert len(NOT_SEEN_ON_SCREEN) == 2
    for name in ("QUIRKS.md", "docs/testing.md", "CAPTURING.md"):
        text = " ".join((ROOT / name).read_text(encoding="utf-8").split())
        assert "`[Red]General` and text under a lone `[Red]`" in text, name
        unseen = r"Two cases have not been seen on screen|The screen has shown neither"
        assert re.search(unseen, text), name
