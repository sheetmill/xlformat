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
    superseded = re.compile(r"earlier version|An earlier|were removed on")
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
