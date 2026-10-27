"""The examples in the README and the documentation are doctests.

Every ``>>>`` line must produce the text shown, so the published examples cannot drift from
the package.
"""

import doctest
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PAGES = [ROOT / "README.md", *sorted((ROOT / "docs").glob("*.md"))]


@pytest.mark.parametrize("page", PAGES, ids=lambda p: p.relative_to(ROOT).as_posix())
def test_examples(page: Path) -> None:
    # Drop the code fences so a closing ``` is not read as expected output.
    lines = page.read_text().splitlines()
    text = "\n".join(line for line in lines if not line.startswith("```"))
    test = doctest.DocTestParser().get_doctest(text, {}, page.name, str(page), 0)
    runner = doctest.DocTestRunner(optionflags=doctest.ELLIPSIS)
    runner.run(test)
    assert runner.failures == 0
