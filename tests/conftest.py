import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--tier",
        choices=["fast", "full"],
        default="full",
        help=(
            "Capture tier: 'full' (default, every captured row, about 25 s) or 'fast' "
            "(the per-format sample, about 10 s); each about 12 s more on a first run, "
            "while the written capture jobs are not yet in pytest's cache."
        ),
    )
    parser.addoption(
        "--record-glyph-slack",
        action="store_true",
        help=(
            "Write each suite's count of glyph-slack rows into tests/captures/glyph_slack.json "
            "instead of requiring the counts listed there."
        ),
    )


@pytest.fixture(scope="session")
def record_glyph_slack(request: pytest.FixtureRequest) -> bool:
    return bool(request.config.getoption("--record-glyph-slack"))


@pytest.fixture(scope="session")
def tier(request: pytest.FixtureRequest) -> str:
    tier = request.config.getoption("--tier")
    assert isinstance(tier, str)
    return tier
