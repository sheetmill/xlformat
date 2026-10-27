import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--tier",
        choices=["fast", "full"],
        default="full",
        help=(
            "Capture tier: 'full' (default, every captured row, a few seconds) or 'fast' "
            "(the per-format sample, under a quarter of a second)."
        ),
    )


@pytest.fixture(scope="session")
def tier(request: pytest.FixtureRequest) -> str:
    tier = request.config.getoption("--tier")
    assert isinstance(tier, str)
    return tier
