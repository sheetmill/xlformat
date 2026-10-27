import xlformat


def test_version_is_a_string() -> None:
    assert isinstance(xlformat.__version__, str)
