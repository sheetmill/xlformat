"""Value normalization: kinds, 15-digit rule, date serials."""

import datetime as dt
from decimal import Decimal

import pytest

from xlformat import Quirks
from xlformat._values import Kind, limit_precision, normalize, to_serial


def test_kinds() -> None:
    assert normalize(None).kind is Kind.EMPTY
    assert normalize(True) == normalize(True)
    assert (normalize(True).kind, normalize(True).text) == (Kind.BOOL, "TRUE")
    assert (normalize(False).kind, normalize(False).text) == (Kind.BOOL, "FALSE")
    assert (normalize("x").kind, normalize("x").text) == (Kind.TEXT, "x")
    assert normalize(3).kind is Kind.NUMBER


def test_bool_is_checked_before_int() -> None:
    assert normalize(True).kind is Kind.BOOL


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (0.1234512345123456, "0.123451234512346"),
        (12345123451234.25, "12345123451234.2"),
        (12345123451234.55, "12345123451234.6"),
        (123451234512345.5, "123451234512345"),
        (123451234512345.52, "123451234512346"),  # exact .515625: above the tie
        (123451234512345.56, "123451234512346"),
        (1234512345123455.5, "1234512345123460"),
        (1234512345123454.5, "1234512345123450"),
        (999999999999999.5, "999999999999999"),
        (9.915, "9.91500000000000"),
    ],
)
def test_fifteen_digit_rule(value: float, expected: str) -> None:
    assert f"{limit_precision(Decimal(value)):f}" == expected


def test_strict_keeps_the_exact_value() -> None:
    v = normalize(0.1, quirks=Quirks.STRICT)
    assert v.number == Decimal(0.1)  # noqa: RUF032  (the exact binary value is the point)
    assert normalize(0.1).number == Decimal("0.100000000000000")


@pytest.mark.parametrize(
    ("value", "serial"),
    [
        (dt.date(1900, 1, 1), 1),
        (dt.date(1900, 2, 28), 59),
        (dt.date(1900, 3, 1), 61),
        (dt.date(2023, 1, 1), 44927),
        (dt.date(9999, 12, 31), 2958465),
    ],
)
def test_date_serials_1900(value: dt.date, serial: int) -> None:
    assert to_serial(value) == serial


def test_date_serials_strict_has_no_lotus_shift() -> None:
    assert to_serial(dt.date(1900, 1, 1), quirks=Quirks.STRICT) == 2
    assert to_serial(dt.date(1900, 3, 1), quirks=Quirks.STRICT) == 61


def test_date_serials_1904() -> None:
    assert to_serial(dt.date(1904, 1, 1), date1904=True) == 0
    assert to_serial(dt.date(2023, 1, 1), date1904=True) == 44927 - 1462


def test_datetime_time_timedelta() -> None:
    assert to_serial(dt.datetime(2023, 1, 1, 12)) == Decimal("44927.5")
    assert to_serial(dt.time(6)) == Decimal("0.25")
    assert to_serial(dt.timedelta(days=1, hours=12)) == Decimal("1.5")


def test_big_int_goes_through_double() -> None:
    assert normalize(2**63).number == Decimal(float(2**63)).quantize(Decimal("1e4"))


def test_errors() -> None:
    with pytest.raises(ValueError, match="non-finite"):
        normalize(float("nan"))
    with pytest.raises(TypeError):
        normalize(object())
