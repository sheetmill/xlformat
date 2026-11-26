"""Date/time rendering: calendar arithmetic, the leap quirk, 1904, elapsed codes."""

import datetime as dt

import pytest

from xlformat import Quirks, format_value
from xlformat._datetime import calendar, civil_from_days, days_from_civil


@pytest.mark.parametrize(
    "date", [dt.date(1, 1, 1), dt.date(1899, 12, 30), dt.date(2000, 2, 29), dt.date(9999, 12, 31)]
)
def test_civil_round_trip(date: dt.date) -> None:
    n = days_from_civil(date.year, date.month, date.day)
    assert civil_from_days(n) == (date.year, date.month, date.day)


@pytest.mark.parametrize(
    ("serial", "expected"),
    [
        (0, (1900, 1, 0, 6)),  # January 0, a Saturday
        (1, (1900, 1, 1, 0)),  # Sunday in Excel, Monday in reality
        (59, (1900, 2, 28, 2)),
        (60, (1900, 2, 29, 3)),  # the day that did not exist
        (61, (1900, 3, 1, 4)),  # Thursday, correct from here on
        (44927, (2023, 1, 1, 0)),
        (2958465, (9999, 12, 31, 5)),
    ],
)
def test_calendar_1900(serial: int, expected: tuple[int, int, int, int]) -> None:
    assert calendar(serial, False, Quirks.EXCEL) == expected


def test_calendar_strict_has_no_leap_day() -> None:
    assert calendar(60, False, Quirks.STRICT) == (1900, 2, 28, 3)
    assert calendar(61, False, Quirks.STRICT) == (1900, 3, 1, 4)
    assert calendar(1, False, Quirks.STRICT) == (1899, 12, 31, 0)


def test_calendar_1904() -> None:
    assert calendar(0, True, Quirks.EXCEL) == (1904, 1, 1, 5)  # a Friday
    assert calendar(44927 - 1462, True, Quirks.EXCEL) == (2023, 1, 1, 0)


def test_rendering_samples() -> None:
    assert format_value(60, "dddd, mmmm d, yyyy", platform="mac") == "Wednesday, February 29, 1900"
    assert format_value(0.5, "h:mm AM/PM", platform="mac") == "12:00 PM"
    assert format_value(0.5, "h a/p", platform="mac") == "12 p"
    assert format_value(1.5, "[h]:mm", platform="mac") == "36:00"
    assert format_value(-1, "[h]", platform="mac") == "-24"  # elapsed-only sections take negatives
    assert format_value(-1, "h:mm", platform="mac") == "#" * 255  # clock sections do not, in 1900
    assert format_value(-1, "m/d/yyyy", date1904=True, platform="mac") == "-1/2/1904"
    assert (
        format_value(44927.9999999884, "yyyy-mm-dd hh:mm:ss.00", platform="mac")
        == "2023-01-02 00:00:00.00"
    )
    assert (
        format_value(44927.9999999884, "yyyy-mm-dd hh:mm:ss.000", platform="mac")
        == "2023-01-01 23:59:59.999"
    )
    assert format_value(1234, "e+", platform="mac") == "1903+"
    assert format_value(0, "[$-F800]", platform="mac") == "Saturday, January 0, 1900"
    assert format_value(2958466, "m/d/yy", platform="mac") == "#" * 255
