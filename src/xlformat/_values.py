"""Input normalization: every accepted value becomes a ``Value`` with a kind and a Decimal.

Numbers are ``Decimal`` built from the exact binary value of the input and then limited to
Excel's 15 significant digits with Excel's tie rule (``QUIRKS.md``, ``fifteen_digit_precision``).
Dates and times become Excel serials under the requested date system, including the Lotus
shift of the 1900 system.
"""

from __future__ import annotations

import datetime as _dt
import math
import operator
from dataclasses import dataclass
from decimal import ROUND_HALF_DOWN, Decimal, localcontext
from enum import Enum

from xlformat._quirks import Quirks

_EPOCH_1900 = _dt.date(1899, 12, 30)  # serial 61 = 1900-03-01 under this epoch
_EPOCH_1900_PRE_LEAP = _dt.date(1899, 12, 31)  # serials 1..59 (Jan 1 .. Feb 28, 1900)
_LEAP_DAY = _dt.date(1900, 3, 1)
_EPOCH_1904 = _dt.date(1904, 1, 1)
_SECONDS_PER_DAY = 86400
_TWO_53 = 2**53


class Kind(Enum):
    """What sort of thing a value is, for section selection and alignment."""

    NUMBER = "number"
    TEXT = "text"
    BOOL = "bool"
    EMPTY = "empty"


@dataclass(frozen=True, slots=True)
class Value:
    """A normalized input.

    ``number`` (15-digit limited) is set for NUMBER, ``text`` for TEXT and BOOL. ``raw`` is the
    value as a double, which is what Excel's clock arithmetic works from (``_datetime.py``).
    """

    kind: Kind
    number: Decimal | None = None
    text: str | None = None
    raw: float = 0.0


def limit_precision(x: Decimal) -> Decimal:
    """Excel's 15-significant-digit rule on a non-negative exact value: ties round down.

    Evidence: ``QUIRKS.md``. An earlier extra step for values at or above 1e14 was an artifact
    of the captures' 16-digit number texts, not of Excel (``CAPTURING.md`` section 2 item 16).
    """
    if x == 0:
        return x
    with localcontext() as ctx:
        ctx.prec = 60
        return x.quantize(Decimal(1).scaleb(x.adjusted() - 14), rounding=ROUND_HALF_DOWN)


def _from_float(f: float) -> Decimal:
    if math.isnan(f) or math.isinf(f):
        msg = f"cannot format non-finite number {f!r}"
        raise ValueError(msg)
    return Decimal(f)


def _from_int(i: int) -> Decimal:
    if -_TWO_53 <= i <= _TWO_53:
        return Decimal(i)
    return _from_float(float(i))  # Excel holds a double; mimic its rounding of big integers


def _serial_of_date(d: _dt.date, date1904: bool, quirks: Quirks) -> int:
    if date1904:
        return (d - _EPOCH_1904).days
    if quirks.leap_year_1900 and d < _LEAP_DAY:
        return (d - _EPOCH_1900_PRE_LEAP).days
    return (d - _EPOCH_1900).days


def _fraction_of_day(t: _dt.time) -> Decimal:
    seconds = t.hour * 3600 + t.minute * 60 + t.second
    return (Decimal(seconds) + Decimal(t.microsecond).scaleb(-6)) / _SECONDS_PER_DAY


def to_serial(
    value: _dt.date | _dt.datetime | _dt.time | _dt.timedelta,
    *,
    date1904: bool = False,
    quirks: Quirks | None = None,
) -> Decimal:
    """Excel serial number for a date, datetime, time or timedelta (exact Decimal)."""
    quirks = quirks or Quirks.EXCEL
    with localcontext() as ctx:
        ctx.prec = 40
        if isinstance(value, _dt.datetime):
            days = _serial_of_date(value.date(), date1904, quirks)
            return Decimal(days) + _fraction_of_day(value.time())
        if isinstance(value, _dt.date):
            return Decimal(_serial_of_date(value, date1904, quirks))
        if isinstance(value, _dt.time):
            return _fraction_of_day(value)
        total = (
            Decimal(value.days)
            + (Decimal(value.seconds) + Decimal(value.microseconds).scaleb(-6)) / _SECONDS_PER_DAY
        )
        return total


def normalize(value: object, *, date1904: bool = False, quirks: Quirks | None = None) -> Value:
    """Turn any accepted input into a ``Value``.

    Raises:
        TypeError: for an unsupported type.
        ValueError: for NaN or an infinity.
    """
    quirks = quirks or Quirks.EXCEL
    if value is None:
        return Value(Kind.EMPTY)
    if isinstance(value, bool):
        return Value(Kind.BOOL, text="TRUE" if value else "FALSE")
    if isinstance(value, str):
        return Value(Kind.TEXT, text=value)
    number: Decimal
    if isinstance(value, int):
        number = _from_int(value)
    elif isinstance(value, float):
        number = _from_float(value)
    elif isinstance(value, Decimal):
        if not value.is_finite():
            msg = f"cannot format non-finite number {value!r}"
            raise ValueError(msg)
        number = value
    elif isinstance(value, _dt.date | _dt.time | _dt.timedelta):
        number = to_serial(value, date1904=date1904, quirks=quirks)
    elif hasattr(value, "__index__"):
        number = _from_int(operator.index(value))
    elif hasattr(value, "__float__"):
        number = _from_float(float(value))
    else:
        msg = f"unsupported value type {type(value).__name__}"
        raise TypeError(msg)
    raw = float(number)
    if quirks.fifteen_digit_precision:
        magnitude = limit_precision(abs(number))
        number = -magnitude if number < 0 else magnitude
    return Value(Kind.NUMBER, number=number, raw=raw)
