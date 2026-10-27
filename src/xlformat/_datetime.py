"""Date and time rendering. See ``ARCHITECTURE.md`` section 7 and ``QUIRKS.md``.

The serial is taken as a double (Excel's own representation, not the 15-digit display value).
Excel's clock arithmetic, reproduced here operation for operation (``CAPTURING.md`` section 2
item 15; 41,266 captured rows, no misses): the serial is biased by half of the finest displayed
unit (a second, or a tenth, hundredth or thousandth of one; never coarser), then day, hour,
minute and second are peeled from successive fractions, each multiplied up and truncated. The
carry therefore propagates into the day, and the floating-point noise of the peel is Excel's
own: a ``General`` token inside a date section leaks the biased seconds before truncation
(``0.499999803`` for serial 44927). Elapsed codes come from the biased total seconds. Civil
dates come from integer arithmetic over the day number, never from ``datetime`` objects, so
January 0 and February 29, 1900 need no special objects.

Sections made only of elapsed codes (``[h]``, ``[mm]``, ``[s]``) have no upper limit and show
negative values with a leading minus; other sections show ``#`` for negatives in the 1900
system, while the 1904 system shows a minus sign followed by the rendering of the magnitude.
"""

from __future__ import annotations

import math
from decimal import Decimal

from xlformat._number import general_text
from xlformat._parser import NumberFormat, Piece, Section, Token, TokenKind
from xlformat._quirks import Quirks
from xlformat._values import limit_precision

_MONTHS = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
_DAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
_SECONDS_PER_DAY = 86400
_DAY_RECIPROCAL = 1 / 86400  # Excel forms the total seconds as x / (1/86400), not x * 86400
_MAX_DAY_1900 = 2958465  # 9999-12-31
_DAYS_1900_TO_1904 = 1462
SYSTEM_DATE = "dddd, mmmm d, yyyy"
SYSTEM_TIME = "h:mm:ss AM/PM"
_SYSTEM_DATE_TAGS = {"x-sysdate", "f800"}
_SYSTEM_TIME_TAGS = {"x-systime", "f400"}


def days_from_civil(y: int, m: int, d: int) -> int:
    """Day count from 0000-03-01 for a proleptic Gregorian date (inverse of civil_from_days)."""
    y -= m <= 2
    era = (y if y >= 0 else y - 399) // 400
    yoe = y - era * 400
    doy = (153 * (m + (-3 if m > 2 else 9)) + 2) // 5 + d - 1
    doe = yoe * 365 + yoe // 4 - yoe // 100 + doy
    return era * 146097 + doe


def civil_from_days(z: int) -> tuple[int, int, int]:
    """Proleptic Gregorian (year, month, day) for a day count from 0000-03-01."""
    era = (z if z >= 0 else z - 146096) // 146097
    doe = z - era * 146097
    yoe = (doe - doe // 1460 + doe // 36524 - doe // 146096) // 365
    y = yoe + era * 400
    doy = doe - (365 * yoe + yoe // 4 - yoe // 100)
    mp = (5 * doy + 2) // 153
    d = doy - (153 * mp + 2) // 5 + 1
    m = mp + 3 if mp < 10 else mp - 9
    return (y + 1 if m <= 2 else y, m, d)


_EPOCH_1900 = days_from_civil(1899, 12, 30)  # serial 61 = 1900-03-01 under this epoch
_EPOCH_1904 = days_from_civil(1904, 1, 1)


def calendar(day: int, date1904: bool, quirks: Quirks) -> tuple[int, int, int, int]:
    """(year, month, day, weekday with Sunday = 0) for an Excel day number."""
    if date1904:
        y, m, d = civil_from_days(_EPOCH_1904 + day)
        return y, m, d, (day + 5) % 7
    weekday = (day + 6) % 7
    if quirks.leap_year_1900:
        if day == 0:
            return 1900, 1, 0, weekday
        if day == 60:
            return 1900, 2, 29, weekday
        if day < 60:
            y, m, d = civil_from_days(_EPOCH_1900 + 1 + day)
            return y, m, d, weekday
    y, m, d = civil_from_days(_EPOCH_1900 + day)
    return y, m, d, weekday


def _fraction_digits(tokens: tuple[Token, ...]) -> int:
    """Fractional-second digits displayed (0 to 3). Excel always rounds to at least the second."""
    digits = 0
    for t in tokens:
        if t.kind is TokenKind.DATE and "." in t.text:
            digits = max(digits, len(t.text.partition(".")[2]))
    return digits


def _system_tokens(section: Section) -> tuple[Token, ...]:
    """Tokens of the system date or time format when the section carries that tag."""
    tag = (section.locale or "").lower()
    if tag in _SYSTEM_DATE_TAGS:
        return NumberFormat.parse(SYSTEM_DATE).sections[0].tokens
    if tag in _SYSTEM_TIME_TAGS:
        return NumberFormat.parse(SYSTEM_TIME).sections[0].tokens
    return section.tokens


def render_datetime(
    section: Section,
    serial: float,
    display: Decimal,
    *,
    date1904: bool,
    quirks: Quirks,
) -> list[Piece] | None:
    """Pieces for a date/time section, or ``None`` when the value cannot be shown.

    ``serial`` is the double Excel works from; ``display`` is the 15-digit value, which Excel
    does not use for clock arithmetic.
    """
    del display
    tokens = _system_tokens(section)
    elapsed_only = all(
        t.kind not in (TokenKind.DATE, TokenKind.MINUTE, TokenKind.AMPM) for t in tokens
    ) and any(t.kind is TokenKind.ELAPSED for t in tokens)
    negative = serial < 0
    if negative and not (date1904 or elapsed_only):
        return None  # decided on the raw sign: -1e-9 shows '#', not January 0, 1900
    x = abs(serial)
    digits = _fraction_digits(tokens)
    scale = 10**digits
    bias = 0.5 / scale  # half of the displayed unit, in seconds
    # Excel's arithmetic (module docstring): bias the serial, then peel and truncate. Every
    # operation and its order matter: the 59.5-second rows separate x / (1/86400) from
    # x * 86400, and the day-44927 rows separate the bias in days from a bias in seconds.
    d = x + bias / _SECONDS_PER_DAY
    day = math.floor(d)
    h = (d - day) * 24
    hour = math.floor(h)
    m = (h - hour) * 60
    minute = math.floor(m)
    biased_seconds = (m - minute) * 60
    second = math.floor(biased_seconds)
    frac = math.floor((biased_seconds - second) * scale)
    elapsed = x / _DAY_RECIPROCAL + bias
    limit = _MAX_DAY_1900 - (_DAYS_1900_TO_1904 if date1904 else 0)
    if day > limit and not elapsed_only:
        return None
    has_ampm = any(t.kind is TokenKind.AMPM for t in tokens)
    year = month = dom = weekday = 0
    if any(t.kind is TokenKind.DATE and t.text[0] in "ymde" for t in tokens):
        year, month, dom, weekday = calendar(day, date1904, quirks)
    out: list[Piece] = [("lit", "-")] if negative else []
    for t in tokens:
        k = t.kind
        if k is TokenKind.DATE:
            out.append(
                (
                    "lit",
                    _date_field(
                        t.text, year, month, dom, weekday, hour, second, frac, digits, has_ampm
                    ),
                )
            )
        elif k is TokenKind.MINUTE:
            out.append(("lit", f"{minute:02d}" if t.text == "mm" else str(minute)))
        elif k is TokenKind.AMPM:
            out.append(("lit", _ampm(t.text, hour)))
        elif k is TokenKind.ELAPSED:
            whole = math.floor(elapsed / {"h": 3600, "m": 60, "s": 1}[t.text[0]])
            out.append(("lit", str(whole).zfill(len(t.text))))
        elif k is TokenKind.PAD:
            out.append(("pad", t.text))
        elif k is TokenKind.FILL:
            out.append(("fill", t.text))
        elif k is TokenKind.CURRENCY:
            if t.text:
                out.append(("lit", t.text))
        elif k in (TokenKind.COLOR, TokenKind.CONDITION):
            continue
        elif k is TokenKind.GENERAL:
            # Excel leaks its biased seconds before truncation, noise included (a whole day
            # shows "0.5"; serial 44927 shows "0.499999803").
            out.append(("lit", general_text(limit_precision(Decimal(biased_seconds)))))
        else:
            out.append(("lit", t.text))
    return out


def _ampm(token: str, hour: int) -> str:
    pm = hour >= 12
    if len(token) == 3:  # A/P keeps the token's case; AM/PM is always upper-case (datetime suite)
        text = "P" if pm else "A"
        return text.lower() if token.islower() else text
    return "PM" if pm else "AM"


def _date_field(
    code: str,
    year: int,
    month: int,
    dom: int,
    weekday: int,
    hour: int,
    second: int,
    frac: int,
    digits: int,
    has_ampm: bool,
) -> str:
    c = code[0]
    if c == "y":
        return f"{year % 100:02d}" if code == "yy" else f"{year:04d}"
    if c == "e":
        return f"{year:04d}"
    if c == "m":
        if code == "m":
            return str(month)
        if code == "mm":
            return f"{month:02d}"
        name = _MONTHS[month - 1]
        if code == "mmm":
            return name[:3]
        if code == "mmmm":
            return name
        return name[0]
    if c == "d":
        if code == "d":
            return str(dom)
        if code == "dd":
            return f"{dom:02d}"
        name = _DAYS[weekday]
        return name[:3] if code == "ddd" else name
    if c == "h":
        h = hour
        if has_ampm:
            h = hour % 12 or 12
        return f"{h:02d}" if code == "hh" else str(h)
    # seconds, possibly with a fraction
    whole, _, fraction = code.partition(".")
    text = f"{second:02d}" if whole == "ss" else str(second)
    if fraction:
        text += "." + str(frac).zfill(digits)[: len(fraction)]
    return text
