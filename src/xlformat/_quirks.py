"""Switchable Excel behaviors. See ``QUIRKS.md`` for the evidence behind each one."""

from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar


@dataclass(frozen=True, slots=True)
class Quirks:
    """Which Excel quirks to reproduce. All on by default (the ``EXCEL`` preset).

    ``Quirks.EXCEL`` has every quirk on and is what ``quirks=None`` means; ``Quirks.STRICT``
    has every quirk off. Individual fields can be set: ``Quirks(leap_year_1900=False)``.

    Attributes:
        leap_year_1900: Serial 60 is February 29, 1900 and weekdays before March 1900 are
            shifted, as in Excel's 1900 date system.
        fifteen_digit_precision: Numbers are displayed from at most 15 significant digits:
            the exact double is rounded to 15 digits with ties down before the format's own
            rounding.
        fraction_denominators: Fraction formats choose the fraction the way Excel does: a
            continued-fraction search with an epsilon for ``# ?/?``, and a half-up nudged by
            one relative epsilon for fixed denominators such as ``# ??/16``. Off, both give
            the closest fraction.
    """

    leap_year_1900: bool = True
    fifteen_digit_precision: bool = True
    fraction_denominators: bool = True

    EXCEL: ClassVar[Quirks]
    STRICT: ClassVar[Quirks]


Quirks.EXCEL = Quirks()
Quirks.STRICT = Quirks(
    leap_year_1900=False, fifteen_digit_precision=False, fraction_denominators=False
)
