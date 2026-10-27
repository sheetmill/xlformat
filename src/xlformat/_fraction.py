"""Fraction sections (``# ?/?``, ``# ??/16``). See ``ARCHITECTURE.md`` section 6 and ``QUIRKS.md``.

Excel's denominator search for free denominators is a continued-fraction expansion stopped at
an epsilon, run on the double, not the closest fraction; ``fraction_denominators`` off uses
``Fraction.limit_denominator``. The fixed-denominator numerator is Excel's nudged half-up in
double arithmetic (``fixed_numerator``), zero misses on every captured row.
"""

from __future__ import annotations

import math
from decimal import ROUND_HALF_DOWN, ROUND_HALF_UP, Decimal
from fractions import Fraction

from xlformat._number import render_number
from xlformat._parser import FractionSpec, NumberSpec, Piece
from xlformat._quirks import Quirks

_NUDGE = 1.0 + 2.0**-52  # 1 + DBL_EPSILON, exactly representable


def excel_fraction(value: float, max_denominator: int) -> Fraction:
    """Excel's continued-fraction search.

    Adopted from ExcelNumberFormat; confirmed by the 2023 capture for one- to three-digit
    denominators.
    """
    b = value
    p2, p1, q2, q1 = 0.0, 1.0, 1.0, 0.0
    p = q = 0.0
    while q1 < max_denominator:
        a = math.floor(b)
        p = a * p1 + p2
        q = a * q1 + q2
        if (b - a) < 0.00000005:
            break
        b = 1 / (b - a)
        p2, p1, q2, q1 = p1, p, q1, q
    if q > max_denominator:
        if q1 > max_denominator:
            p, q = p2, q2
        else:
            p, q = p1, q1
    return Fraction(int(p), int(q))


def fixed_numerator(value: Decimal, denominator: int) -> int:
    """Numerator for a fixed denominator, as Excel computes it.

    ``value`` is the non-negative value (a double, possibly scaled by percent signs, as a
    Decimal); the result is the numerator of the whole value, ``W * D + n``.

    The rule (``CAPTURING.md`` section 2 item 22; found with a larger capture of half-step values,
    included as ``fraction_threshold`` and ``fraction_random``; every
    fixed-denominator captured row, 38,904 of them across ``fraction``, ``fraction_ties``,
    ``fraction_threshold``, ``fraction_random``, ``capture2023`` and ``capture2026``, zero
    misses): the fraction part of the double is scaled up by one relative epsilon,
    ``f' = fl(f * (1 + 2**-52))``, the product ``t = fl(f' * D)`` is formed in double
    arithmetic, and the numerator is ``floor(t)`` plus one when ``t - floor(t) >= 0.5``.
    Phrasing the half-up as ``floor(t + 0.5)`` is wrong at k = 0, where ``t + 0.5`` rounds up to
    1 for ``t`` within a quarter ulp of 0.5.

    The constant is DBL_EPSILON exactly: the window is (0.995, 1.005) * 2**-52, bounded below
    by 0.7504999999999997 under ``# ???/1000`` (Excel rounds up) and above by
    0.09374999999999997 under ``# ??/16`` (Excel rounds down); 2**-53 misses 172 rows and
    2**-51 misses 377. The nudge is applied to the fraction part, not the whole value (rows with
    whole parts 1 and 12 decide like their whole-part-0 twins; nudging ``x`` misses 105), and
    before the product, not after (misses 52); dividing by ``1 / D`` misses 29.

    Why it looked like a tolerance: for a fraction part with mantissa ``m`` in [1, 2), the nudge
    adds ``round(m)`` ulps, one below ``m = 1.5`` and two above, so a deficit of two ulps below
    the half step rounds up only in the upper half of each binade; no tolerance relative to the
    half step, the value or the denominator fitted every row.
    """
    x = float(value)
    whole = math.floor(x)
    fraction = x - whole  # exact: the fractional bits of x
    scaled = fraction * _NUDGE * denominator  # two roundings, left to right, as Excel's C code
    n = math.floor(scaled)
    if scaled - n >= 0.5:
        n += 1
    return whole * denominator + n


def fraction_value(raw: float) -> Decimal:
    """The value fractions work from: the exact double, as a Decimal.

    An earlier version rounded it to 15 significant digits; the recapture with exact number
    texts showed the step changes nothing (CAPTURING.md section 2 item 18).
    """
    return Decimal(abs(raw))


_MAX_NUMERATOR = 2**31 - 1  # a larger numerator shows '#' (fraction suite: 49382715605/4)


def render_fraction(
    spec: FractionSpec, raw: float, negative: bool, quirks: Quirks
) -> list[Piece] | None:
    """Pieces for a fraction section, or ``None`` when the numerator overflows 32 bits.

    ``raw`` is the double (sign ignored).
    """
    v = fraction_value(raw)
    x = float(v)
    for _ in range(spec.percent):
        x *= 100.0  # in double arithmetic, as Excel does (1234567.891 under ?/?% is 1111111102/9)
    if spec.fixed_denominator and quirks.fraction_denominators:
        frac = Fraction(fixed_numerator(Decimal(x), spec.fixed_denominator), spec.fixed_denominator)
    elif spec.fixed_denominator:
        # the closest fraction with this denominator: plain half-up on the exact product
        n = int((Decimal(x) * spec.fixed_denominator).quantize(Decimal(1), rounding=ROUND_HALF_UP))
        frac = Fraction(n, spec.fixed_denominator)
    elif quirks.fraction_denominators:
        frac = excel_fraction(x, 10 ** len(spec.denominator) - 1)
    else:
        frac = Fraction(x).limit_denominator(10 ** len(spec.denominator) - 1)
    whole = 0
    if spec.whole is not None:
        whole = math.floor(frac)
        frac -= whole
        if x > 99_999_999_999_999:  # 15+ integer digits: Excel rounds half down here
            whole = int(Decimal(x).quantize(Decimal(1), rounding=ROUND_HALF_DOWN))
    numerator, denominator = frac.numerator, frac.denominator
    if spec.fixed_denominator:
        numerator = numerator * (spec.fixed_denominator // denominator)
        denominator = spec.fixed_denominator
    if numerator > _MAX_NUMERATOR:
        return None
    show_fraction = spec.whole is None or "0" in spec.numerator
    out: list[Piece] = []
    if negative and not (spec.whole is None and numerator == 0):
        out.append(("lit", "-"))  # a bare zero fraction drops the sign ('0/1' for -0.0001)
    out.extend(spec.leading)
    if spec.whole is not None:
        whole_digits = [p[1] for p in spec.whole.integer if p[0] == "digit"]
        if whole == 0 and "0" not in whole_digits and "0" in spec.numerator:
            pass  # '# 0/0' with 0.125 shows '1/8', with 0 shows '0/1': whole part and space vanish
        elif whole == 0 and frac == 0:
            # the whole value is zero: Excel shows a 0 even for '#' placeholders ('0    ')
            out.extend(_forced_zero(spec.whole))
            out.extend(spec.whole_trailing)
        else:
            out.extend(render_number(spec.whole, Decimal(whole), False))
            out.extend(spec.whole_trailing)
    if frac or show_fraction:
        n_text = str(numerator)
        if "0" in spec.numerator:
            n_text = n_text.rjust(len(spec.numerator), "0")
        elif "?" in spec.numerator:
            n_text = n_text.rjust(len(spec.numerator), " ")
        d_text = str(denominator)
        if not spec.fixed_denominator:
            if "0" in spec.denominator:
                d_text = d_text.rjust(len(spec.denominator), "0")
            elif "?" in spec.denominator:
                d_text = d_text.ljust(len(spec.denominator), " ")
        out.append(("lit", n_text))
        out.extend(spec.slash)
        out.append(("lit", d_text))
    elif "?" in spec.numerator:
        out.append(("lit", " " * len(spec.numerator)))
        out.extend(
            ("lit", " " * len(text)) if kind == "lit" else (kind, text) for kind, text in spec.slash
        )
        out.append(("lit", " " * len(spec.denominator)))
    out.extend(spec.trailing)
    return out


def _forced_zero(whole: NumberSpec) -> list[Piece]:
    """Render a zero whole part with its last placeholder forced to show a digit."""
    last = max(i for i, p in enumerate(whole.integer) if p[0] == "digit")
    forced = NumberSpec(
        leading=whole.leading,
        integer=tuple(("digit", "0") if i == last else p for i, p in enumerate(whole.integer)),
        decimal=whole.decimal,
        trailing=whole.trailing,
        has_point=whole.has_point,
        grouping=whole.grouping,
        scale=whole.scale,
        percent=whole.percent,
        n_integer=whole.n_integer,
        n_decimal=whole.n_decimal,
    )
    return render_number(forced, Decimal(0), False)
