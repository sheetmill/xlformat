"""Rendering of number sections and of ``General``. See ``ARCHITECTURE.md`` section 6."""

from __future__ import annotations

from collections.abc import Callable
from decimal import ROUND_HALF_UP, Decimal, localcontext

from xlformat._parser import NumberSpec, Piece, ScientificSpec

_UNFILLED = {"0": "0", "?": " ", "#": ""}
_GENERAL_SCI_AT = Decimal("99999999999.5")  # 12 integer digits switch to scientific
_GENERAL_INT_AT = Decimal("9999999999.5")  # 11 integer digits show as a rounded integer
_GENERAL_SMALL = Decimal("0.0001")


def render_number(spec: NumberSpec, magnitude: Decimal, negative: bool) -> list[Piece]:
    """Pieces for a number section. ``magnitude`` is non-negative and already 15-digit limited."""
    with localcontext() as ctx:
        ctx.prec = 400  # a double can need 309 integer digits plus the decimals
        v = magnitude
        if spec.percent:
            v *= Decimal(100) ** spec.percent
        if spec.scale:
            v /= Decimal(1000) ** spec.scale
        q = v.quantize(Decimal(1).scaleb(-spec.n_decimal), rounding=ROUND_HALF_UP)
        int_part = int(q)
        text = f"{q:f}"
    out: list[Piece] = []
    if negative and q != 0:
        out.append(("lit", "-"))
    out.extend(spec.leading)
    out.extend(_integer_pieces(spec, str(int_part) if int_part else ""))
    if spec.has_point:
        out.append(("lit", "."))
    if spec.n_decimal:
        out.extend(_decimal_pieces(spec, text.partition(".")[2]))
    out.extend(spec.trailing)
    return out


def _integer_pieces(spec: NumberSpec, digits: str) -> list[Piece]:
    placeholders = [p[1] for p in spec.integer if p[0] == "digit"]
    n = len(placeholders)
    if n == 0:
        cells = list(digits)
        extra_count = len(cells)
    else:
        extra = digits[:-n] if len(digits) > n else ""
        own = digits[-n:] if digits else ""
        own = own.rjust(n, "\0")
        cells = list(extra) + [
            c if c != "\0" else _UNFILLED[ph] for c, ph in zip(own, placeholders, strict=True)
        ]
        extra_count = len(extra)
    total = len(cells)

    def cell(i: int) -> str:
        if spec.grouping and i and (total - i) % 3 == 0:
            left = cells[i - 1]
            sep = "," if left.isdigit() else (" " if left == " " else "")
            return sep + cells[i]
        return cells[i]

    out: list[Piece] = []
    k = extra_count
    first = True
    for piece in spec.integer:
        if piece[0] != "digit":
            out.append(piece)
            continue
        if first:
            out.extend(("lit", cell(i)) for i in range(extra_count))
            first = False
        out.append(("lit", cell(k)))
        k += 1
    if n == 0:
        out = [("lit", cell(i)) for i in range(total)] + out
    return out


def _decimal_pieces(spec: NumberSpec, frac: str) -> list[Piece]:
    placeholders = [p[1] for p in spec.decimal if p[0] == "digit"]
    chars = list(frac)
    for i in range(len(chars) - 1, -1, -1):
        if chars[i] == "0" and placeholders[i] in "#?":
            chars[i] = _UNFILLED[placeholders[i]]
        else:
            break
    out: list[Piece] = []
    k = 0
    for piece in spec.decimal:
        if piece[0] == "digit":
            out.append(("lit", chars[k]))
            k += 1
        else:
            out.append(piece)
    return out


def _strip(text: str) -> str:
    return text.rstrip("0").rstrip(".") if "." in text else text


def _sci(a: Decimal) -> str:
    m = a.quantize(Decimal(1).scaleb(a.adjusted() - 5), rounding=ROUND_HALF_UP)
    e = m.adjusted()
    mant = _strip(f"{m.scaleb(-e):f}")
    return f"{mant}E{'-' if e < 0 else '+'}{abs(e):02d}"


def general_text_fitting(
    x: Decimal,
    width: float,
    measure: Callable[[str], float],
    sci_width: float | None = None,
) -> str | None:
    """``General`` reduced to fit ``width`` under ``measure``; ``None`` when nothing fits.

    Derived from the 986 width rows of ``width_general`` (CAPTURING.md section 2 item 23).
    Starting from the unconstrained text: a decimal form drops decimals one at a time, rounding
    half-up and stripping trailing zeros; when the decimals are gone the rounded integer is shown
    if it fits; when the integer part itself does not fit, scientific notation drops mantissa
    digits down to ``dE+ee``. Below one, a decimal whose digits all round away goes scientific if
    that fits and shows ``0`` (``-0`` for negatives) otherwise, and a decimal that carried into a
    new power of ten (``0.00005`` as ``0.0001``) or shows fewer significant digits than the
    scientific form that fits (``0.0000002`` against ``1.5E-07``) loses to it. A value whose
    unconstrained form is already scientific stays scientific; when not even ``dE+ee`` fits it
    shows ``0`` below width 5 if it is below one, else nothing fits (the caller's ``#`` run).
    """
    full = general_text(x)
    if measure(full) <= width:
        return full
    sign = "-" if x < 0 else ""
    a = abs(x)
    # Among literals Excel admits the scientific form of a value whose integer part does not
    # fit without charging the sign (General"x" with -123456 at width 6 shows -1E+05x), while
    # a value that rounds to zero still pays for it (-0.001 shows 0x). Bare General charges it
    # in both places (-12345678 at width 5 is a '#' run). ``sci_width`` is that budget.
    sci_budget = width if sci_width is None else sci_width
    with localcontext() as ctx:
        ctx.prec = 400
        if "E" in full:
            sci = _sci_fitting(a, sign, width, measure)
            if sci is None and a < 1 and width < 5:
                return sign + "0"
            return sci
        n_decimals = len(full.partition(".")[2])
        for d in range(n_decimals - 1, -1, -1):
            q = a.quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
            text = _strip(f"{q:f}") if d else str(int(q))
            if measure(sign + text) > width:
                continue
            if q == 0:
                sci = _sci_fitting(a, sign, width, measure)
                return sci if sci is not None else sign + "0"
            if a < 1:
                sci = _sci_fitting(a, sign, width, measure)
                carried = q < 1 and q.adjusted() != a.adjusted()  # 0.99999 at width 6 is '1'
                if sci is not None and (carried or _digits(text) < _digits(sci)):
                    return sci
            return sign + text
        return _sci_fitting(a, sign, sci_budget, measure)


def _digits(text: str) -> int:
    """Significant digits shown by a General text (leading zeros and the exponent excluded)."""
    mantissa = text.partition("E")[0]
    return len(mantissa.replace("-", "").replace(".", "").lstrip("0"))


def _sci_fitting(
    a: Decimal, sign: str, width: float, measure: Callable[[str], float]
) -> str | None:
    """Scientific notation with as many mantissa decimals as fit (at most five), else None."""
    for decimals in range(5, -1, -1):
        m = a.quantize(Decimal(1).scaleb(a.adjusted() - decimals), rounding=ROUND_HALF_UP)
        e = m.adjusted()
        mant = _strip(f"{m.scaleb(-e):f}")
        text = f"{sign}{mant}E{'-' if e < 0 else '+'}{abs(e):02d}"
        if measure(text) <= width:
            return text
    return None


def general_text(x: Decimal) -> str:
    """``General`` with no width limit: Excel's 11-character rule. ``x`` is 15-digit limited."""
    with localcontext() as ctx:
        ctx.prec = 400  # a double can need 309 integer digits plus the decimals
        a = abs(x)
        if a == 0:
            return "0"
        if a >= _GENERAL_SCI_AT:
            s = _sci(a)
        elif a == a.to_integral_value():
            s = str(int(a))
        elif a >= _GENERAL_INT_AT:
            s = str(int(a.quantize(Decimal(1), rounding=ROUND_HALF_UP)))
        elif a >= 1:
            q = a.quantize(Decimal(1).scaleb(a.adjusted() - 9), rounding=ROUND_HALF_UP)
            s = _strip(f"{q:f}") if q < _GENERAL_INT_AT else str(int(q))
        else:
            exponent = a.normalize().as_tuple().exponent
            fits = a >= _GENERAL_SMALL or (isinstance(exponent, int) and exponent >= -9)
            if fits:
                q = a.quantize(Decimal("1e-9"), rounding=ROUND_HALF_UP)
                s = _strip(f"{q:f}")
            else:
                s = _sci(a)
        if s in ("0", "0.0"):
            s = "0"
        return "-" + s if x < 0 and s != "0" else s


def _last_digit(pieces: tuple[Piece, ...]) -> int:
    return max(i for i, p in enumerate(pieces) if p[0] == "digit")


def render_scientific(spec: ScientificSpec, magnitude: Decimal, negative: bool) -> list[Piece]:
    """Pieces for a scientific section (ARCHITECTURE.md section 6).

    With ``n`` integer placeholders the mantissa has ``(exponent mod n) + 1`` integer digits,
    so ``##0.0E+0`` is engineering notation; zero renders every placeholder as ``0``; the
    exponent is padded to its placeholder count and signed per ``E+`` / ``E-``.
    """
    m = spec.mantissa
    n_int = max(m.n_integer, 1)
    with localcontext() as ctx:
        ctx.prec = 400
        v = magnitude  # a % in a scientific section is a literal and does not scale (captured)
        if v == 0:
            exponent = 0
            coefficient = Decimal(0)
        else:
            e = v.adjusted()
            int_digits = e % n_int + 1
            exponent = e - (int_digits - 1)
            coefficient = v.scaleb(-exponent)
            # rounding may carry (9.95 under ##0.0E+0 is 10.0E+0; 99.95 under 0.0E+0 is 1.0E+2)
            q = coefficient.quantize(Decimal(1).scaleb(-m.n_decimal), rounding=ROUND_HALF_UP)
            if q.adjusted() + 1 > int_digits and n_int == 1:
                exponent += 1
                coefficient = v.scaleb(-exponent)
            elif q.adjusted() + 1 > n_int:
                exponent += n_int
                coefficient = v.scaleb(-exponent)
    out: list[Piece] = []
    if negative and magnitude != 0:
        out.append(("lit", "-"))
    if magnitude == 0:
        zero_spec = NumberSpec(
            leading=m.leading,
            integer=tuple(("digit", "0") if p[0] == "digit" else p for p in m.integer),
            decimal=m.decimal,  # only integer placeholders turn into zeros ('#.#E+#' is '0.E+0')
            trailing=m.trailing,
            has_point=m.has_point,
            grouping=m.grouping,
            scale=0,
            percent=0,
            n_integer=m.n_integer,
            n_decimal=m.n_decimal,
        )
        out.extend(render_number(zero_spec, Decimal(0), False))
    else:
        plain = NumberSpec(
            leading=m.leading,
            integer=m.integer,
            decimal=m.decimal,
            trailing=m.trailing,
            has_point=m.has_point,
            grouping=m.grouping,
            scale=0,
            percent=0,
            n_integer=m.n_integer,
            n_decimal=m.n_decimal,
        )
        out.extend(render_number(plain, coefficient, False))
    sign = "-" if exponent < 0 else ("+" if spec.exp_text[1] == "+" else "")
    out.append(("lit", spec.exp_text[0] + sign))
    digits = str(abs(exponent))  # at least one digit, even for '#' placeholders ('1.E+0')
    exponent_pieces = spec.exponent
    if exponent == 0 and not any(p == ("digit", "0") for p in spec.exponent):
        exponent_pieces = tuple(
            ("digit", "0") if p[0] == "digit" and i == _last_digit(spec.exponent) else p
            for i, p in enumerate(spec.exponent)
        )
    exp_spec = NumberSpec(
        leading=(),
        integer=exponent_pieces,
        decimal=(),
        trailing=(),
        has_point=False,
        grouping=False,
        scale=0,
        percent=0,
        n_integer=spec.n_exponent,
        n_decimal=0,
    )
    out.extend(render_number(exp_spec, Decimal(digits), False))
    out.extend(spec.trailing)
    return out
