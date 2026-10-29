"""Format-string parsing: sections, tokens, classification and the compiled ``NumberFormat``.

The pipeline is described in ``ARCHITECTURE.md`` sections 2 to 4. Parsing is pure and cached;
``FormatError`` is raised only for strings Excel refuses to open from a file (section 11 there,
evidence in ``CAPTURING.md`` section 2 item 9).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from enum import Enum
from functools import lru_cache
from typing import TYPE_CHECKING

from xlformat._errors import FormatError

if TYPE_CHECKING:
    from xlformat._assemble import FormattedCell
    from xlformat._quirks import Quirks
from xlformat._measure import FontSpec, Measure


class TokenKind(Enum):
    """What a token means. ``MINUTE`` exists only after ``m`` resolution."""

    # members are singletons compared by identity, so identity hashing is equivalent and avoids
    # Enum's Python-level __hash__ in the parser's per-kind dictionaries
    __hash__ = object.__hash__

    LITERAL = "literal"
    DIGIT = "digit"  # 0 # ?
    POINT = "point"
    COMMA = "comma"
    PERCENT = "percent"
    EXP = "exp"  # E+ E- (e+ e- rejected when digits are present)
    SLASH = "slash"
    GENERAL = "general"
    AT = "at"
    DATE = (
        "date"  # normalized code: y yy yyyy m mm mmm mmmm mmmmm d dd ddd dddd h hh s ss e, s.0 ...
    )
    MINUTE = "minute"  # m or mm resolved as minutes
    AMPM = "ampm"  # AM/PM or A/P
    ELAPSED = "elapsed"  # [h] [mm] [sss] ...
    COLOR = "color"
    CONDITION = "condition"
    CURRENCY = "currency"  # [$sym-locale]
    PAD = "pad"  # _x
    FILL = "fill"  # *x


@dataclass(frozen=True, slots=True)
class Token:
    """One token. ``text`` is the literal text, the normalized code, or the color name."""

    kind: TokenKind
    text: str
    condition: tuple[str, Decimal] | None = None
    locale: str | None = None


class SectionKind(Enum):
    """How a section renders. Classification order is fixed; see ``ARCHITECTURE.md`` section 3."""

    EMPTY = "empty"
    DATETIME = "datetime"
    FRACTION = "fraction"
    SCIENTIFIC = "scientific"
    NUMBER = "number"
    TEXT = "text"
    GENERAL = "general"


Piece = tuple[str, str]
"""A render piece: ("lit", text), ("digit", "0#?"), ("pad", ch), ("fill", ch), ("general", "")."""


@dataclass(frozen=True, slots=True)
class NumberSpec:
    """A number section compiled for rendering (``_number.py``)."""

    leading: tuple[Piece, ...]
    integer: tuple[Piece, ...]
    decimal: tuple[Piece, ...]
    trailing: tuple[Piece, ...]
    has_point: bool
    grouping: bool
    scale: int  # divide by 1000 this many times
    percent: int  # multiply by 100 this many times
    n_integer: int
    n_decimal: int


@dataclass(frozen=True, slots=True)
class ScientificSpec:
    """A scientific section: the mantissa as a NumberSpec plus the exponent part."""

    mantissa: NumberSpec
    exp_text: str  # "E+" or "E-"
    exponent: tuple[Piece, ...]  # digit placeholders and literals after the exponent marker
    n_exponent: int
    trailing: tuple[Piece, ...]


@dataclass(frozen=True, slots=True)
class FractionSpec:
    """A fraction section split into its parts (see ``_fraction.py``)."""

    leading: tuple[Piece, ...]
    whole: NumberSpec | None  # integer placeholders before the fraction, or None
    whole_trailing: tuple[Piece, ...]  # literals between the whole part and the numerator
    numerator: str  # placeholder characters, e.g. "??"
    slash: tuple[Piece, ...]  # pieces between numerator and denominator, usually ("lit", "/")
    denominator: str  # placeholder characters, or the digits of a fixed denominator
    fixed_denominator: int | None
    trailing: tuple[Piece, ...]
    percent: int


@dataclass(frozen=True, slots=True)
class Section:
    """One compiled section of a format string.

    Attributes:
        kind: How the section renders: number, scientific, fraction, date and time, text,
            ``General`` or empty.
        tokens: The section's tokens, in order (internal).
        color: The section color in canonical spelling (``'Red'``, ``'Color12'``), or ``None``.
        condition: The condition as ``(operator, bound)``, such as ``('>=', Decimal('100'))``,
            or ``None``.
        locale: The locale or system tag of ``[$-...]``, such as ``'409'`` or ``'x-sysdate'``;
            parsed but not applied, except for the system date and time tags.
        number: Compiled number layout, for number sections (internal).
        scientific: Compiled mantissa and exponent, for scientific sections (internal).
        fraction: Compiled fraction parts, for fraction sections (internal).
        has_text_placeholder: True when the section contains ``@``.
        has_general: True when the section contains ``General``.
        source: The section's text as written in the format string.
    """

    kind: SectionKind
    tokens: tuple[Token, ...]
    color: str | None = None
    condition: tuple[str, Decimal] | None = None
    locale: str | None = None
    number: NumberSpec | None = None
    scientific: ScientificSpec | None = None
    fraction: FractionSpec | None = None
    has_text_placeholder: bool = False
    has_general: bool = False
    source: str = ""


_COLORS = {
    c.lower(): c for c in ("Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow")
}
_COLOR_INDEX_RE = re.compile(r"^color\s*([1-9]\d*)$", re.I)
_CONDITION_RE = re.compile(r"^(<=|>=|<>|<|>|=)\s*(\S+)$")
_ELAPSED_RE = re.compile(r"^([hms])\1*$", re.I)

_TOKEN_RE = re.compile(
    r"""
    "(?P<quoted>[^"]*)"
  | (?P<badquote>"[^"]*)$
  | \\(?P<escaped>.)
  | (?P<badescape>\\)$
  | \[(?P<bracket>[^\]]*)\]
  | (?P<badbracket>\[[^\]]*)$
  | _(?P<pad>.)
  | (?P<badpad>_)$
  | \*(?P<fill>.)
  | (?P<badfill>\*)$
  | (?P<general>general)
  | (?P<at>@)
  | (?P<ampm>am/pm|a/p)
  | (?P<exp>e[+-])
  | (?P<digit>[0#?])
  | (?P<date>y+|m+|d+|h+|s+|e+)
  | (?P<char>.)
    """,
    re.X | re.S | re.I,
)


def _normalize_date_code(run: str) -> str:
    c = run[0].lower()
    n = len(run)
    if c == "y":
        return "yy" if n <= 2 else "yyyy"
    if c == "m":
        return "m" * min(n, 5) if n <= 5 else "mmmm"
    if c == "d":
        return "d" * min(n, 4)
    if c == "h":
        return "h" if n == 1 else "hh"
    if c == "s":
        return "s" if n == 1 else "ss"
    return "e"


def _bracket_token(inner: str, raw: str) -> Token:
    stripped = inner.strip()
    if not stripped:
        msg = f"empty bracket {raw!r}"
        raise FormatError(msg)
    if stripped.lower() in _COLORS:
        return Token(TokenKind.COLOR, _COLORS[stripped.lower()])
    m = _COLOR_INDEX_RE.match(stripped)
    if m:
        index = int(m.group(1))
        if index > 56:
            msg = f"color index out of range in {raw!r}"
            raise FormatError(msg)
        return Token(TokenKind.COLOR, f"Color{index}")
    m = _CONDITION_RE.match(stripped)
    if m:
        try:
            rhs = Decimal(m.group(2))
        except InvalidOperation:
            msg = f"invalid condition {raw!r}"
            raise FormatError(msg) from None
        return Token(TokenKind.CONDITION, raw, condition=(m.group(1), rhs))
    if stripped.startswith("$"):
        symbol, _, locale = stripped[1:].partition("-")
        return Token(TokenKind.CURRENCY, symbol, locale=locale or None)
    if _ELAPSED_RE.match(stripped):
        return Token(TokenKind.ELAPSED, stripped.lower())
    msg = f"unrecognized bracket {raw!r}"
    raise FormatError(msg)


# regex groups whose token is their text under one kind; the rest are handled one by one
_SIMPLE_KINDS: dict[str, TokenKind] = {
    "escaped": TokenKind.LITERAL,
    "pad": TokenKind.PAD,
    "fill": TokenKind.FILL,
    "ampm": TokenKind.AMPM,  # spelling kept: am/pm renders lower-case
    "exp": TokenKind.EXP,
    "digit": TokenKind.DIGIT,
}
_CHAR_KINDS: dict[str, TokenKind] = {
    ".": TokenKind.POINT,
    ",": TokenKind.COMMA,
    "%": TokenKind.PERCENT,
    "/": TokenKind.SLASH,
}


def tokenize(section: str) -> list[Token]:
    """Split one section into tokens. Raises ``FormatError`` for unclosed quotes and brackets."""
    tokens: list[Token] = []
    for m in _TOKEN_RE.finditer(section):
        kind = m.lastgroup
        text = m.group(kind) if kind else ""
        simple = _SIMPLE_KINDS.get(kind) if kind else None
        if simple is not None:
            tokens.append(Token(simple, text))
        elif kind == "char":
            tokens.append(Token(_CHAR_KINDS.get(text, TokenKind.LITERAL), text))
        elif kind == "quoted":
            if text:
                tokens.append(Token(TokenKind.LITERAL, text))
        elif kind in ("badquote", "badbracket"):
            msg = f"unclosed {text[0]!r} in {section!r}"
            raise FormatError(msg)
        elif kind in ("badescape", "badpad", "badfill"):
            msg = f"dangling {text!r} at end of {section!r}"
            raise FormatError(msg)
        elif kind == "bracket":
            tokens.append(_bracket_token(text, m.group(0)))
        elif kind == "general":
            tokens.append(Token(TokenKind.GENERAL, "General"))
        elif kind == "at":
            tokens.append(Token(TokenKind.AT, "@"))
        else:  # date
            tokens.append(Token(TokenKind.DATE, _normalize_date_code(text)))
    return tokens


def split_sections(format_string: str) -> list[str]:
    """Split on ``;`` outside quotes, escapes and brackets."""
    sections: list[str] = []
    current: list[str] = []
    i = 0
    n = len(format_string)
    while i < n:
        ch = format_string[i]
        if ch == '"':
            end = format_string.find('"', i + 1)
            if end < 0:
                current.append(format_string[i:])
                break
            current.append(format_string[i : end + 1])
            i = end + 1
            continue
        if ch == "[":
            end = format_string.find("]", i + 1)
            if end < 0:
                current.append(format_string[i:])
                break
            current.append(format_string[i : end + 1])
            i = end + 1
            continue
        if ch in "\\_*" and i + 1 < n:
            current.append(format_string[i : i + 2])
            i += 2
            continue
        if ch == ";":
            sections.append("".join(current))
            current = []
            i += 1
            continue
        current.append(ch)
        i += 1
    sections.append("".join(current))
    if len(sections) > 4:
        msg = f"more than four sections in {format_string!r}"
        raise FormatError(msg)
    return sections


def _resolve_datetime(tokens: list[Token], source: str) -> list[Token]:
    """Fold fractional seconds, reject stray digits, resolve ``m`` as month or minute."""
    out: list[Token] = []
    i = 0
    while i < len(tokens):
        t = tokens[i]
        if (
            t.kind is TokenKind.POINT
            and out
            and out[-1].kind is TokenKind.DATE
            and out[-1].text in ("s", "ss")
            and i + 1 < len(tokens)
            and tokens[i + 1].kind is TokenKind.DIGIT
            and tokens[i + 1].text == "0"
        ):
            zeros = 0
            j = i + 1
            while j < len(tokens) and tokens[j].kind is TokenKind.DIGIT and tokens[j].text == "0":
                zeros += 1
                j += 1
            if zeros > 3:
                msg = f"more than three fractional-second digits in {source!r}"
                raise FormatError(msg)
            out[-1] = Token(TokenKind.DATE, out[-1].text + "." + "0" * zeros)
            i = j
            continue
        if t.kind is TokenKind.DIGIT:
            msg = f"digit placeholder in date/time section {source!r}"
            raise FormatError(msg)
        if t.kind is TokenKind.EXP:  # ``e+`` without digits: year code then literal
            out.append(Token(TokenKind.DATE, "e"))
            out.append(Token(TokenKind.LITERAL, t.text[1]))
            i += 1
            continue
        out.append(t)
        i += 1
    tokens = out
    date_kinds = (TokenKind.DATE, TokenKind.ELAPSED)
    last_m = -1
    for ix, t in enumerate(tokens):
        if t.kind is TokenKind.DATE and t.text in ("m", "mm"):
            minute = False
            for back in range(ix - 1, last_m, -1):
                prev = tokens[back]
                if prev.kind is TokenKind.DATE and prev.text in ("h", "hh"):
                    minute = True
                    break
                if prev.kind is TokenKind.ELAPSED and prev.text[0] == "h":
                    minute = True
                    break
            if not minute:
                for fwd in range(ix + 1, len(tokens)):
                    nxt = tokens[fwd]
                    if nxt.kind in date_kinds:
                        minute = nxt.kind is TokenKind.DATE and nxt.text[0] == "s"
                        break
            if minute:
                tokens[ix] = Token(TokenKind.MINUTE, t.text)
            last_m = ix
    return tokens


def _pieces(tokens: list[Token]) -> tuple[Piece, ...]:
    pieces: list[Piece] = []
    for t in tokens:
        if t.kind is TokenKind.DIGIT:
            pieces.append(("digit", t.text))
        elif t.kind is TokenKind.PAD:
            pieces.append(("pad", t.text))
        elif t.kind is TokenKind.FILL:
            pieces.append(("fill", t.text))
        elif t.kind is TokenKind.GENERAL:
            pieces.append(("general", ""))
        elif t.kind is TokenKind.CURRENCY:
            if t.text:
                pieces.append(("lit", t.text))
        elif t.kind in (TokenKind.COLOR, TokenKind.CONDITION):
            continue
        else:
            pieces.append(("lit", t.text))
    return tuple(pieces)


def _merge_commas(tokens: list[Token], as_literal: bool) -> list[Token]:
    out: list[Token] = []
    for t in tokens:
        if t.kind is TokenKind.COMMA and out and out[-1].kind is TokenKind.COMMA:
            continue
        if t.kind is TokenKind.COMMA and as_literal:
            out.append(Token(TokenKind.LITERAL, ","))
        else:
            out.append(t)
    return out


def _compile_number(tokens: list[Token], source: str) -> NumberSpec:
    """Apply Excel's comma rules and split a number section into its parts."""
    percent = sum(1 for t in tokens if t.kind is TokenKind.PERCENT)
    tokens = [Token(TokenKind.LITERAL, "%") if t.kind is TokenKind.PERCENT else t for t in tokens]
    tokens = [Token(TokenKind.LITERAL, "/") if t.kind is TokenKind.SLASH else t for t in tokens]
    dp = [i for i, t in enumerate(tokens) if t.kind in (TokenKind.DIGIT, TokenKind.POINT)]
    first, last = dp[0], dp[-1]
    point_ix = next((i for i, t in enumerate(tokens) if t.kind is TokenKind.POINT), None)
    leading = tokens[:first]
    if point_ix is None:
        integer = tokens[first : last + 1]
        decimal: list[Token] = []
    else:
        integer = tokens[first:point_ix]
        decimal = tokens[point_ix + 1 : last + 1]
    trailing = tokens[last + 1 :]
    scale = 0
    while trailing and trailing[0].kind is TokenKind.COMMA:
        scale += 1
        trailing.pop(0)
    leading = _merge_commas(leading, True)
    trailing = _merge_commas(trailing, True)
    integer = _merge_commas(integer, False)
    decimal = _merge_commas(decimal, False)
    grouping = False
    kept: list[Token] = []
    for i, t in enumerate(integer):
        if (
            t.kind is TokenKind.COMMA
            and 0 < i < len(integer) - 1
            and integer[i - 1].kind is TokenKind.DIGIT
            and integer[i + 1].kind is TokenKind.DIGIT
        ):
            grouping = True
            continue
        kept.append(t)
    integer = kept

    def drop_or_literal(part: list[Token]) -> list[Token]:
        out: list[Token] = []
        for t in part:
            if t.kind is TokenKind.COMMA:
                if not out or out[-1].kind in (TokenKind.DIGIT, TokenKind.POINT):
                    continue
                out.append(Token(TokenKind.LITERAL, ","))
            else:
                out.append(t)
        return out

    integer = drop_or_literal(integer)
    decimal = drop_or_literal(decimal)
    # later points inside the decimal part are literals (0.0.0 shows 1234.5.7)
    decimal = [Token(TokenKind.LITERAL, ".") if t.kind is TokenKind.POINT else t for t in decimal]
    return NumberSpec(
        leading=_pieces(leading),
        integer=_pieces(integer),
        decimal=_pieces(decimal),
        trailing=_pieces(trailing),
        has_point=point_ix is not None,
        grouping=grouping,
        scale=scale,
        percent=percent,
        n_integer=sum(1 for t in integer if t.kind is TokenKind.DIGIT),
        n_decimal=sum(1 for t in decimal if t.kind is TokenKind.DIGIT),
    )


def _compile_scientific(body: list[Token], exp: Token, source: str) -> ScientificSpec:
    """Split at the exponent marker: mantissa (a number spec), exponent digits, trailing."""
    at = body.index(exp)
    mantissa_tokens = body[:at]
    rest = body[at + 1 :]
    digit_ix = [i for i, t in enumerate(rest) if t.kind is TokenKind.DIGIT]
    exponent = rest[: digit_ix[-1] + 1]
    trailing = rest[digit_ix[-1] + 1 :]
    if not any(t.kind is TokenKind.DIGIT for t in mantissa_tokens):
        msg = f"exponent without mantissa digits in {source!r}"
        raise FormatError(msg)
    return ScientificSpec(
        mantissa=_compile_number(mantissa_tokens, source),
        exp_text=exp.text,
        exponent=_pieces(exponent),
        n_exponent=len(digit_ix),
        trailing=_pieces(trailing),
    )


def _compile_fraction(body: list[Token], source: str) -> FractionSpec:
    """Split a fraction section around its bar."""
    percent = sum(1 for t in body if t.kind is TokenKind.PERCENT)
    body = [Token(TokenKind.LITERAL, "%") if t.kind is TokenKind.PERCENT else t for t in body]
    slash = next(i for i, t in enumerate(body) if t.kind is TokenKind.SLASH)

    def is_digit(t: Token) -> bool:
        return t.kind is TokenKind.DIGIT

    # numerator: the digit run nearest before the bar
    num_end = slash
    while num_end > 0 and not is_digit(body[num_end - 1]):
        num_end -= 1
    num_start = num_end
    while num_start > 0 and is_digit(body[num_start - 1]):
        num_start -= 1
    if num_start == num_end:
        msg = f"fraction without numerator in {source!r}"
        raise FormatError(msg)
    # denominator: placeholders or literal digits right after the bar
    den_start = slash + 1
    while den_start < len(body) and not (
        is_digit(body[den_start])
        or (body[den_start].kind is TokenKind.LITERAL and body[den_start].text.isdigit())
    ):
        den_start += 1
    den_end = den_start
    fixed = body[den_start].kind is TokenKind.LITERAL
    while den_end < len(body) and (
        (not fixed and is_digit(body[den_end]))
        or (
            fixed
            and (
                (body[den_end].kind is TokenKind.LITERAL and body[den_end].text.isdigit())
                or (is_digit(body[den_end]) and body[den_end].text == "0")  # the 0 of "10"
            )
        )
    ):
        den_end += 1
    denominator = "".join(t.text for t in body[den_start:den_end])
    # whole part: every digit placeholder before the numerator run
    whole_ix = [i for i in range(num_start) if is_digit(body[i])]
    if whole_ix:
        whole_tokens = body[whole_ix[0] : whole_ix[-1] + 1]
        whole = _compile_number(whole_tokens, source)
        leading = body[: whole_ix[0]]
        whole_trailing = body[whole_ix[-1] + 1 : num_start]
    else:
        whole = None
        leading = body[:num_start]
        whole_trailing = []
    return FractionSpec(
        leading=_pieces(_merge_commas(leading, True)),
        whole=whole,
        whole_trailing=_pieces(whole_trailing),
        numerator="".join(t.text for t in body[num_start:num_end]),
        slash=_pieces(body[num_end:den_start]),
        denominator=denominator,
        fixed_denominator=int(denominator) if fixed else None,
        trailing=_pieces(_merge_commas(body[den_end:], True)),
        percent=percent,
    )


_AT_MOST_ONE: tuple[tuple[TokenKind, str], ...] = (
    (TokenKind.COLOR, "color"),
    (TokenKind.CONDITION, "condition"),
    (TokenKind.CURRENCY, "currency or locale tag"),
    (TokenKind.GENERAL, "General"),
    (TokenKind.ELAPSED, "elapsed-time code"),
    (TokenKind.EXP, "exponent"),
)


def compile_section(source: str) -> Section:
    """Tokenize, validate and classify one section."""
    tokens = tokenize(source)
    # one pass: the count of every kind, and the first token of each
    counts: dict[TokenKind, int] = {}
    first: dict[TokenKind, Token] = {}
    for t in tokens:
        counts[t.kind] = counts.get(t.kind, 0) + 1
        first.setdefault(t.kind, t)
    for kind, what in _AT_MOST_ONE:
        if counts.get(kind, 0) > 1:
            msg = f"more than one {what} in section {source!r}"
            raise FormatError(msg)
    color_tok = first.get(TokenKind.COLOR)
    cond_tok = first.get(TokenKind.CONDITION)
    currency = first.get(TokenKind.CURRENCY)
    color = color_tok.text if color_tok else None
    condition = cond_tok.condition if cond_tok else None
    locale = currency.locale if currency else None
    if color_tok or cond_tok:
        body = [t for t in tokens if t.kind not in (TokenKind.COLOR, TokenKind.CONDITION)]
    else:
        body = tokens
    has_at = TokenKind.AT in counts
    has_general = TokenKind.GENERAL in counts
    has_digits = TokenKind.DIGIT in counts
    has_point = TokenKind.POINT in counts
    has_date = TokenKind.DATE in counts or TokenKind.ELAPSED in counts or TokenKind.AMPM in counts
    # A bare system date or time tag ([$-x-sysdate], [$-x-systime]) is itself the format: Excel
    # shows the system long date or time and ignores the rest of the section.
    if (locale or "").lower() in ("x-sysdate", "x-systime", "f800", "f400"):
        has_date = True  # [$-F800] and [$-F400] are the LCID spellings of the same tags
    exp = first.get(TokenKind.EXP)

    def section(
        kind: SectionKind,
        toks: list[Token],
        *,
        number: NumberSpec | None = None,
        scientific: ScientificSpec | None = None,
        fraction: FractionSpec | None = None,
        has_text_placeholder: bool = False,
    ) -> Section:
        return Section(
            kind,
            tuple(toks),
            color=color,
            condition=condition,
            locale=locale,
            number=number,
            scientific=scientific,
            fraction=fraction,
            has_text_placeholder=has_text_placeholder,
            has_general=has_general,
            source=source,
        )

    if not body:
        return section(SectionKind.EMPTY, [])
    if has_date or (exp is not None and not has_digits):
        body = _resolve_datetime(body, source)
        if has_at:
            msg = f"text placeholder in date/time section {source!r}"
            raise FormatError(msg)
        return section(SectionKind.DATETIME, body)
    if has_digits and has_at:
        msg = f"text placeholder in number section {source!r}"
        raise FormatError(msg)
    if has_digits and TokenKind.SLASH in counts:
        _check_fraction(body, source)
        return section(SectionKind.FRACTION, body, fraction=_compile_fraction(body, source))
    if exp is not None:
        if exp.text[0] == "e":
            msg = f"lower-case exponent {exp.text!r} in {source!r}"
            raise FormatError(msg)
        after = body[body.index(exp) + 1 :]
        if not any(t.kind is TokenKind.DIGIT for t in after):
            msg = f"exponent without digits in {source!r}"
            raise FormatError(msg)
        return section(
            SectionKind.SCIENTIFIC, body, scientific=_compile_scientific(body, exp, source)
        )
    if has_digits or has_point:
        spec = _compile_number(body, source)
        return section(SectionKind.NUMBER, body, number=spec)
    if has_at:
        return section(SectionKind.TEXT, body, has_text_placeholder=True)
    return section(SectionKind.GENERAL, body)


def _check_fraction(body: list[Token], source: str) -> None:
    slashes = [i for i, t in enumerate(body) if t.kind is TokenKind.SLASH]
    if any(t.kind is TokenKind.POINT for t in body[: slashes[0]]):
        msg = f"decimal point before the fraction bar in {source!r}"  # Excel refuses the file
        raise FormatError(msg)
    if len(slashes) > 1:
        msg = f"more than one fraction bar in {source!r}"
        raise FormatError(msg)
    after = body[slashes[0] + 1 :]
    if not any(
        t.kind is TokenKind.DIGIT or (t.kind is TokenKind.LITERAL and t.text.isdigit())
        for t in after
    ):
        msg = f"fraction without denominator in {source!r}"
        raise FormatError(msg)


@dataclass(frozen=True, slots=True)
class NumberFormat:
    """A parsed format string. Create with :meth:`parse`, which is cached.

    Attributes:
        source: The format string as given (``General`` for the empty string).
        sections: The compiled sections, one to four, in order.
        numeric: The sections that take numbers (internal).
        text_section: The section that takes text and booleans, or ``None`` (internal).
    """

    source: str
    sections: tuple[Section, ...]
    numeric: tuple[Section, ...] = field(repr=False)
    text_section: Section | None = field(repr=False)

    @classmethod
    def parse(cls, format_string: str) -> NumberFormat:
        """Compile a format string; the result is cached per string.

        Args:
            format_string: An Excel number format; ``""`` means ``General``.

        Raises:
            FormatError: The format string is one Excel refuses to open from a file.
        """
        return _parse_cached(format_string)

    @property
    def is_date(self) -> bool:
        """True when any section is a date/time section."""
        return any(s.kind is SectionKind.DATETIME for s in self.sections)

    @property
    def is_text(self) -> bool:
        """True when the first section is a text section (every value renders as text)."""
        return self.sections[0].kind is SectionKind.TEXT

    @property
    def is_general(self) -> bool:
        """True for ``General`` and the empty format."""
        first = self.sections[0]
        return (
            len(self.sections) == 1
            and first.kind is SectionKind.GENERAL
            and len(first.tokens) == 1
            and first.tokens[0].kind is TokenKind.GENERAL
        ) or (first.kind is SectionKind.EMPTY and len(self.sections) == 1)

    def format(
        self,
        value: object,
        *,
        width: int | None = None,
        column_width: float | None = None,
        font: FontSpec | None = None,
        normal_font: FontSpec | None = None,
        measure: Measure | None = None,
        date1904: bool = False,
        quirks: Quirks | None = None,
    ) -> str:
        """The text Excel displays for ``value`` under this format.

        Takes the same keyword arguments as :func:`xlformat.format_value`, which calls it.
        """
        from xlformat._render import render

        return render(
            self,
            value,
            width=width,
            column_width=column_width,
            font=font,
            normal_font=normal_font,
            measure=measure,
            date1904=date1904,
            quirks=quirks,
        ).text

    def format_cell(
        self,
        value: object,
        *,
        width: int | None = None,
        column_width: float | None = None,
        font: FontSpec | None = None,
        normal_font: FontSpec | None = None,
        measure: Measure | None = None,
        date1904: bool = False,
        quirks: Quirks | None = None,
    ) -> FormattedCell:
        """Text plus color, alignment, fill and overflow information.

        Takes the same keyword arguments as :func:`xlformat.format_value`; returns a
        :class:`xlformat.FormattedCell`, as :func:`xlformat.format_cell` does.
        """
        from xlformat._render import render

        return render(
            self,
            value,
            width=width,
            column_width=column_width,
            font=font,
            normal_font=normal_font,
            measure=measure,
            date1904=date1904,
            quirks=quirks,
        )


@lru_cache(maxsize=1024)
def _parse_cached(format_string: str) -> NumberFormat:
    if format_string == "":
        format_string = "General"
    raw_sections = split_sections(format_string)
    sections: list[Section] = []
    for raw in raw_sections:
        section = compile_section(raw)
        # An empty section renders nothing, even the first (";" hides every number). One that
        # carries a color or condition is General: Excel stores "[Red]" as "[Red]General".
        if section.kind is SectionKind.EMPTY and (
            section.color is not None or section.condition is not None
        ):
            section = Section(
                SectionKind.GENERAL,
                (Token(TokenKind.GENERAL, "General"),),
                color=section.color,
                condition=section.condition,
                has_general=True,
                source=raw,
            )
        sections.append(section)
    text_sections = [i for i, s in enumerate(sections) if s.has_text_placeholder]
    if len(text_sections) > 1 or (text_sections and text_sections[0] != len(sections) - 1):
        msg = f"text section must be the last section: {format_string!r}"
        raise FormatError(msg)
    if len(sections) == 4 or text_sections:
        text_section: Section | None = sections[-1]
        numeric = tuple(sections[:-1])
    else:
        text_section = None
        numeric = tuple(sections)
    for i, s in enumerate(numeric):
        if s.condition is not None and i > 1:
            msg = f"condition in section {i + 1} of {format_string!r}"
            raise FormatError(msg)
    if text_section is not None and text_section.condition is not None:
        msg = f"condition on the text section of {format_string!r}"
        raise FormatError(msg)
    return NumberFormat(format_string, tuple(sections), numeric, text_section)
