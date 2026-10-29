"""Section selection and dispatch to the renderers. See ``ARCHITECTURE.md`` section 4."""

from __future__ import annotations

from decimal import Decimal

from xlformat._assemble import Align, FormattedCell, assemble, overflow
from xlformat._datetime import render_datetime
from xlformat._fraction import render_fraction
from xlformat._measure import (
    FontSpec,
    Measure,
    character_measure,
    column_budget,
    column_units,
    glyph_measure,
)
from xlformat._number import general_text, render_number, render_scientific
from xlformat._parser import NumberFormat, Piece, Section, SectionKind
from xlformat._quirks import Quirks
from xlformat._text import render_text_section
from xlformat._values import Kind, normalize


def _holds(section: Section, x: Decimal) -> bool:
    assert section.condition is not None
    op, rhs = section.condition
    if op == "<":
        return x < rhs
    if op == "<=":
        return x <= rhs
    if op == ">":
        return x > rhs
    if op == ">=":
        return x >= rhs
    if op == "<>":
        return x != rhs
    return x == rhs


def _implies_negative(section: Section) -> bool:
    """True when the section's condition can only hold for negative numbers.

    Excel then treats the section like the implicit negative section and drops the sign:
    ``[<0]0.0`` shows ``1.0`` for -1, ``[<=-1]0`` shows ``1``, but ``[<=0]0`` shows ``-1``.
    """
    assert section.condition is not None
    op, rhs = section.condition
    return (op == "<" and rhs <= 0) or (op == "<=" and rhs < 0)


def _complement_non_positive(section: Section) -> bool:
    """True when every value failing the section's condition is zero or negative."""
    assert section.condition is not None
    op, rhs = section.condition
    return op in (">", ">=") and rhs <= 0


class Overflow:
    """Marker: the value cannot be shown by the format (renders as a run of ``#``)."""


OVERFLOW = Overflow()


def select(fmt: NumberFormat, x: Decimal) -> tuple[Section | Overflow | None, bool]:
    """Choose the section for a number.

    Returns the section (None means General, ``OVERFLOW`` means a ``#`` run) and whether the
    sign is kept. Every branch is pinned by captured rows (``ARCHITECTURE.md`` section 4;
    suites ``numeric_core``, ``conditions_rounding``, ``lone_zero``, ``selection_fallback``).
    """
    numeric = fmt.numeric
    n = len(numeric)
    if n == 0:
        return None, True
    s0 = numeric[0]
    conditioned0 = s0.condition is not None
    conditioned1 = n >= 2 and numeric[1].condition is not None
    if not conditioned0 and not conditioned1:
        # The classic layouts: 1 section for everything; 2 sections for >= 0 and < 0; 3 or 4
        # for > 0, < 0 and = 0. The negative section renders the magnitude without a sign.
        if n == 1 or (n == 2 and x >= 0) or (n >= 3 and x > 0):
            return s0, True
        if x < 0:
            return numeric[1], False
        return numeric[2], True
    if s0.condition is not None:
        if _holds(s0, x):
            return s0, not _implies_negative(s0)
        if n == 1:
            # A lone condition that fails: General, except that a General body shows '#'.
            if s0.kind is SectionKind.GENERAL and s0.has_general:
                return OVERFLOW, True
            return None, s0.condition[1] >= 0  # a negative bound drops the sign (numeric_core)
    elif x > 0:
        # An unconditioned first section before a conditioned second one takes the positives
        # only: 0;[>1]0 shows '#' for 0 and 0;[>1]0;0.0 shows '0.0' (selection_fallback).
        return s0, True
    s1 = numeric[1]
    if conditioned1:
        if _holds(s1, x):
            return s1, not _implies_negative(s1)
        if n == 2:
            # Both sections refused the value: '#' ([<=0]0;[>1]0 with 0.3; 0;[>1]0 with -2).
            return OVERFLOW, True
        return numeric[2], True  # the third section takes the rest with its sign (-0.3 -> -0.3)
    if n == 2:
        # "everything else" after a failed condition keeps the sign, unless the failing range
        # holds only non-positive values ([>0]0;"neg" shows "neg", [>100]0.00;0 shows "-1").
        return s1, not _complement_non_positive(s0)
    if x < 0:
        return s1, False  # with a third section present, the second is the negative one
    return numeric[2], True  # ...and the third takes zero and the positives that failed


def _renders_zero(cell: FormattedCell) -> bool:
    """True when the width-less text shows no nonzero digit (a ``#`` run is not zero)."""
    return not cell.overflow and not any(c in "123456789" for c in cell.full_text)


def resolve_width(
    width: int | None,
    column_width: float | None,
    font: FontSpec | None,
    measure: Measure | None,
) -> tuple[float | None, Measure, float | None]:
    """Budget, measure and column width in units for the ways a caller limits width (§8).

    The units are ``None`` outside Excel mode, where the budget is the width in the measure's
    own units.
    ``width`` is a count of characters (or the budget in ``measure``'s own units when a measure
    is given); ``column_width`` is Excel's Column Width number, measured with the glyph table of
    ``font``. Passing both raises ``ValueError``.
    """
    if width is not None and column_width is not None:
        msg = "pass width (characters) or column_width (Excel units), not both"
        raise ValueError(msg)
    if column_width is not None:
        if measure is not None:
            msg = "a custom measure goes with width, not column_width"
            raise ValueError(msg)
        if column_width <= 0:
            msg = f"column_width must be positive, not {column_width!r}"
            raise ValueError(msg)
        return (
            column_budget(column_width, font),
            glyph_measure(font),
            column_units(column_width, font),
        )
    if width is not None and (not isinstance(width, int) or isinstance(width, bool) or width < 1):
        msg = f"width must be a positive int or None, not {width!r}"
        raise ValueError(msg)
    budget = None if width is None else float(width)
    return budget, (measure or character_measure), None


def render(
    fmt: NumberFormat,
    value: object,
    *,
    width: int | None = None,
    column_width: float | None = None,
    font: FontSpec | None = None,
    measure: Measure | None = None,
    date1904: bool = False,
    quirks: Quirks | None = None,
) -> FormattedCell:
    """Render ``value`` under ``fmt``."""
    budget, measure, units = resolve_width(width, column_width, font, measure)
    v = normalize(value, date1904=date1904, quirks=quirks)
    if v.kind is Kind.EMPTY:
        return assemble([], None, "left", width=budget, measure=measure)
    if v.kind is not Kind.NUMBER:
        text = v.text or ""
        align: Align = "left" if v.kind is Kind.TEXT else "center"
        section = fmt.text_section
        is_text = v.kind is Kind.TEXT  # booleans overflow like numbers, text never does...
        if section is None or text == "":  # an empty string shows nothing, literals included
            return assemble(
                [("lit", text)], None, align, width=budget, measure=measure, overflows=not is_text
            )
        pieces = render_text_section(section, text, None)
        # ...except through a section whose '*' fill comes after the text, which shows '#' when
        # the text does not fit ('@*-' with 'Hello World' at width 6 is '#', '*-@' is the text).
        kinds = [k for k, _ in pieces]
        fill_after_text = "fill" in kinds and kinds.index("fill") > kinds.index("lit")
        overflows = not is_text or fill_after_text
        return assemble(
            pieces, section.color, align, width=budget, measure=measure, overflows=overflows
        )
    assert v.number is not None
    x = v.number
    chosen, keep_sign = select(fmt, x)
    cell = _render_number(
        fmt, chosen, keep_sign, x, v.raw, date1904, quirks, budget, measure, units
    )
    if (
        keep_sign
        and x < 0
        and isinstance(chosen, Section)
        and chosen.kind is SectionKind.NUMBER
        and _renders_zero(cell)
    ):
        # Excel renders a negative as its magnitude plus a minus sign; when the magnitude
        # renders as zero it renders the positive magnitude again, from section selection
        # onward (lone_zero and conditions_rounding suites): [<=0]0 shows 0.001 for -0.001,
        # [<=0.1]0 shows 0.3 for -0.3, [<=0]0;[>1]0 shows a '#' run for -0.3, [<=0.5]0 shows 0.
        # [<0]0 shows 0 because its sign was already dropped, so nothing is re-rendered.
        magnitude = abs(x)
        chosen, keep_sign = select(fmt, magnitude)
        cell = _render_number(
            fmt, chosen, keep_sign, magnitude, abs(v.raw), date1904, quirks, budget, measure, units
        )
    return cell


def _render_number(
    fmt: NumberFormat,
    chosen: Section | Overflow | None,
    keep_sign: bool,
    x: Decimal,
    raw: float,
    date1904: bool,
    quirks: Quirks | None,
    width: float | None = None,
    measure: Measure = character_measure,
    units: float | None = None,
) -> FormattedCell:
    """Render a number under the selected section."""
    sole = len(fmt.numeric) <= 1  # a lone bare General section is the General format
    if isinstance(chosen, Overflow):
        return overflow(None, width, None, measure)
    negative = keep_sign and x < 0
    magnitude = abs(x)
    pieces: list[Piece]
    general_value: Decimal | None = None
    if chosen is None:
        general_value = x if keep_sign else magnitude
        pieces = [("general", general_text(general_value))]
    elif chosen.kind is SectionKind.EMPTY:
        pieces = []
    elif chosen.kind is SectionKind.NUMBER:
        assert chosen.number is not None
        pieces = render_number(chosen.number, magnitude, negative)
    elif chosen.kind is SectionKind.SCIENTIFIC:
        assert chosen.scientific is not None
        pieces = render_scientific(chosen.scientific, magnitude, negative)
    elif chosen.kind is SectionKind.FRACTION:
        assert chosen.fraction is not None
        rendered = render_fraction(chosen.fraction, raw, negative, quirks or Quirks.EXCEL)
        if rendered is None:
            return overflow(chosen.color, width, None, measure)
        pieces = rendered
    elif chosen.kind in (SectionKind.GENERAL, SectionKind.TEXT):
        pieces = render_text_section(chosen, None, magnitude)
        if len(pieces) == 1 and pieces[0][0] == "general":
            # Bare General converts the signed value: -0.05 at width 3 is "-0" (width_general).
            general_value = x if keep_sign else magnitude
            pieces = [("general", general_text(general_value))]
        else:
            if chosen.has_general:
                general_value = magnitude
            if negative and magnitude != 0:
                # Dropped when the General reduces to zero at a width: General"x" shows "0x".
                pieces.insert(0, ("sign", "-"))
    elif chosen.kind is SectionKind.DATETIME:
        signed_raw = raw if keep_sign else abs(raw)
        rendered = render_datetime(
            chosen, signed_raw, magnitude, date1904=date1904, quirks=quirks or Quirks.EXCEL
        )
        if rendered is None:
            return overflow(chosen.color, width, None, measure)
        pieces = rendered
    return assemble(
        pieces,
        chosen.color if chosen else None,
        "right",
        width=width,
        measure=measure,
        general_value=general_value,
        units=units,
        general_format=chosen is None
        or (sole and chosen.kind is SectionKind.GENERAL and len(pieces) == 1),
    )
