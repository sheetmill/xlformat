"""Turn render pieces into text and a ``FormattedCell``, applying the width decision.

``ARCHITECTURE.md`` section 8: both width modes are a measure function plus a width in its
units. A ``*`` fill repeats its character until the text measures the width (the last fill in a
section wins); ``_x`` pads measure as one space; ``General`` reduces its digits before it
overflows; numbers, dates and booleans that still do not fit become a run of ``#`` that fills
the width, text never does.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Literal

from xlformat._measure import MAX_TEXT, Measure, character_measure, general_cost
from xlformat._number import GENERAL_SMALL, general_fallback, general_text_fitting
from xlformat._parser import Piece

Align = Literal["left", "right", "center"]


@dataclass(frozen=True, slots=True)
class FormattedCell:
    """What Excel shows for one cell: the text and the display information around it.

    Attributes:
        text: What the cell shows under the given width.
        full_text: The width-less rendering (what copy and paste gives); always set.
        color: Section color, ``'Red'`` ... ``'Color56'``, or ``None``.
        align: Excel's default horizontal alignment for the value kind.
        fill: ``(index, char)`` of a ``*`` fill in ``text``, or ``None``.
        overflow: True when ``text`` is a run of ``#``.
    """

    text: str
    full_text: str
    color: str | None
    align: Align
    fill: tuple[int, str] | None
    overflow: bool


OVERFLOW_WIDTH = 255  # Excel's widest column; the clipboard shows this many '#'


def _hashes(width: float | None, measure: Measure, general: bool = False) -> str:
    """As many ``#`` as fit, at most 255.

    None in a column too narrow for one, except that General always shows one
    (``column_narrow``, ``limits``; CAPTURING.md section 2 item 28).
    """
    if width is None:
        return "#" * OVERFLOW_WIDTH
    count = min(max(int(width // measure("#")), 0), MAX_TEXT)
    return "#" * max(count, 1 if general else 0)


def overflow(
    color: str | None,
    width: float | None = None,
    full_text: str | None = None,
    measure: Measure = character_measure,
    general: bool = False,
) -> FormattedCell:
    """A value the section cannot show: a ``#`` run filling the width (255 without a width)."""
    text = _hashes(width, measure, general)
    # full_text is always the width-less rendering: 255 '#' for a value the section cannot show
    full = _hashes(None, measure) if full_text is None else full_text
    return FormattedCell(text, full, color, "right", None, True)


def _join(pieces: list[Piece], general: str | None) -> tuple[str, tuple[int, str] | None]:
    """Text with pads as one space and fills left out; the last fill's position and character."""
    parts: list[str] = []
    position = 0
    fill: tuple[int, str] | None = None
    for kind, text in pieces:
        if kind == "pad":
            parts.append(" ")
            position += 1
        elif kind == "fill":
            fill = (position, text)  # Excel honors the last fill in a section
        else:
            piece = general if (kind == "general" and general is not None) else text
            parts.append(piece)
            position += len(piece)
    return "".join(parts), fill


def assemble(
    pieces: list[Piece],
    color: str | None,
    align: Align,
    *,
    width: float | None = None,
    measure: Measure = character_measure,
    overflows: bool = True,
    general_value: Decimal | None = None,
    units: float | None = None,
    general_format: bool = False,
) -> FormattedCell:
    """Join pieces and apply the width decision.

    ``general_value`` is the number behind a ``general`` piece, for digit reduction, which
    Excel does by characters against the column width in ``units`` (``general_cost``), not by
    glyph widths; ``overflows`` is False for text, which never becomes a ``#`` run.
    ``general_format`` is True for the format ``General`` itself, as opposed to a General
    section among others or among literals.
    """
    full_text, fill = _join(pieces, None)
    if overflows and len(full_text) > MAX_TEXT:
        # Excel shows at most 255 characters: 1E+300 under '0' is 301 digits and shows a run of
        # 255 '#' without a width as well as in any column (limits)
        return overflow(color, width, None, measure, general=general_value is not None)
    if width is None:
        return FormattedCell(full_text, full_text, color, align, fill, False)
    general: str | None = None
    has_general = general_value is not None and any(k == "general" for k, _ in pieces)
    if has_general:
        assert general_value is not None
        # The General token's budget is the column width in units less what the literals (and
        # the sign) take, in digit units of this measure: '( ' and ' )' cost 1.67 in Excel mode
        # and 4 in character mode. The number itself is counted by characters (general_cost).
        total_units = width if units is None else units
        digit = measure("0")
        rest, _ = _join([p for p in pieces if p[0] != "general"], None)
        # rounded: 8.17 units less ' units' (25 / 6) is 4.0, not 3.999..., and fits '3.14'
        budget = round(total_units - measure(rest) / digit, 9)
        if general_format:
            # the General format keeps one character in a column narrower than one unit: 1 and
            # 0.5 show '1' in a 1-point column, 12 and -1 are '#' (column_narrow)
            budget = max(budget, 1.0)
        cost = general_cost(measure("-") / digit)
        if budget < 1:
            # not one character left: a value from 0.0001 to under 10 shows zero ('( General )'
            # with -1.5 or -0.05 at 2.67 units is '( 0 )'), any other nothing (-1.5E-07 '###')
            a = abs(general_value)
            general = "0" if a == 0 or GENERAL_SMALL <= a < 10 else None
        else:
            general = general_text_fitting(general_value, budget, cost)
        if general is None:
            # nothing fits by characters: the shortest form shows if its glyphs fit the column
            # (1E+06 at 4.67 units, -1E+05x at 6), else the '#' run (general_points suites);
            # the General format keeps the integer on a tie, a General section the scientific
            general = general_fallback(general_value, prefer_integer=general_format)
            text, _ = _join(pieces, general)
            if measure(text) > width:
                return overflow(color, width, full_text, measure, general=True)
        if rest and general.strip("-") == "0":
            # among literals a negative whose General reduced to zero drops its sign, and a '-'
            # literal just before the General: General"x" with -0.001 at width 6 shows "0x"
            # (numeric_core_width), General;-General with -0.05 at width 1 shows "0"
            # (general_literals_points); bare General keeps "-0" (width_general)
            general = "0"
            k = next(i for i, p in enumerate(pieces) if p[0] == "general")
            pieces = [
                p
                for i, p in enumerate(pieces)
                if p[0] != "sign" and not (i == k - 1 and p == ("lit", "-"))
            ]
    text, fill = _join(pieces, general)
    if fill is not None:
        count = int((width - measure(text)) // measure(fill[1]))
        if count > 0:
            position, char = fill
            text = text[:position] + char * count + text[position:]
            if len(text) > MAX_TEXT:
                # more than 255 characters with the fill: a '#' run, text too (limits)
                return overflow(color, width, full_text, measure, general=has_general)
        else:
            fill = None
    # General among literals too must fit by glyph width ('0 units' at width 1 is '#'); the
    # General format alone fits by characters ('1' shows in a 1-point column)
    if overflows and not (has_general and general_format) and measure(text) > width:
        return overflow(color, width, full_text, measure, general=has_general)
    return FormattedCell(text, full_text, color, align, fill, False)
