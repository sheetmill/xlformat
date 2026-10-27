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

from xlformat._measure import Measure, character_measure, general_cost
from xlformat._number import general_text_fitting
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


def _hashes(width: float | None, measure: Measure) -> str:
    if width is None:
        return "#" * OVERFLOW_WIDTH
    return "#" * max(int(width // measure("#")), 1)


def overflow(
    color: str | None,
    width: float | None = None,
    full_text: str | None = None,
    measure: Measure = character_measure,
) -> FormattedCell:
    """A value the section cannot show: a ``#`` run filling the width (255 without a width)."""
    text = _hashes(width, measure)
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
) -> FormattedCell:
    """Join pieces and apply the width decision.

    ``general_value`` is the number behind a ``general`` piece, for digit reduction, which
    Excel does by characters against the column width in ``units`` (``general_cost``), not by
    glyph widths; ``overflows`` is False for text, which never becomes a ``#`` run.
    """
    full_text, fill = _join(pieces, None)
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
        unsigned, _ = _join([p for p in pieces if p[0] not in ("general", "sign")], None)
        budget = total_units - measure(rest) / digit
        sci_budget = total_units - measure(unsigned) / digit if rest else None
        general = general_text_fitting(general_value, budget, general_cost, sci_budget)
        if general is None:
            return overflow(color, width, full_text, measure)
        if general.strip("-") == "0":
            # a negative whose General reduced to zero drops its sign: General"x" with -0.001
            # at width 6 shows "0x" (numeric_core_width)
            pieces = [p for p in pieces if p[0] != "sign"]
    text, fill = _join(pieces, general)
    if fill is not None:
        count = int((width - measure(text)) // measure(fill[1]))
        if count > 0:
            position, char = fill
            text = text[:position] + char * count + text[position:]
        else:
            fill = None
    if overflows and not has_general and measure(text) > width:
        return overflow(color, width, full_text, measure)
    return FormattedCell(text, full_text, color, align, fill, False)
