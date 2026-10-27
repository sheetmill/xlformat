"""Rendering of text and General sections (literals around ``@`` or ``General``)."""

from __future__ import annotations

from decimal import Decimal

from xlformat._number import general_text
from xlformat._parser import Piece, Section, TokenKind


def render_text_section(section: Section, text: str | None, number: Decimal | None) -> list[Piece]:
    """Pieces for a TEXT or GENERAL section. One of ``text`` or ``number`` is given."""
    out: list[Piece] = []
    for t in section.tokens:
        if t.kind is TokenKind.AT:
            out.append(("lit", text or ""))
        elif t.kind is TokenKind.GENERAL:
            if number is not None:
                out.append(("general", general_text(number)))
            else:
                out.append(("lit", text or ""))
        elif t.kind is TokenKind.PAD:
            out.append(("pad", t.text))
        elif t.kind is TokenKind.FILL:
            out.append(("fill", t.text))
        elif t.kind is TokenKind.CURRENCY:
            if t.text:
                out.append(("lit", t.text))
        elif t.kind in (TokenKind.COLOR, TokenKind.CONDITION):
            continue
        else:
            out.append(("lit", t.text))
    return out
