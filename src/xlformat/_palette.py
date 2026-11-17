"""Excel's default color palette: section color names to ``RRGGBB``.

The eight names (``[Red]`` ...) and ``[Color1]`` to ``[Color56]`` are the legacy 56-color
palette of the workbook. Several indices share one color (``Color25`` and ``Color11`` are both
navy), so callers compare colors, not names. Evidence: the ``colors`` capture suite reads the
font color of each cell's display format back from Excel (``CAPTURING.md`` section 2 item 21).
"""

from __future__ import annotations

from collections.abc import Sequence
from types import MappingProxyType

_INDEXED: tuple[str, ...] = (
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
)  # fmt: skip

DEFAULT_PALETTE = MappingProxyType(
    {
        "Black": "000000",
        "White": "FFFFFF",
        "Red": "FF0000",
        "Green": "00FF00",
        "Blue": "0000FF",
        "Yellow": "FFFF00",
        "Magenta": "FF00FF",
        "Cyan": "00FFFF",
        **{f"Color{i}": rgb for i, rgb in enumerate(_INDEXED, 1)},
    }
)
"""Canonical section color name (``FormattedCell.color``) to ``RRGGBB`` in the default palette."""

_NAMES = ("Black", "White", "Red", "Green", "Blue", "Yellow", "Magenta", "Cyan")
_HEX = frozenset("0123456789ABCDEF")


def _entry(value: str) -> str:
    """``RRGGBB`` from an ``RRGGBB`` or ``AARRGGBB`` palette entry."""
    text = value.strip().upper()
    if len(text) == 8:
        text = text[2:]
    if len(text) != 6 or not set(text) <= _HEX:
        msg = f"not an RRGGBB or AARRGGBB color: {value!r}"
        raise ValueError(msg)
    return text


def section_rgb(color: str, palette: Sequence[str] | None = None) -> str:
    """The ``RRGGBB`` Excel draws a section color in (CAPTURING.md section 2 item 50).

    Args:
        color: A canonical section color name, as ``FormattedCell.color`` holds it.
        palette: The workbook's ``<indexedColors>`` as stored, 64 entries of ``RRGGBB`` or
            ``AARRGGBB``; ``None`` for the default palette.

    The eight names (``Black`` to ``Cyan``) are drawn in their standard colors whatever the
    palette holds; ``ColorN`` is the palette's entry N + 7. For an index a shorter palette lacks
    the default palette's color is used: a choice of this package, as what Excel does with a
    short palette is not captured.

    Raises:
        KeyError: ``color`` is not a section color name.
        ValueError: The palette entry used is not a hex color.
    """
    if color in _NAMES:
        return DEFAULT_PALETTE[color]
    default = DEFAULT_PALETTE[color]  # raises KeyError for an unknown name
    if palette is None:
        return default
    index = int(color[5:]) + 7
    if index >= len(palette):
        return default
    return _entry(palette[index])
