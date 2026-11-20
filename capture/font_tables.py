"""The few numbers of each face's font file that place its rules (underline, strikethrough).

    uv run python -m capture.font_tables      # print them, one line per face

Reads the TrueType files Excel for Mac draws with (its bundled fonts, else macOS's) with the
standard library: ``head.unitsPerEm``, ``post.underlinePosition`` and ``underlineThickness``,
``OS/2.yStrikeoutPosition``, ``yStrikeoutSize`` and ``usWinDescent``. The package stores these
numbers, not the files (``xlformat._measure._RULE_TABLES``).
"""

from __future__ import annotations

import struct
import sys
from pathlib import Path

BUNDLED = Path("/Applications/Microsoft Excel.app/Contents/Resources/DFonts")
SYSTEM = Path("/System/Library/Fonts/Supplemental")
FILES = {
    "Calibri": BUNDLED / "Calibri.ttf",
    "Aptos Narrow": BUNDLED / "Aptos-Narrow.ttf",
    "Arial": BUNDLED / "arial.ttf",
    "Times New Roman": BUNDLED / "times.ttf",
    "Courier New": SYSTEM / "Courier New.ttf",
    "Georgia": SYSTEM / "Georgia.ttf",
}


def tables(path: Path) -> dict[str, int]:
    data = path.read_bytes()
    count = struct.unpack(">H", data[4:6])[0]
    offsets = {}
    for i in range(count):
        tag, _, offset, _ = struct.unpack(">4sIII", data[12 + 16 * i : 28 + 16 * i])
        offsets[tag.decode("latin-1")] = offset
    head, post, os2 = offsets["head"], offsets["post"], offsets["OS/2"]
    position, thickness = struct.unpack(">hh", data[post + 8 : post + 12])
    strike_size, strike_position = struct.unpack(">hh", data[os2 + 26 : os2 + 30])
    win_descent = struct.unpack(">H", data[os2 + 76 : os2 + 78])[0]
    return {
        "upem": struct.unpack(">H", data[head + 18 : head + 20])[0],
        "underline_position": position,
        "underline_thickness": thickness,
        "strikeout_position": strike_position,
        "strikeout_size": strike_size,
        "win_descent": win_descent,
    }


def main(argv: list[str]) -> None:
    """Command-line entry point."""
    if argv:
        sys.exit(__doc__)
    for face, path in FILES.items():
        print(face, tables(path))


if __name__ == "__main__":
    main(sys.argv[1:])
