"""Excel's built-in number formats: the format string each ``numFmtId`` below 164 stands for.

A cell can name its number format by id alone; Excel writes id 14 for every date typed as a short
date. The strings are what Excel for Mac reports for each id in en-US (CAPTURING.md section 2 item
49), which differ from other tables at 14 (``m/d/yy``), 44 and 47.
"""

from __future__ import annotations

from collections.abc import Mapping
from types import MappingProxyType

_EXPLICIT: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: "$#,##0_);($#,##0)",
    6: "$#,##0_);[Red]($#,##0)",
    7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    23: "General",
    24: "General",
    25: "General",
    26: "General",
    27: "m/d/yy",
    28: "m/d/yy",
    29: "m/d/yy",
    30: "m/d/yy",
    31: "m/d/yy",
    32: "h:mm:ss",
    33: "h:mm:ss",
    34: "h:mm:ss",
    35: "h:mm:ss",
    36: "m/d/yy",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mm:ss.0",
    48: "##0.0E+0",
    49: "@",
    50: "m/d/yy",
    51: "m/d/yy",
    52: "m/d/yy",
    53: "m/d/yy",
    54: "m/d/yy",
    55: "m/d/yy",
    56: "m/d/yy",
    57: "m/d/yy",
    58: "m/d/yy",
    59: "0",
    60: "0.00",
    61: "#,##0",
    62: "#,##0.00",
    63: "$#,##0_);($#,##0)",
    64: "$#,##0_);[Red]($#,##0)",
    65: "$#,##0.00_);($#,##0.00)",
    66: "$#,##0.00_);[Red]($#,##0.00)",
    67: "0%",
    68: "0.00%",
    69: "# ?/?",
    70: "# ??/??",
    71: "m/d/yy",
    72: "m/d/yy",
    73: "d-mmm-yy",
    74: "d-mmm",
    75: "mmm-yy",
    76: "h:mm",
    77: "h:mm:ss",
    78: "m/d/yy h:mm",
    79: "mm:ss",
    80: "[h]:mm:ss",
    81: "mm:ss.0",
}
CUSTOM_FORMAT_START = 164
"""The first id a file defines itself (``<numFmt>``); ids below it are built in."""

BUILTIN_FORMATS: Mapping[int, str] = MappingProxyType(
    {i: _EXPLICIT.get(i, "General") for i in range(CUSTOM_FORMAT_START)}
)
"""Every built-in id, 0 to 163, and the format string Excel for Mac uses for it in en-US."""


def builtin_format(number_format_id: int) -> str:
    """The format string Excel uses for the built-in ``numFmtId`` ``number_format_id``.

    The table is Excel for Mac's in en-US (CAPTURING.md section 2 item 49). What a cell shows under
    some ids depends on the viewer's system: Excel for Mac follows a custom short-date format set in
    macOS for ids 14 and 22, and pads the day of 15 and 16 to match, while the region alone changes
    nothing on the Mac. For Excel on Windows, Microsoft's "Format numbers as dates or times"
    says that "Date and time formats that begin with an asterisk (*) respond to changes in
    regional date and time settings that are specified in Control Panel"
    (https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309).
    Later keyword-only arguments may choose those forms; this call keeps returning the en-US
    table's string.

    Raises:
        ValueError: ``number_format_id`` is negative or a custom id (164 or more), whose string the
            file itself holds.
    """
    if not 0 <= number_format_id < CUSTOM_FORMAT_START:
        msg = (
            f"{number_format_id!r} is not a built-in number format id (0 to "
            f"{CUSTOM_FORMAT_START - 1}); custom ids are defined in the file"
        )
        raise ValueError(msg)
    return BUILTIN_FORMATS[number_format_id]
