"""Excel's error values (``#DIV/0!``, ``#N/A``, ...) as a value type. See ``ARCHITECTURE.md`` §11.

The members are every error text current Excel documents. The seven classic codes are the
ones ECMA-376 lists and ERROR.TYPE numbers 1 to 7; ``#GETTING_DATA`` is ERROR.TYPE 8; the rest
are the error cell values of the Office JavaScript API (``Excel.ErrorCellValueType``, ExcelApi
1.16: ``#BLOCKED!``, ``#BUSY!``, ``#CALC!``, ``#CONNECT!``, ``#EXTERNAL!``, ``#FIELD!``,
``#SPILL!``; ``#PYTHON!`` and ``#TIMEOUT!`` in preview), and ``#UNKNOWN!``, a data type the
running version does not support (Microsoft support, "#UNKNOWN! error").
"""

from __future__ import annotations

import re
from enum import StrEnum

_CODE = re.compile(r"#[A-Z0-9_/]+[!?]?")


class ExcelError(StrEnum):
    """An error value, which a cell shows as its text whatever its number format.

    ``ExcelError("#DIV/0!")`` is :attr:`DIV0`. A text shaped like an error code that this
    version does not list (``#`` and capitals, digits, ``_`` or ``/``, optionally ending in
    ``!`` or ``?``) gives an unlisted member that renders like any other error, so a code from a
    newer Excel still renders; other text raises ``ValueError``. A reader of a file passes
    ``ExcelError(cell.value)`` for a cell whose type is ``e``.

    Compare errors with ``==``, not ``is``: each lookup of an unlisted code makes a new object.
    Two of them with the same text are equal, hash equal and are the same dictionary key, and
    every error equals its text (``ExcelError.NA == "#N/A"``).
    """

    NULL = "#NULL!"
    DIV0 = "#DIV/0!"
    VALUE = "#VALUE!"
    REF = "#REF!"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NA = "#N/A"
    GETTING_DATA = "#GETTING_DATA"
    SPILL = "#SPILL!"
    CALC = "#CALC!"
    FIELD = "#FIELD!"
    BLOCKED = "#BLOCKED!"
    CONNECT = "#CONNECT!"
    BUSY = "#BUSY!"
    EXTERNAL = "#EXTERNAL!"
    UNKNOWN = "#UNKNOWN!"
    PYTHON = "#PYTHON!"
    TIMEOUT = "#TIMEOUT!"

    @classmethod
    def _missing_(cls, value: object) -> ExcelError | None:
        if not isinstance(value, str) or not _CODE.fullmatch(value):
            return None
        member = str.__new__(cls, value)
        member._name_ = value
        member._value_ = value
        return member
