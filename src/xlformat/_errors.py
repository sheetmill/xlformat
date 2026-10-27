"""Exceptions raised by xlformat."""


class FormatError(ValueError):
    """The format string is one Excel refuses to open from a file.

    Raised only for strings Excel itself cannot render, where the file would need repair, such
    as a lower-case ``e+`` after digits or five sections. Strings that Excel's Format Cells
    dialog rejects but a file may contain render as Excel renders them. Callers that read files
    typically catch it and fall back to ``General``, which is what Excel's repair does to the
    cell. It is a subclass of ``ValueError``.
    """
