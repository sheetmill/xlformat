"""Where a run of decimals after a point counts as fractions of a second.

Usage: ``uv run python -m capture.elapsed_forms``. Runs ``capture/invalid.py``'s check over the
forms below. Results are recorded in ``CAPTURING.md`` section 2 item 42. ``[s]*-.00`` is left out:
a batch stopped on a workbook with that format (item 42).
"""

from __future__ import annotations

FORMS = (
    "h.00", "hh.0", "m.00", "mm.00", "d.00", "yy.0", "yyyy.00", "e.0", "b.0", "g.0", "h:mm.00",
    "mm:ss.00", "s .00", "ss .0", "s x.00", "[s] x.00", '[s]".".00', "[s]\\.00", "[s]_).00", ".00",
    ".00[s]", ".0 [s]", ".00 s", ".0 yyyy", "[s].00.0", "[s].0 .0", "s.00.0", "s.0 .0", "ss.0 yyyy",
    "yyyy ss.0", "AM/PM.0", "h AM/PM.00", "h.00 AM/PM", "m/d/yy.00", "d h.0", "[h]h.0", "[s]s.0",
    "[s].00 [s]",
)  # fmt: skip

REFUSED = (
    "[s]\\.00",
    ".00[s]",
    ".0 [s]",
    ".00 s",
    ".0 yyyy",
    "[s].00 [s]",
    "[s].0000",
    "[s].#",
    "[s].?",
    "[s].0#",
    "[s],00",
    "[h]:[s].00",
    "[s]0",
    "[s] 0",
    "[s]E+00",
    "[s].00%",
    "[s].00 AM/PM",
    "[s]/2",
)
"""The forms Excel refuses to open from a file: ``FormatError``."""

ACCEPTED = (
    *(f for f in FORMS if f not in REFUSED),
    "[s].",
    "[s].0",
    "[s].00",
    "[s].000",
    "[ss].00",
    "[sss].000",
    "[S].00",
    "[m].00",
    "[mm].0",
    "[h].00",
    "[hh].0",
    "[s] .00",
    "[s].00 x",
    "x[s].00",
    "[m]:ss.00",
    "[mm]:ss.000",
    "[h]:mm:ss.00",
    "[h]:mm.00",
    "[s].00;[s].0",
    "[Red][s].00",
    "[s].00*-",
    "[s].00 d",
    "m.0 d",
    "m .0",
    "m x.0",
    "d m.0",
    "m.0 s",
)
"""The forms of the ``elapsed_decimals`` suite."""

NOT_MODELED = ("d [s].00", "yyyy [s].00")
"""An elapsed code after a date code: Excel shows numbers no rule here explains, ``0 4363200.50``
for half a second under ``d [s].00`` (``QUIRKS.md``). Left out of the suite."""

if __name__ == "__main__":
    from capture.invalid import check

    check({form: form for form in FORMS})
