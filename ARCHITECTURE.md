# Architecture

Companions: `CAPTURING.md` (test data) and `QUIRKS.md` (Excel evidence per quirk). Callers with
error cells: see section 5.

## 1. Scope and principles

`xlformat` turns a value and an Excel number-format string into the text Excel displays.
Pure Python, standard library only, Python 3.11+. Priorities in order: match Excel 16 for Mac,
then speed, then code quality. Excel's behavior is the specification; ECMA-376 section 18.8.31
is a secondary reference and is wrong or silent in places where Excel is definite.

Out of scope for 0.1: locales other than en-US (format codes are parsed, names are English,
`[$-x-sysdate]` renders the en-US long date), fonts other than Calibri in Excel mode, conditional formatting,
right-to-left and East Asian calendar codes (`e`, `g`, `b`), cell styles beyond what the format
string itself implies.

## 2. Pipeline

```
format string ──parse──▶ NumberFormat (cached)
                            │  sections: tuple[Section, ...]
value ──normalize──▶ Value  │
        (kind, serial/text)─┴─select section──▶ pieces ──assemble(width, measure)──▶ str | FormattedCell
```

1. **Parse** (`_parser.py`): split into at most four sections on `;` outside quotes and escapes;
   tokenize each section; classify it; compile it into a `Section` dataclass. Parsing is pure,
   raises `FormatError` only when Excel itself has no rendering for the string (section 11), and
   is cached (`functools.lru_cache`, size 1024) keyed on the string alone. Quirks bind at format
   time, not parse time: none of the three quirks affects parsing.
2. **Normalize the value** (`_values.py`): every input becomes `Value(kind, number, text)` where
   `kind` is `NUMBER`, `TEXT`, `BOOL`, `EMPTY`; numbers are `Decimal`. See section 5.
3. **Select the section** (section 4).
4. **Render** (`_number.py`, `_datetime.py`, `_text.py`): the section renderer returns a list of
   pieces: literal text, digit text, and markers (`Fill(char)`, `Pad(char)`).
5. **Assemble** (`_assemble.py`): compute the unconstrained text, then apply the width decision
   (overflow, General reduction, fill, pad) through a measure function (section 8), produce the
   string and, for `format_cell()`, the `FormattedCell` record with color and alignment.

## 3. Format model

Tokens per section, produced by one regex pass (the old tokenizer approach, kept):

| token | notes |
|---|---|
| `"..."`, `\x`, and the always-literal characters `$ - + ( ) : space ! ^ & ~ { } < > =` | literals; `\` and `"` are consumed |
| `[Red]` ... `[Color56]` | section color; case-insensitive; `[Color 3]` with a space is accepted and kept as `[Color3]` (`CAPTURING.md` §2 item 9) |
| `[<=100]`, `[<>0]`, `[=1]` | condition, sections 1 and 2 only; a condition in the text section is an error |
| `[$€-x-euro2]`, `[$-409]`, `[$-F800]` | currency symbol (emitted as a literal) and locale tag (recorded, not applied) |
| `[h]`, `[mm]`, `[ss]` and longer runs | elapsed time; at most one per section; nothing may precede it except literals |
| `General` (case-insensitive), `@` | General placeholder (allowed in numeric and text sections, with literals around it); text placeholder |
| `0 # ?`, `.`, `,`, `%`, `E+ E-`, `/` | numeric codes. Lower-case `e+`/`e-` after digits make Excel refuse the file (`FormatError`, `CAPTURING.md` §2 item 9). A second `.` is a literal inside the decimal digits (`0.0.0` shows `1234.5.7`). `%` between mantissa and `E` is a literal and does not scale (`0%E+0` shows `1%E+3`) |
| `y m d h s e`, `AM/PM`, `A/P` runs | date/time codes, case-insensitive, normalized (`yyy`→`yyyy`, `mmmmmm`→`mmmm`, `ddddd`→`dddd`, `sss`→`ss`). `e` is a four-digit year in en-US (`e+` shows `1903+` for serial 1234) |
| `_x`, `*x` | pad (width of `x`, one space in text output) and fill (repeat `x`). Excel accepts two fills in a file and honors the last one (`*-0*-` shows `1235` then dashes); the dialog refuses it |
| anything else | literal |

Classification of a section, in this order: contains a date/time code → **datetime** (numeric
codes are then errors, except the `0` runs of `s.0`, `s.00`, `s.000`); contains `/` between digit
runs → **fraction**; contains `E+`/`E-` after a digit run → **scientific**; contains `0 # ? .` →
**number**; contains `@` → **text**; only `General` and literals → **general**; nothing → **empty**.

Compiled `Section` fields (frozen dataclass): `kind`, `color`, `condition` (op, Decimal),
`locale`, `pieces` (literals and markers), and per kind: integer placeholders with their
inter-digit literals, decimal placeholders, grouping flag, scale (10^-3 per trailing comma),
percent count, exponent placeholders and sign style, fraction numerator/denominator placeholders
or fixed denominator, resolved date fields (`m` already resolved to month or minute, `h` already
knows whether AM/PM is present, fractional-second digit count).

`m` resolution (from the 2023 prototype, verified by the Excel captures): scanning left to right, an `m` run
is a minute if an `h` run precedes it since the previous `m` run; otherwise a minute if the next
date/time code is a second; otherwise a month. The general rule is that `m` next to `h` or `s` is
a minute.

Comma rules (from the 2023 prototype's docstring, which the Excel captures confirm): commas between digit
placeholders before the decimal point turn on grouping; commas immediately after the last digit
placeholder scale by a thousand each; commas before any placeholder, or after a literal, are
literals; commas after the decimal point are dropped.

## 4. Section selection

Given the normalized value and the compiled sections (n = 1..4):

- **Text and booleans** use the text section if there is one (the fourth, or the last section if
  it contains `@`); otherwise the text is displayed as is (General). Booleans render as `TRUE` /
  `FALSE` through that path, so `_(@_)` gives `' TRUE '`.
- **Empty** (`None`) renders as `''`.
- **Numbers**, conditions present (`_render.py: select`; suites `numeric_core`,
  `conditions_rounding`, `lone_zero`, `selection_fallback`): the condition sees the unrounded
  value (`[=0]"zero";0` shows `0` for 0.001). A conditioned first section applies when it holds
  and keeps the sign (`-1` under `[<=9999999]###-####;...` gives `--1`) unless the condition can
  only hold for negatives (`<` with bound ≤ 0, `<=` with bound < 0), which drops it like the
  implicit negative section (`[<0]0.0` shows `1.0` for -1). When it fails: with one section,
  General (sign dropped when the bound is negative; a `General` body shows `#`); with two, an
  unconditioned second section is "everything else" with the sign kept (`[>100]0.00;0` shows
  `-1`) unless the failing range holds only non-positive values (`[>0]0;"neg"` shows `neg`), and
  a conditioned second section applies when it holds or else the value shows a `#` run
  (`[<=0]0;[>1]0` with 0.3); with three, an unconditioned second section is the negative one
  (`[>1]0;0;0.0` shows `2` for -2) and the third takes the rest (`0.3`, `0.0`), while a failing
  conditioned second section sends the value to the third with its sign. An unconditioned first
  section before a conditioned second takes positives only: `0;[>1]0` shows `#` for 0 and
  negatives, `0;[>1]0;0.0` shows `-0.3` and `0.0`. A negative rendered with its sign whose
  magnitude renders as zero is rendered again as the positive magnitude, from section selection
  onward (`[<=0]0` shows `0.001` for -0.001, `[<=0.1]0` shows `0.3` for -0.3, `[<=0.5]0` shows
  `0`; `[<0]0` shows `0` because its sign was already dropped).
- **Numbers**, no conditions: 1 section → all values, with an automatic leading `-` for
  negatives placed before the whole section text (`x( General )` gives `-x( 1234 )`); 2 sections
  → `>= 0` and `< 0`, the negative section renders the absolute value and adds no sign; 3 or 4 →
  `> 0`, `< 0`, `= 0`. The sign test uses the unrounded value: `-0.4` under `0;(0)` gives `(0)`,
  while under the single section `0` it gives `0` (a negative that rounds to zero drops the sign).
- An empty section renders nothing (this is how `0;;` hides negatives), except that an empty first
  section, or an empty section carrying a color or condition, behaves as General.
- Values a datetime section cannot show (negative, or `>= 2958466` after rounding to the displayed
  precision) render as overflow (section 8).

## 5. Value model

Accepted inputs and their normalization:

| input | becomes |
|---|---|
| `bool` (checked first) | `BOOL`; text `TRUE`/`FALSE` |
| `int`, `float`, `Decimal`, objects with `__index__` | `NUMBER`, `Decimal`; floats via exact `Decimal(float)` (NaN and infinities raise `ValueError`) |
| `str` | `TEXT` |
| `None` | `EMPTY` |
| `datetime.datetime`, `date` | `NUMBER`: serial under the selected date system (the Lotus shift applies to 1900-01-01..1900-02-28 so that round-tripping agrees with Excel) |
| `datetime.time` | `NUMBER`: fraction of a day |
| `datetime.timedelta` | `NUMBER`: days as a fraction (for elapsed formats) |
| anything else | `TypeError` |

**Decimal, not float**, is the working type: Excel's display is a decimal process (15
significant digits, half-up rounding at the displayed place) applied to a binary double, and
`Decimal(float)` reproduces the double exactly so the decimal process can be applied faithfully.
`float` arithmetic would reintroduce the errors the 2023 prototype spent effort avoiding. Performance
cost is acceptable (a few microseconds per value) and is measured, not assumed.

Precision rules applied at normalization, each a quirk (section 9):
- `fifteen_digit_precision`: round the magnitude to 15 significant digits (tie rule below).
  Evidence: `0.1234512345123456` under thirty decimals shows `0.123451234512346000...`.
- The tie rule (`CAPTURING.md` §2 items 8, 15 and 16; `QUIRKS.md`): the exact binary value is
  rounded to 15 significant digits with **ties down** (`12345123451234.25` under `0.0` shows
  `….2`, `123451234512345.5` under `0` shows `…345`, `….515625` shows `…346`). Zero failures on
  69,690 numeric captured rows. An earlier extra step at or above 1e14 ("round to 16 digits
  first") fitted the 16-digit number texts openpyxl had written, not Excel; the formula-written
  rows removed it. The 2023 prototype's "exact .5 rounds down at 15 integer digits" was an
  approximation of the same artifact.

**Error values.** `#N/A`, `#DIV/0!` and the other error values are not inputs in 0.1. Excel
displays them regardless of number format, and openpyxl exposes them as strings with
`cell.data_type == 'e'`, so a renderer bypasses the formatter for those cells. A string such as
`'#N/A'` passed in is just text.

Fraction formats work from the exact double, not the 15-digit value, and run the denominator
search and the fixed-denominator rounding in double arithmetic, which is what reproduces Excel.

## 6. Numbers

- **Rounding**: `Decimal.quantize` with `ROUND_HALF_UP` at the number of decimal placeholders
  after percent and comma scaling.
- **Digit placement**: from the quantized digit tuple, right to left for the integer part (`0`
  pads with zeros, `?` with spaces, `#` with nothing; grouping inserts `,` every three digits
  when on; inter-digit literals are emitted where they sit), left to right for the decimal part
  (trailing `#` positions beyond the value's digits are omitted, `?` become spaces).
- **General**: integer values print as integers; otherwise up to 11 characters (10 significant
  digits plus the point), trailing zeros stripped; at or above `99,999,999,999.5` or below
  `0.0001` (with more than nine decimals needed) switch to `d.dddddE+ee` with up to five
  mantissa decimals and a two-digit exponent. These thresholds come from the 2023 prototype and the
  Excel captures; the width-dependent form is in section 8.
- **Scientific**: with `n` integer placeholders the mantissa has `(exponent mod n) + 1` integer
  digits (so `##0.0E+0` is engineering notation); zero renders every integer placeholder as `0`
  (`000.0E+0`, but `#.#E+#` shows `0.E+0`); the exponent is padded to its placeholder count,
  shows at least one digit, and is signed per `E+`/`E-`; `%` is a literal and does not scale.
  1,770 rows of 2023 and the 960-row `scientific` suite pass.
- **Fraction**: whole part (optional) from the integer digits, numerator and denominator digit
  counts from the placeholders, fixed denominator when digits follow `/`. Fractions work from
  the exact double (a 15-digit rounding step changed nothing and was removed); the
  fixed-denominator numerator is Excel's nudged half-up in double arithmetic: with `f` the
  fraction part, `f' = fl(f · (1 + 2⁻⁵²))`, `t = fl(f' · D)`, numerator `floor(t)` plus one when
  `t − floor(t) ≥ 0.5` (`CAPTURING.md` §2 item 22). Free denominators use a continued-fraction search bounded by
  `10^digits - 1` with the epsilon from ExcelNumberFormat, which the Excel captures confirm for the
  2023 cases; `fraction_denominators` off uses `Fraction.limit_denominator`. Zero fraction: `?`
  numerators show spaces (the bar too), `#` show nothing, `0` show `0/1`; a whole value of zero
  shows `0` even under `#`; `%` scales. A decimal point before the bar is a `FormatError`
  (Excel refuses the file). All 5,310 rows of 2023 pass.

**Fixed-denominator numerator** (`fraction_ties`, `fraction_threshold`,
`fraction_random`; `CAPTURING.md` §2 item 22). An
empirical tolerance of `5e-18 · D²` had missed 32 rows near half steps; the nudge by one relative
epsilon on the fraction part, before the product, fits all 38,904 fixed-denominator rows with no
known deviation left. `strict` rounds the exact product half-up, the closest fraction.

## 7. Dates and times

In `_datetime.py`; every rule below is pinned by the 2023 suites
(35,990 date/time rows), `datetime` (3,630 rows) and `date1904` (336 rows).

**Excel's clock arithmetic, operation for operation** (`CAPTURING.md` §2 item 15; fitted to the
formula-written rows, zero misses on 41,266 date and time rows). The serial `x` is the double
Excel stores, not the 15-digit display value (`29.5` seconds shows `30`, and 23:59:59.5 on the
last day overflows). With `scale` = 1, 10, 100 or 1000 for the finest displayed unit (a second,
or a tenth, hundredth or thousandth; never coarser):

```
d = x + (0.5 / scale) / 86400          bias of half the displayed unit, in days
day = floor(d);        h = (d - day) * 24
hour = floor(h);       m = (h - hour) * 60
minute = floor(m);     s = (m - minute) * 60
second = floor(s);     fraction = floor((s - second) * scale)
elapsed [h], [m], [s] = floor((x / (1/86400) + 0.5 / scale) / unit)
```

Every field is truncated, so the bias rounds half-up and the carry propagates into the day:
23:59:59.999 under `h:mm` is `0:00` the next day, under `ss.000` it is `59.999`. The
floating-point noise of the peel is Excel's: a `General` token inside a date section shows `s`
before truncation (`0.5` for a whole day, `0.499999803` for serial 44927), which is how the
order of operations was pinned (nine digits per row over 66 serials). The clock and elapsed
paths differ in the last bit and Excel shows it: 59.5 seconds under `[mm]:ss` is `01:59`, and
44927 + 59.5 s shows `64694881:59`. The 59.5-second rows separate `x / (1/86400)` from
`x * 86400` (59.5 against 59.49999999999999); the day-44927 rows separate the bias in days from a
bias in seconds.

- **1900 system** (default): day `d >= 61` → civil date `1899-12-30 + d`; `1 <= d <= 59` →
  `1899-12-31 + d`; `d == 60` → February 29, 1900; `d == 0` → "January 0, 1900". The weekday is
  `(d + 6) mod 7` with Sunday = 0 (serial 1 is a Sunday; it was a Monday). With `leap_year_1900`
  off the mapping is `1899-12-30 + d` for all `d` (LibreOffice's model) and weekdays are true.
- **1904 system**: day `d` → `1904-01-01 + d`; weekday `(d + 5) mod 7`; no leap quirk. A
  negative value renders as `-` followed by the rendering of its magnitude (`-1` under
  `m/d/yyyy` is `-1/2/1904`, `-0.5` under `h:mm` is `-12:00`).
- **Negatives in 1900**: a section made only of elapsed codes shows `-` and the magnitude
  (`[h]` with -1 is `-24`) and has no upper limit; any other section shows the `#` run, decided
  on the raw sign (-1e-9 is `#`, not January 0).
- **Limits**: a clock section whose rounded value reaches day 2958466 (year 10000) shows `#`;
  under 1904 the limit is 1462 days lower. 23:59:59.5 on 9999-12-31 overflows under `ss` and
  shows `59.5` under `ss.0`.
- Civil conversion is integer arithmetic (days-from-civil and back, proleptic Gregorian), not
  `datetime`, so year 0 to 9999 and the two synthetic 1900 dates need no special objects.
- `h` shows 1..12 when the section contains `AM/PM` or `A/P`. `AM/PM` always renders upper-case
  whatever its spelling; `A/P` keeps the token's case (`a/p` gives `a`/`p`).
- `e` is a four-digit year code. `[$-x-sysdate]` and `[$-F800]` render `dddd, mmmm d, yyyy`
  (the system long date in en-US; Excel's canned string says `dd`, the display says `d`);
  `[$-x-systime]` and `[$-F400]` render `h:mm:ss AM/PM`; a bare tag is the whole format.
- `General` inside a date section shows Excel's biased seconds before truncation (`s` above):
  a whole day shows `0.5`, 23:59:59.999 shows `0.499`, serial 365 shows `0.499999999`. The
  noise is reproduced because the arithmetic is.
- Names are English; month and weekday names, `mmmmm` (initial) and `ddd` (three letters).

## 8. Width model

Two ways a caller limits width, one mechanism underneath.

**Character mode** (first to ship): `width: int`, a literal character count. Exact for
monospace and terminal output. It is not an approximation of Excel; it is its own contract.

**Excel mode** (second): the caller passes an Excel column width, the number from the Column
Width dialog (8.43 is the default), and gets what Excel displays: the same overflow decision,
the same count of `#`, the same General digit reduction, the same `*` fill count. Excel decides
by glyph widths, not characters: "Tuesday, January 4, 2000" and "Sunday, December 3, 2000"
are both 24 characters, yet at column width 21.17 only the second shows `####` (verified by
hand).
Excel mode therefore needs per-glyph advance tables. Expected outputs come from Excel for Mac, whose text
renderer differs from Windows, so borderline fits are matched for Mac and the documentation
says so.

**Glyph tables** are measured per font and per point size and bundled as data
(`_glyphs_calibri.py`, generated by `capture/glyph_fit.py` from the `measure*` suites). The first
design stored one table per font in em fractions and scaled it at run time; Excel for Mac hints
advances per size (digits are 6 points at both 11 and 12, the hyphen 3 then 4), and no
size-independent fraction reproduces that, so the measured tables replaced it (the Excel-mode
paragraph below). Calibri is the only bundled font: openpyxl writes Calibri 11 as the workbook
default and Excel for Mac's own default is Calibri 12. Aptos Narrow, the default of current
Windows Excel, ships inside Excel for Mac and is measured there the same way (tables planned for
0.2); whether Windows Excel's widths differ from the Mac's is a separate, open question. Reading
installed fonts is an optional extra later, never a requirement.

**The seam**: internally both modes are "a measure function plus a width in its units": `len`
for character mode, a glyph-advance sum for Excel mode. The same seam is exposed for callers
with special needs (a measure callable plus a width). The pipeline is built around it from the
start so that character mode and Excel mode share every line except the measure. The public
spelling is in section 11: `width` for character mode, `column_width` and `font` for Excel mode.

What the width logic does, and why a caller cannot do it afterwards from the plain string:
- `*x` repeats `x` so the text measures exactly the width (never negative); the last `*`
  in a section is honored (`*-0*-` shows `1235` then dashes).
- Numbers and dates whose text measures wider than the width become a run of `#` that fills the
  width; text never overflows (clipping or spilling is the renderer's job). The caller cannot
  apply this rule from outside because it depends on the value kind and on the section, not on
  the string.
- General reduces before it overflows: it drops decimals (rounding) to fit, then tries
  scientific with fewer mantissa digits, then `#`. The exact order is reverse-engineered from the
  `width_general` capture suite before implementation, not assumed.
- `_x` counts one character in character mode and the advance of `x` in Excel mode.

`width=None` is the unconstrained rendering: `*x` emits nothing, `_x` emits one space, General
uses the 11-character rule, and a value a datetime section cannot show gives 255 `#`. That is
Excel's clipboard text; `Range.Text` gives as many `#` as the column holds (30 at width 30), and
255 is the widest column Excel allows, so the two agree at the limit (`CAPTURING.md` §2 items 1
and 2). sheetpeek always passes a width, so the 255 case is for callers without a layout.

Observed in the `width_general` capture (`CAPTURING.md` §2 item 11): General drops decimals one at
a time, rounding, until the text fits (`999.9999999` becomes `1000` as soon as one decimal goes);
when the integer part does not fit it switches to scientific with as many mantissa decimals as
fit, down to `1E+07`; below that, `#`. Values below one reduce to `0` rather than overflow.
Booleans overflow like numbers; text never does. Fill and overflow counts in Excel are
glyph-measured (35 dashes and 53 spaces at width 30), so character mode and Excel mode differ
on them by design.

Consequences for the captures: the character-mode suite uses digit-only output (digits are tabular in
Excel's default fonts, so character and pixel agree); the Excel-mode suite uses everything,
including dates with month and day names. Both are captured with `string value` at set column
widths, and `autofit` followed by reading `column width` back gives Excel's own measurement of
a text, which is the data a glyph table is built or verified from.

**Character mode** (`_measure.py`, `_assemble.py`, `general_text_fitting`
in `_number.py`; suites `width_general`, `width_fill`, `numeric_core_width`; `CAPTURING.md` §2 item
23). `width: int` is a count of characters, `measure` is `len`. The `#` run is `width` characters
long by contract; Excel's count is glyph-measured (six at width 5 in Calibri 11), so in character
mode at a width the capture test treats any two `#` runs as equal (4,907 of 5,396 such rows
differ in count); width-less rows and every Excel-mode row are compared exactly. Rows where Excel fitted a text longer than the width
(commas, points, spaces and minus signs are narrower than digits: `1,000.0` fits width 6,
`-12345678` fits 8.43) or overflowed a text no longer than it (`%` is wider: `10000%` overflows
width 6) are glyph slack and glyph tightness; character mode cannot reproduce them and the test
skips them (2,224 of the 32,832 `numeric_core_width` rows, 11 of `width_fill`, 16 of
`width_general`); they are the evidence Excel mode will be fitted to. Rows at fractional widths
(8.43) have no character-mode meaning and are skipped. Two refinements of the rules above came
from the suites: text does overflow to `#` when its section carries a `*` fill (`@*-` with
`Hello World` at width 6), and a negative whose `General` reduces to zero at the width drops its
sign when the section has literals (`General"x"` shows `0x`) but keeps it when the section is a
bare `General` (`-0.05` at width 3 shows `-0`). `FormattedCell.fill` reports the position in
`text`. Renders with a width measure under 7 µs.

**Excel mode** (`_measure.py`, `_glyphs_calibri.py`; suites `measure11`,
`measure12` and the width suites; `CAPTURING.md` §2 items 24 and 25). Excel for Mac works in whole
points. A column of `w` units is `round(MDW · w + 5)` points, MDW being the maximum digit width of
the Normal font: 5 points for Calibri 9 and 10, 6 for 11 and 12 (so 8.43 snaps to 8.5 units, 56
points), 7 for 14, 8 for 16; text has those points less the 5 of padding: `MDW · w`. A text
measures the floored sum of its glyph advances in points at the cell's size; the advances are
hinted per size (digits are 6 at 11 and 12, `-` is 3 at 11 and 4 at 12, `%` 8 and 9.1), so the
bundled table holds every measured size (9, 10, 11, 12, 14, 16) as measured and scales the
nearest measured table with per-glyph rounding for other sizes; glyphs outside the table measure
as a digit. The autofit width of a text is that measure plus 7 points at every size, which is how
the tables were fitted (`capture/glyph_fit.py`, 515 of 515 space-less texts exact per size). The
space (2, 2, 2, 3, 3, 4 points from 9 to 16) comes from the `space_fill` suites, rendered through
the package with every candidate advance (`CAPTURING.md` §2 item 26); `_x` pads and `?` blanks
measure as spaces. A `#` run holds `floor(budget / 5)` characters at 11 points (six at width 5),
a `*` fill `floor((budget − text) / advance)` repeats (35 dashes at width 30 in Calibri 12). Fits
are `measure ≤ budget`: `1,000.0` (36 points) fits width 6 and `10000%` (38) does not. `General`
alone reduces by characters, not points: one unit per character and half for a leading minus
against the column width in units (`-12345678` fits 8.5, `-1234.6` does not fit 6, `-2E-07` needs
6); among literals the General token gets the units less the literals' and the sign's advances in
digit units (`( 1234 )` shows at width 6, `( 0.125 )` becomes `( 0.13 )`), and the scientific form
of a value whose integer part does not fit is admitted without charging the sign (`-1E+05x` at
width 6) while a value that rounds to zero still pays for it (`0x`). Text overflows only through a
section whose `*` fill follows it (`@*-` shows `#`, `*-@` shows the text). Every width row of
every suite passes in Excel mode, the 2,251 glyph-slack rows included; character mode is
unchanged. Renders in Excel mode measure 3.6 to 8.7 µs. `width_general14` passes
at Calibri 14 with no change to the model.

## 9. Quirks

```python
@dataclass(frozen=True, slots=True)
class Quirks:
    leap_year_1900: bool = True
    fifteen_digit_precision: bool = True   # includes the large-value tie rule, section 5
    fraction_denominators: bool = True

    EXCEL: ClassVar[Quirks]   # all True (the default everywhere)
    STRICT: ClassVar[Quirks]  # all False
```

Names are shared with `xlformula` where a quirk exists in both packages; an early candidate,
`half_down_above_15_digits`, is folded into `fifteen_digit_precision` (section 5).
Quirks bind at format
time: `format_value`, `format_cell` and the `NumberFormat` methods take
`quirks: Quirks | None = None` (None means `Quirks.EXCEL`); `NumberFormat.parse` does not, and
the parse cache is keyed on the string alone. If a parse-affecting quirk ever appears, revisit.
`QUIRKS.md` holds, per quirk, the Excel evidence (captured rows), what `strict` does instead, and
the first version it appeared in. Candidates that need evidence before they are added:
the `[$-x-sysdate]` `dd` versus `d` discrepancy. Scientific zero rendering (`000.0E+0`), the
section-selection sign rules, the last `*` fill winning and `General` inside a date section are
Excel semantics matched unconditionally, not quirks (`QUIRKS.md`).

The full set of Excel captures runs under `EXCEL`; `STRICT` gets unit tests that pin the alternative behavior
on a handful of values.

## 10. Output types: plain versus rich

Design principle: the plain call returns a plain `str`, and a caller who wants more than the
string asks for it explicitly.

What Excel shows beyond the characters: the section color (eight names or `Color1..56`), the
default horizontal alignment (numbers and dates right, text left, booleans centered, derived from
the value kind, not the format), the position of a `*` fill, and whether the cell overflowed.

Options considered, with precedent:

**A. `str` subclass with attributes** (`xlfmt.Formatted`; `markupsafe.Markup`,
`bs4.NavigableString`, lxml's smart strings). One function, zero friction, `isinstance(x, str)`.
Against: attributes vanish silently on `.strip()`, `+`, f-strings and `json`; `type(x) is str`
is false; every caller pays for the subclass whether they want it or not; and it breaks the
principle above, since every caller gets more than a string without asking. The lxml and
BeautifulSoup experience is that string subclasses carrying context are a recurring source of
confusion.

**B. Two entry points: plain function returns `str`, rich function returns a result object**
(`subprocess.check_output` versus `subprocess.run` → `CompletedProcess`; `os.listdir` versus
`os.scandir`; `re.search` → `Match`). Explicit, the plain path stays plain and fast, the record
can grow fields without breaking anyone, and it prints readably. Against: two names to learn.
A variant is a `details=True` flag with `@overload` (as `open()` and `Path.read_text` do by
mode); workable but flag-selected return types are the less idiomatic form.

**C. Query functions on the compiled format** (`numfmt.formatColor(pattern, value)`,
`numfmt.getFormatInfo(pattern)`, `isDateFormat`; ExcelNumberFormat's `NumberFormat` object).
Pattern-level facts (`is_date`, `is_text`, has colors) belong here. Against, for value-level
facts: color needs section selection, so `format_value` plus `color_for(value)` does that work
twice, and fill position is a property of the rendered text, not the pattern.

**Chosen: B plus the pattern-level part of C.** `format_value()`
returns `str`; `format_cell()` returns `FormattedCell`; `NumberFormat` (the compiled object, as
`re.compile` returns a `Pattern`) exposes `is_date`, `is_text`, `is_general`, `sections` and
both methods. Not A. The rich function is not called `render`: in sheetpeek, which imports it
next to its own functions, rendering means producing HTML or terminal output from a sheet.
`format_cell` takes a value and a format string, not an openpyxl cell; the name describes what
comes back, the display of one cell.

```python
@dataclass(frozen=True, slots=True)
class FormattedCell:
    text: str                  # what the cell shows under the given width
    full_text: str             # the width=None rendering; always set, never None
    color: str | None          # 'Red', 'Blue', ..., 'Color12'; canonical spelling
    align: Literal["left", "right", "center"]
    fill: tuple[int, str] | None   # index in text and character of a `*` fill
    overflow: bool             # text is a run of '#'
```

`full_text` is what copy-and-paste gives a user and what the 2023 capture holds; it is computed
before the width decision, so it costs nothing. `overflow` means only "`text` is a `#` run";
General digit reduction and `*` fill also make `text` differ from `full_text` without being
overflow, and a caller who needs to know compares the two fields. `xlformat.DEFAULT_PALETTE` maps the
canonical color names to `RRGGBB` in the default palette (the eight names and `Color1` to
`Color56`, with the palette's duplicates), confirmed value by value by the `colors` suite.

## 11. Public API

```python
from xlformat import format_value, format_cell, NumberFormat, FormattedCell, Quirks, FormatError

def format_value(value, format_string="General", *, width=None, column_width=None, font=None,
                 measure=None, date1904=False, quirks=None) -> str
def format_cell(value, format_string="General", *, width=None, column_width=None, font=None,
                measure=None, date1904=False, quirks=None) -> FormattedCell

class NumberFormat:
    @classmethod
    def parse(cls, format_string: str) -> "NumberFormat"      # cached; raises FormatError
    def format(self, value, *, width=None, column_width=None, font=None, measure=None,
               date1904=False, quirks=None) -> str
    def format_cell(self, value, *, width=None, column_width=None, font=None, measure=None,
                    date1904=False, quirks=None) -> FormattedCell
    is_date: bool; is_text: bool; is_general: bool; sections: tuple[Section, ...]

class FormatError(ValueError): ...
```

Width parameters: `width: int | None` is character mode, a count of
characters; `column_width: float | None` is Excel mode in Column Width dialog units, measured with
the glyph table of `font`; passing both raises `ValueError`. `font: str | tuple[str, float] | None`
names the font and size, default `("Calibri", 11)` (openpyxl's workbook default; sheetpeek passes
the workbook's Normal style); a font without a bundled table raises `LookupError` naming the
bundled fonts (Calibri only, measured at 9, 10, 11, 12, 14 and 16 points and scaled from the
nearest measured size elsewhere; Aptos Narrow tables are planned for 0.2). `measure:
Callable[[str], float] | None` is the hook for special cases and is used with `width` as the budget
in the measure's own units. `DEFAULT_PALETTE` maps color names to `RRGGBB`.

Behavioral contract:
- `FormatError` is raised only when Excel itself has no rendering for the format string **as it
  appears in a file**. "What the Format Cells dialog refuses" is the wrong test for a library
  that mostly reads files: of 21 dialog-rejected strings written with openpyxl, Excel refused to
  open the file for 16 (lower-case `e+` after digits, five sections, unclosed `[` or `"`, two
  colors, two conditions, `@` in two sections or a middle section, date and digit codes mixed,
  `E+` without digits, `/` without denominator, two elapsed codes, two `General`) and rendered
  the other five (`*-0*-`, `e+`, `0.0.0`, `0%E+0`; `CAPTURING.md` §2 item 9; the refusal is Excel's
  repair dialog, confirmed by hand). The first group is `FormatError`; the second is behavior to
  match. Consumers catch `FormatError` and fall back to `General`, which is what Excel's recovery
  does to the cell. Invalid-format strings are captured rows.
- Invalid values for a valid format never raise; they render as `#`.
- Unsupported types raise `TypeError`; `None` renders as `''`; an empty format string means
  `General`.
- Thread-safe (immutable compiled objects, `lru_cache`).

This signature is the formatter contract the sibling packages `sheetpeek` and `xlformula` rely
on, with `format_cell` / `FormattedCell` as the rich pair.

## 12. Module layout

```
src/xlformat/
  __init__.py      public names, __version__
  _parser.py       tokenizer, section split, classification, Section/NumberFormat
  _values.py       input normalization, date serial conversion both ways
  _number.py       number and scientific rendering
  _fraction.py     fraction rendering: denominator search, fixed denominators
  _datetime.py     civil date arithmetic, time rounding, date/time rendering
  _text.py         text and General sections
  _render.py       section selection and dispatch to the renderers
  _assemble.py     width decision through a measure function; fill, pad, overflow; FormattedCell
  _measure.py      measure functions: len (character mode); glyph measure and column model (Excel mode)
  _glyphs_calibri.py  measured Calibri advances per size (generated by capture/glyph_fit.py)
  _palette.py      DEFAULT_PALETTE
  _quirks.py       Quirks
  _errors.py       FormatError
tests/
  capture/          TSV suites and loader (see CAPTURING.md)
  test_*.py        unit tests per module, strict-preset tests, API tests
capture/            generator, capture scripts (not part of the wheel)
```

## 13. Performance

Targets: parse under 50 µs for a typical format; format under 10 µs for `#,##0.00`, under 20 µs
for dates; every captured row under 3 s. Techniques: parse once and cache, precompute everything
derivable from the format, no regex at render time, no `datetime` objects, Decimal only where
precision needs it.

Measured 2026-10-24 with every section kind implemented (Python 3.11, Apple Silicon):

| operation | time |
|---|---|
| parse `#,##0.00` (cache cleared) | 16.5 µs |
| parse `General` | 4.4 µs |
| parse the four-section accounting format | 61 µs (over target) |
| format `#,##0.00` with 1234.5 | 4.9 µs |
| format `General` with 1234567.891 | 3.2 µs |
| format `m/d/yy` | 5.7 µs |
| format `yyyy-mm-dd hh:mm:ss.000` | 8.0 µs |
| format `0.00E+00` | 8.9 µs |
| format `# ?/?` and `# ??/??` | 7.1 µs and 7.5 µs |
| full capture run, 30 suites, 270,549 rows, both width modes | about 3 s |

## 14. What the code taught

- A bare `[$-x-systime]` or `[$-x-sysdate]` tag is the whole format: the section has no date
  codes, yet Excel shows the system time or long date. Classification treats a section whose
  locale is one of these as date/time.
- Excel honors the **last** `*` fill in a section, not the first (`*-0*-` shows `1235` then
  dashes), so `FormattedCell.fill` records the last one.
- The 2023 prototype stripped trailing `#` placeholders from the decimal part, which makes `0.0#`
  behave like `0.0`; the 2023 capture never had such a format so it passed. The renderer keeps
  them and drops trailing zero digits only over `#` and `?` positions (`1.55` → `1.55`, `1` →
  `1.0` under `0.0#`, `1.` under `0.##`); the `numeric_core` suite covers this.
- `-0.5` under `0` renders `-1` (half-up on the magnitude, sign kept); under `0.00` a value that
  rounds to zero drops its sign. Both follow from "round the magnitude, then decide the sign".
- Grouping separators are decided per digit place from what sits to their left: a digit gives
  `,`, a `?` pad gives a space, an unfilled `#` gives nothing. This reproduces `0,000` → `0,001`,
  `#,###` → `` for zero and `#,##0` for every magnitude without special cases.
- An empty first section hides numbers (`;` shows nothing for every number; `;;;` for text and
  booleans too), while `[Red]` or `[>1]` alone is kept by Excel as `[Red]General` / `[>1]General`.
- Sign rules with conditions, from 14 formats: a section whose condition can only hold for
  negatives (`<` with bound ≤ 0, `<=` with bound < 0) drops the sign like the implicit negative
  section; the "everything else" second section keeps it; a lone failing condition falls back to
  General, dropping the sign when the bound is negative, and shows `#` when the body is
  `General`. Four `[<=0]0` rows (negatives that round to zero show General of the magnitude)
  were known deviations until the re-render rule (`CAPTURING.md` §2 item 17) explained them.
- An empty-string value renders nothing, literals included (`@" units"` with `""` is empty).
- The clock arithmetic needed the double, not the 15-digit value, and two different paths for
  clock and elapsed fields (section 7). The first attempt rounded date-only formats to the
  nearest day and broke 4,116 rows; Excel always rounds to the second.
- The epoch constants were first typed by hand and were two days off; they are now derived
  from `days_from_civil`, and a round-trip test guards them.
- **openpyxl writes floats with 16 significant digits** (`"%.16g"`), so the captures' constants
  denoted a neighboring double for about half of all floats, and three rules had been fitted
  to the wrong doubles: the 16-digit step of the tie rule, the fraction-first clock nesting and
  the millisecond-first elapsed path. Formula cells (`=44927+(6*3600+30*60+30.5)/86400`) give
  Excel the double Python holds, and reading the significand back through `INT(B*2^P/2^30)`
  proved Excel's parser correctly rounded. Floats are now written as `repr` text
  (`CAPTURING.md` §2 item 16). Lesson: before fitting a rule to boundary rows, make sure the
  capture's input is the double you think it is.
- **Nine digits per row beat a thousand rows.** The leaked `General` seconds (`0.499999803`)
  pinned the order of Excel's floating-point operations where half-second rows, which only say
  "up" or "down", could not.

## 15. Decisions

- Negative serials under the 1904 system render as a minus sign followed by the magnitude's
  rendering, and `strict` does the same for lack of a stated alternative (section 7; `date1904`
  suite).
- `FormattedCell.fill` reports the position of the fill in `text`, and is `None` without a `*`
  (section 10).
- Width-less overflow is 255 `#`, Excel's clipboard rendering (`CAPTURING.md` §2 items 1 and 2).
- Only en-US for 0.1: locale tags are parsed and exposed on `Section` but not applied, apart
  from the system date and time tags (sections 1 and 7).
- Error values such as `#N/A` are the caller's responsibility, not inputs (section 5).
- The fraction rules ship with zero deviations (`QUIRKS.md`; `CAPTURING.md` §2 item 22).
- Lower-case `e+` after digits is a `FormatError`, and `e` alone is a year code (`CAPTURING.md`
  §2 item 9).
- A condition in section 3 or 4, or on the text section, is a `FormatError`, because Excel refuses
  the file (`CAPTURING.md` §2 item 9, second run).
