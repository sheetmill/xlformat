# Quirks

Excel behaviors that `xlformat` reproduces on purpose and that a caller may switch off through
`Quirks` (see `ARCHITECTURE.md` §9). Every entry names the switch, states what Excel does, what
`strict` does instead, and the evidence (rows captured from Excel 16 for Mac). Names are
shared with `xlformula` where the quirk exists in both.

## `fifteen_digit_precision`

**Excel**: a number is displayed from at most 15 significant digits. The exact binary value of
the double is rounded to 15 significant digits with ties **down**; only then is the format's own
rounding (half-up at the displayed decimals) applied. An earlier version of this entry had an
extra step for values at or above 1e14 (round to 16 digits first); it fitted number texts that
openpyxl had rounded to 16 significant digits, not Excel (`CAPTURING.md` §2 items 15 and 16).

**Strict**: the exact value is used with no 15-digit rounding.

**Evidence** (`tests/captures/rounding_15digit.tsv`, 135 rows, recaptured 2026-10-24 with exact
number texts; `clock_formulas`, 270 numeric rows written as constants and as formulas; the 1,695
float rows of the 2023 capture under `0`, `0.00` and thirty decimals; zero failures on 69,690
numeric captured rows in all):

| value written | exact double | format | Excel | note |
|---|---|---|---|---|
| 0.1234512345123456 | …3455999… | 30 decimals | 0.123451234512346000… | 15 digits, 16th rounds up |
| 12345123451234.25 | exact | `0.0` | 12345123451234.2 | exact tie at the 16th digit rounds down |
| 12345123451234.75 | exact | `0.0` | 12345123451234.7 | same |
| 12345123451234.55 | ….55078125 | `0.0` | 12345123451234.6 | above the tie rounds up |
| 123451234512345.5 | exact | `0` | 123451234512345 | 15 integer digits: tie down |
| 123451234512345.52 | ….515625 | `0` | 123451234512346 | above the tie (was `…345` for the 16-digit text `….5`) |
| 123451234512345.56 | ….5625 | `0` | 123451234512346 | above the tie |
| 1234512345123455.25 | exact | `0` | 1234512345123460 | ….455.25 is above the tie at the 16th digit |
| 1234512345123454.5 | exact | `0` | 1234512345123450 | ….454.5 is below it |
| 999999999999999.5 | exact | `0` | 999999999999999 | tie down, no carry |

## `leap_year_1900`

**Excel**: in the 1900 date system serial 60 displays as February 29, 1900, a day that did not
exist, inherited from Lotus 1-2-3. Serials 1 to 59 therefore map to January 1 to February 28,
1900 and serial 0 displays as "January 0, 1900"; from serial 61 on, dates are correct. Weekdays
come from the serial (`(serial + 6) mod 7`, Sunday = 0), so January 1, 1900 (serial 1) shows as
a Sunday although it was a Monday, and serial 0 shows as a Saturday.

**Strict**: serial `d` maps to `1899-12-30 + d` for every `d` (LibreOffice's model), weekdays
are the true ones, and there is no February 29, 1900. Dates in January and February 1900 then
differ from Excel by one day, which is the price of a correct calendar.

**Evidence** (2023 capture, formats `[$-x-sysdate]dddd, mmmm dd, yyyy`, `ddd, m/d/yy`, `m/d/yy`):

| serial | Excel (`dddd, mmmm d, yyyy`) |
|---|---|
| 0 | Saturday, January 0, 1900 |
| 1 | Sunday, January 1, 1900 |
| 59 | Tuesday, February 28, 1900 |
| 60 | Wednesday, February 29, 1900 |
| 61 | Thursday, March 1, 1900 |
| 62 | Friday, March 2, 1900 |
| 123456 | Wednesday, January 3, 2238 |
| 44927 | Sunday, January 1, 2023 |

Serials 59 to 62 come from the `datetime` suite (2026-10-23); the rest from the 2023 capture.
Under the 1904 system (`date1904` suite) serial 0 is Friday, January 1, 1904 and there is no
leap quirk; negative serials render as a minus sign followed by the magnitude's date or time.

## `fraction_denominators`

**Excel**: the fraction a `?/?`-style section shows is not the closest one.

- *Free denominators* (`# ?/?`, `# ??/??`, `# ???/???`): a continued-fraction expansion of the
  double that stops when the remainder drops below 5e-8 or the next denominator would exceed
  the placeholder limit, taking the last convergent that fits. 0.91 under `# ?/?` shows `1`,
  not `8/9`, the closest fraction with a one-digit denominator, and 0.123 under `# ??/??` shows `8/65`.
- *Fixed denominators* (`# ?/10`, `# ??/16`, `# ???/1000`): a nudged half-up in double
  arithmetic. With `W = floor(x)` and `f = x − W` (exact), `f' = fl(f · (1 + 2⁻⁵²))`,
  `t = fl(f' · D)`, and the numerator is `floor(t)`, plus one when `t − floor(t) ≥ 0.5`. The
  nudge is one relative epsilon (DBL_EPSILON exactly: the window that fits every row is
  0.995 to 1.005 times `2⁻⁵²`) applied to the fraction part, before the product. Because the
  nudge adds one ulp of `f` below a mantissa of 1.5 and two above, a product two ulps below a
  half step rounds up only in the upper half of each binade, which is why no tolerance fitted
  the data (`CAPTURING.md` §2 item 22).

**Strict**: the closest fraction. Free denominators use `fractions.Fraction.limit_denominator`;
fixed denominators round the exact product half-up.

**Evidence** (the `fraction` suite, 2,205 rows; `fraction_ties`, 237 rows at half steps;
`fraction_threshold`, 710 rows, every value as `repr` text and as a formula pinning the double;
`fraction_random`, 30,300 rows predicted before capture; the 5,310 fraction rows of the 2023
capture and of `capture2026`; and 6,351,456 free-denominator rows of unpublished 2023
captures, `CAPTURING.md` §2 item 22): zero misses in every suite. The fixed rule
replaced an empirical tolerance on 2026-10-25; the tolerance missed 484 of 38,904
fixed-denominator rows.

| value written | format | Excel | strict | note |
|---|---|---|---|---|
| 1.15 | `# ?/10` | `1 1/10` | `1 1/10` | `1.15 · 10` is exactly 11.5 in doubles, yet Excel rounds down: the nudged fraction part times 10 is 1.4999… |
| 1.515 | `# ??/100` | `1 52/100` | `1 51/100` | the fraction part times 100 is 51.49999…, the nudge carries it over the half |
| 0.9499999999999997 | `# ?/10` | `1` | `9/10` | 3.55e-15 below the half step, rounds up |
| 0.8499999999999996 | `# ?/10` | `8/10` | `8/10` | the same deficit in the same binade rounds down: the nudge is relative to `f` |
| 0.7504999999999997 | `# ???/1000` | `751/1000` | `750/1000` | lower bound of the constant's window |
| 0.09374999999999997 | `# ??/16` | `1/16` | `1/16` | upper bound of the window: two ulps below 1.5, rounds down |
| 0.03124999999999999 | `# ??/16` | `0` | `0` | 1.5 ulp below 0.5; `floor(t + 0.5)` would show `1/16` |
| 0.031249999999999997 | `# ??/16` | `1/16` | `0` | half an ulp below 0.5, rounds up |

## Known deviations

None. Rows Excel renders in a way the package does not reproduce would be listed in
`tests/test_captures.py` and here, with the evidence, and counted on every run. History: 99
entries (floating noise in the leaked `General` seconds, half-second boundaries on serial 44927)
were removed on 2026-10-24 when Excel's clock arithmetic was reproduced operation for operation
and the captures began writing exact number texts (`CAPTURING.md` §2 items 15 and 16); 32 entries
(`fraction_ties`, the fixed-denominator numerator near half steps) were removed on 2026-10-25
when the nudged half-up rule replaced the empirical tolerance (`CAPTURING.md` §2 item 22).
Fifteen color rows of the `colors` suite are skipped under `tests/test_captures.py:
DISPLAY_FORMAT_COLOR_MISREPORTS`: Excel's DisplayFormat API reports automatic black where the sheet
paints the section color (verified by eye); the package is right there (`CAPTURING.md` §2 item 21).

## Candidates (evidence needed before they become switches)

- `[$-x-sysdate]`: the canned format says `dd` but Excel displays the day without a leading zero.

Matched unconditionally as Excel semantics, not as quirks (no caller would want otherwise):
scientific zero rendering (`0` under `##0.0E+0` shows `000.0E+0`, every placeholder a zero;
`scientific` suite), the section-selection sign rules, the last `*` fill winning, and `General`
inside a date section.
