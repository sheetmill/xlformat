# Capturing Excel's output: regenerating and extending the suites

Section 2 records probe results in the order they were found; a later item that overturns an
earlier one says so in both places. Section 3 describes the capture pipeline.

## 1. What exists

`tests/captures/capture2023.tsv.gz` (89,680 rows) is a 2023 clipboard capture made with an earlier,
unpublished prototype: openpyxl wrote one row per (value, format) pair with the format
applied to column A as a real number format, columns B and C carry `"<type> <repr>"` and
`repr(format)`; the workbook was opened in Excel for Mac, the cells selected and copied, and the
clipboard text saved. The clipboard text is the displayed text with padding spaces but without
column-width effects: `1234567.911` under `General` is in the file although it cannot fit the
default 8.43-character column. So the clipboard is not `Range.Text`, and width needs a different
capture.

## 2. Driving Excel for Mac with AppleScript

Excel 16.113.4's scripting dictionary (`sdef "/Applications/Microsoft Excel.app"`) has what the
capture needs on `range`:

| property | meaning |
|---|---|
| `string value` (read-only, `any`) | "Returns the text for the range": the `Range.Text` equivalent |
| `value`, `formula`, `number format` | read/write |
| `column width` | read/write, in characters, per column |
| `horizontal alignment`, `display format` | alignment and the effective display format object (its font color may reflect number-format colors; to test) |

and on `workbook`: `date 1904` (read/write). Commands `open workbook`, `close ... saving no`,
`copy range`, `autofit`.

**Running scripts against Excel for Mac.** The terminal that runs `osascript` needs macOS
Automation permission for Microsoft Excel; without it the first script fails with `-1743`.
What the scripts in `capture/` do, and why:

- scripts are non-interactive: `display alerts` off, no dialogs; every script that turns alerts
  off turns them back on before exiting, on the error path too (body wrapped in
  `try ... on error ... end try`, restore in both branches). The setting set from AppleScript
  persists until Excel quits (VBA's automatic reset applies only to macros), and while it is
  off Excel silently fails to open or repair files opened by hand;
- workbooks to be opened live inside Excel's sandbox container,
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/xlformat/`, and are opened by that
  resolved path. Excel for Mac is sandboxed: opening a file anywhere else makes macOS show a
  "Grant Access" dialog per file, which a script cannot dismiss. Verified 2026-10-23: three
  workbooks built there opened and captured in 4.6 s, 0.9 s and 0.8 s with nobody at the keyboard, so no
  dialog was pending; the Launch Services fallback (`open -a "Microsoft Excel"`) was not needed.
  `capture/build/` keeps manifests, generated scripts and logs only;
- captures use the clipboard for width-less text; `capture/capture.py` saves the text clipboard
  before the run and restores it after, so anything copied meanwhile is lost;
- Excel's dictionary claims many ordinary words as terms (`row`, `lines`, `clip`, `tab`, `text`
  behave unexpectedly inside a `tell` block); prefix script variables with `xl` and build output
  with ASCII 30/31 separators, never `tab`;
- a script operates only on workbooks it creates or opens by path, closes them without saving
  unless saving is the point, and never touches any other open workbook.

**Probe results (2026-10-23, Excel 16.113.4, macOS en_US).** Scripts: `capture/probe_behavior.applescript`,
`capture/probe_timing.applescript`, `capture/probe_file.applescript` and
`capture/probe_file_many.applescript`, which read the workbooks `capture/build_probe_workbook.py`
writes (items 8 and 9); raw output under `capture/build/*.out` (gitignored).

1. **`string value` is `Range.Text` and reflects column width.** In a new workbook, column
   width 6 versus 30: `-1` under `m/d/yy` gives `######` versus thirty `#`; `1234567.911` under
   `General` gives `1E+06` versus `1234567.911`; `1234.5` under `0.00*-` gives `######` versus
   `1234.50` followed by 35 dashes; `123456789` under `0.00` gives `######` versus
   `123456789.00`; text `hello world text` is returned whole at both widths (text never
   overflows); `123456789012` under `General` gives `1E+11` versus `1.23457E+11`.
2. **Width-less overflow is the clipboard's 255 `#`, not `string value`'s.** `string value` gives
   as many `#` as fit the column (30 at width 30), so the 255 of the 2023 capture is the
   clipboard rendering, equal to a 255-wide column (Excel's maximum). The library's
   `width=None` answer stays 255 `#`.
3. **Fill and overflow are glyph-measured, not character-counted.** 35 dashes at width 30, 53
   spaces in the accounting format at width 30, six `#` at width 5 for an overflowing number:
   Excel fills pixels with the glyph's advance width. Character mode will therefore differ from
   Excel on fill counts by design; Excel mode needs the glyph tables.
4. **Bulk reads work and are 75 times faster.** 1,000 cells: per-cell `string value` 16.8 s
   (16.8 ms each); `string value of every cell of range` 0.22 s; the multi-cell
   `string value of range "A1:A8"` also returns a list (unlike VBA's `Null`); `value of range`
   0.018 s; `copy range` plus `pbpaste` 0.034 s. Capture is bulk only.
5. **Number-format color is readable**: `color of font object of display format of range` gave
   `{0, 0, 255}` for `-1234.5` under `[Red]0.00;[Blue]-0.00` while the cell's own font color
   was black. Color capture goes through `display format`. `horizontal alignment` reports
   `general` for everything (the setting, not the resolved alignment), so alignment is derived
   by the library, not captured.
6. **Fonts.** Excel for Mac's standard font is Calibri 12; workbooks written by openpyxl have
   Calibri 11 as the Normal style, and `string value` and autofit in those workbooks use it.
   Column width to points in the Calibri 12 workbook: `points = 6 * width + 5` (width 1 → 11,
   8.43 → 56, 10 → 65, 30 → 185), so one width unit is a 6-point digit and padding is 5 points.
7. **Autofit measures text in width units.** Calibri 11 workbook, `General`: `1` → 1.33, `12` →
   2.33, `12345678` → 8.33, `12345678901` → 11.33 (digits are one unit each plus 0.33 padding);
   `-12345678` → 8.83 (`-` is half a unit); `1E+15` → 5.0, `1.23457E+11` → 11.0 (`E`, `+`, `.`
   are narrower than a digit); `hello world text` → 12.33 (16 characters); `TRUE` → 4.17. In the
   Calibri 12 workbook: `Tuesday, January 4, 2000` → 21.0 and `Sunday, December 3, 2000` →
   22.33, which matches what the sheet shows by hand. Every width-general value has its measurement in
   `tests/captures/width_general.measure.tsv`; that file is the seed for the Calibri 11 glyph table.
8. **Rounding at 15 digits (point d).** Workbook `probe_halfdown.xlsx`, formats `0`, `General`,
   `0.0`:

   | value (exact binary) | `0` | `0.0` |
   |---|---|---|
   | 123451234512344.5 | 123451234512344 | 123451234512344.0 |
   | 123451234512345.5 | 123451234512345 | 123451234512345.0 |
   | 123451234512345.52 (…345.515625) | 123451234512345 | 123451234512345.0 |
   | 123451234512345.48 (…345.484375) | 123451234512345 | 123451234512345.0 |
   | 123451234512345.6 (…345.59375) | 123451234512346 | 123451234512346.0 |
   | 999999999999999.5 | 999999999999999 | 999999999999999.0 |
   | 1234512345123455.5 | 1234512345123460 | 1234512345123460.0 |
   | 1234512345123454.5 | 1234512345123450 | 1234512345123450.0 |
   | 1234512345123456.5 | 1234512345123460 | 1234512345123460.0 |
   | 12345123451234.5 | 12345123451235 | 12345123451234.5 |
   | 99999999999999.5 | 100000000000000 | 99999999999999.5 |
   | 2.5 / 0.125 / 1.0625 / 1234.5 | 3 / 0 / 1 / 1235 | 2.5 / 0.1 / 1.1 / 1234.5 |

   `General` shows `1.23451E+14` and friends for all large values (no information). Models
   tested against these 48 rows and the 1,695 float rows of the 2023 capture under `0`, `0.00`
   and thirty decimals (`capture/rounding_models.py`):
   - correct rounding of the exact binary value to 15 significant digits, ties up: fits every
     2023 row, fails 8 of today's (all values with 15 or more integer digits);
   - ties down at 15 digits, or truncation: fail both sets;
   - round the exact value to 16 significant digits, then to 15 with ties down: fits all 48 of
     today's rows, fails 23 2023 rows such as `0.1234512345123455` (exactly
     0.12345123451234550626…, which Excel shows as `…346`, so for small values the 16th digit is
     not what decides);
   - **hybrid, fits everything seen**: below 1e14 round the exact value to 15 digits with ties up;
     at or above 1e14 round to 16 significant digits first, then to 15 with ties down.
   **Grid run (suite `rounding_15digit`, 135 rows, script `capture/rounding_models.py`).** 14, 15
   and 16 integer digits with fractions `.25 .4999 .5 .5001 .51 … .59 .75` under `0`, `0.0`,
   `0.00`. Decisive rows: `12345123451234.25` under `0.0` shows `….2` and `.75` shows `….7`, so
   exact ties round **down** even below 1e14 (ties-up was wrong; the 2023 rows never contained
   an exact tie at the 16th digit). `123451234512345.51` to `.55` (exact `.515625` to `.546875`)
   show `…345` while `.56` (`.5625`) shows `…346`. Model failure counts, grid / 2023 floats:

   | model | grid | 2023 |
   |---|---|---|
   | exact value → 15 digits, ties up | 34 | 6 |
   | exact value → 15 digits, ties down | 18 | 0 |
   | exact value → 15 digits, ties even | 32 | 6 |
   | truncate to 15 digits | 64 | 281 |
   | exact → 16 digits (even) → 15 digits ties down | 2 | 11 |
   | hybrid: ties up below 1e14, 16-then-15-down above | 4 | 0 |
   | **best fit: ties down below 1e14; 16 digits then 15 ties down at or above** | **0** | **0** |

   **Superseded by items 15 and 16**: the 16-digit step fitted number texts openpyxl had rounded
   to 16 digits, not Excel. With exact inputs the rule is plain: the exact double rounded to 15
   significant digits with ties down, zero misses on every numeric row. `QUIRKS.md` carries the
   settled table.
9. **Formats Excel refuses in a file (point a).** 21 formats the dialog rejects were each written
   with openpyxl into their own workbook and opened with alerts off. Sixteen made Excel refuse
   to open the file (`open workbook` fails with "Parameter error (-50)"): `0.00e+00`,
   `0.00e-00`, `##0.0e+0`, `0.00E+00;0.00e+00`, `0;0;0;0;0`, `[Red`, `0.00"`, `[Red][Blue]0`,
   `[>1][<5]0`, `@;@`, `0;@;0`, `yyyy 0`, `0 yyyy`, `0.00E+`, `# ?/`, `[hh]:mm:[ss]`,
   `General General`. These are `FormatError`. **Confirmed by hand**: after relaunching Excel,
   double-clicking one of these workbooks brought up the repair dialog ("We found a problem
   with some content in 'inv_01.xlsx'. Do you want us to try to recover as much as we can?"),
   so `-50` is Excel refusing the file with the dialog suppressed. What the recovered cell shows
   is expected to be General (Excel's recovery drops the style), not yet observed. Five opened
   and rendered, so they are formats to match, not errors: `*-0*-` (two fills; the second one fills: `1235` then dashes), `e+` (`e`
   is a year code: `1234.5678` → `1903+`, `0` → `1900+`, `-0.5` → `#` run), `0.0.0` (the second
   point is a literal inside the decimal digits: `1234.5678` → `1234.5.7`), `0%E+0` (`%` is a
   literal here and does not scale: `1234.5678` → `1%E+3`, `-0.5` → `-5%E-1`). Setting
   `0.00e+00` through the AppleScript `number format` property (the dialog path) is accepted but
   the cell reads back as `General`.

   **Second run (same day, `capture/invalid.py`, 17 candidates, one workbook each from the
   container, no dialogs).** Refused, now `FormatError` with unit tests: `[Color57]0`,
   `[Color0]0`, `[]0`, `[abc]0`, `ss.0000` (four fractional-second digits), `yyyy @`, `0 @`,
   `@ 0`, `# ?/?/?`, `[>1]0;[>2]0;[>3]0` (condition in section 3), `[>1]0;[>2]0;[>3]0;@`,
   `0;0;0;[>1]@` (condition on the text section). Every refusal the parser had marked
   "unverified" is confirmed. Rendered, so behavior to match: `[Color 3]0` (Excel keeps it as
   `[Color3]0`, the space is dropped); `0 x` → `1235 x`, `-1 x` (a bare unknown letter is a
   literal); `x0` → `x1235`, `-x1` (the sign precedes the literal); `[Red]` alone is kept as
   `[Red]General`; `[>1]` alone is kept as `[>1]General` and shows General for values that pass
   the condition and a `#` run (`#########` at the probe's column width) for values that fail it:
   a lone condition that fails renders like an invalid value, not as General (the 2023 prototype was
   wrong here; captured width-less in `numeric_core` as `[>1]`, `[>100]0.00` and `[<0]0.0`).
10. **Pipeline validation.** `capture/capture.py capture2023_sample --diff-2023`: 1,993 rows (every
    45th of the 2023 capture, all 152 formats) rebuilt with openpyxl, captured with `copy range`
    plus `pbpaste`: **0 differences**. The clipboard method is reproduced exactly and the 2023
    build and 16.113.4 agree on every sampled row.
11. **`width_general` captured**: 58 values × (clipboard + 17 widths) = 1,044 rows plus 58
    autofit measurements, in `tests/captures/width_general.tsv` and `.measure.tsv`. First reading of
    the data: General drops decimals one at a time (rounding; `999.9999999` becomes `1000` as
    soon as one decimal is dropped) until the text fits; when the integer part itself does not
    fit, scientific with as many mantissa decimals as fit, down to `1E+07`; when not even that
    fits, `#` to the width (`####` at width 4, `##` at 2). Values below one reduce to `0` rather
    than overflow (`0.000123` at width 4 is `0`, at width 5 `1E-04`, at width 6 `0.0001`).
    Negative values need one more unit for the sign (`-12345678` is `######` at width 5 where
    `12345678` is `1E+07`). Booleans overflow to `#` like numbers (`TRUE` at width 3 is `###`);
    text never does. The exact reduction rule is derived from this file when the width work
    starts.
12. **Conditions and rounding (2026-10-24, suite `conditions_rounding`, 187 rows, 11 formats).**
    The hypothesis "a condition sees the value after rounding to the section's precision" is
    false: `[=0]"zero";0` shows `0` for 0.001 and `[<>0]0;"zero"` shows `0` for it, so the
    unrounded value is compared. Two sign rules came out instead. The "everything else" second
    section drops the sign when the failing range can only hold non-positive values
    (`[>0]0;"neg"` shows `neg` for -0.001; `[>100]0.00;0` shows `-1`). A lone `[<=0]` section
    whose rendering of a negative value rounds to zero shows General of the magnitude
    (`[<=0]0` shows `0.001` for -0.001 and `-1` for -0.5; `[<=0]0.0` shows `0.001` for -0.001 and
    `-0.1` for -0.05), while `[<0]0`, `[<1]0`, `[>=-1]0`, `[<=0]0;0` and `0;[<=0]0` all show
    `0`. Both are implemented as observed; the second is narrow (one section, `<=`, bound 0)
    and its mechanism is unknown. The four `[<=0]0` known deviations are gone.
13. **Clock boundaries (2026-10-24, suite `clock_boundaries`, 186 rows).** Ten half-second times
    built four ways (sum of fractions, nested products, total seconds, and on serial 44927)
    under `ss`, `[mm]:ss`, `h:mm:ss`, plus 66 serials under `m/d/yy General`. For fractions
    below one day, three of the four constructions give the same double and Excel agrees with
    `((frac*24)*60)*60` rounded half-up in every case but one (`06:30:30.5` as a sum of
    fractions rounds up where its neighbor, one ulp away, rounds down). On serial 44927 plus a
    time, Excel rounds `.5` down in seven of ten cases regardless of the sign of the exact
    distance from the half (+1.1e-7 s rounds down, +2.2e-7 s rounds up). Nine candidate
    computations were tried (direct, nested, fraction-first, peeled hours/minutes/seconds,
    millisecond-first, single precision); the best, fraction-first nesting, fits 248 of 264
    seconds-bearing rows across the suites and is what the package uses; the 21 remaining rows
    (7 values x 3 formats) are known deviations. The leaked `General` noise is a step function of
    the serial (`0.5` to 100, `0.499999999` from 200 to 3000, `0.500000038` at 5000 to 7000,
    `0.49999996` at 10000 to 15000, `0.500000117` at 20000 to 30000, `0.499999803` from 36526 to
    200000, `0.500002317` at 500000, `0.499997288` at 1e6 to 2e6, `0.500017405` at the limit): the
    size of double rounding error on total seconds, but no formula tried (multiplying by 86400
    or by 24, 60, 60 in any order; dividing by 1/86400, 1/1440 or 1/24; from the fraction or the
    whole serial) reproduces it. Those 54 rows plus the 24 in `datetime` stay known deviations.
    **Resolved by items 15 and 16**: the noise is the peel of day, hour, minute and second from
    successive fractions of a biased serial, and the half-second misses were doubles openpyxl had
    written with 16 digits.
14. **Fraction formats refused in a file**: of 22 candidate fraction formats, only
    `#,##0.00 ?/?` (a decimal point before the bar) is refused; `?/?`, `0/0`, `#/#`, `0 0/0`,
    `# ???/1000`, `?/?%` and `# ?/?;(# ?/?)` all render. `%` scales in a fraction section
    (`-0.5` under `?/?%` is `-50/1%`), unlike in a scientific one.
15. **Formula-written boundary cases (2026-10-24, suite `clock_formulas`, 642 rows; comparison
    script `capture/formula_diff.py`).** Every boundary value of `clock_boundaries` and
    `rounding_15digit` written twice: as the constant and as a formula with literals of at most
    15 digits (`=44927+(6*3600+30*60+30.5)/86400`, `=123451234512345*10+5.25`) that Excel
    evaluates in IEEE arithmetic, so the formula's double is exactly Python's (`Case.formula`;
    the tag carries `formula:` plus the text, the schema is unchanged). First capture: 16 of 210
    constant/formula pairs differed, and the formula rows contradicted the package on 71 rows.
    Taking the formula rows as truth, a brute-force fit over operation orders, scored against
    the leaked `General` seconds (nine digits per row over 66 serials) and every date and time
    row in every suite, gave Excel's clock arithmetic exactly (`ARCHITECTURE.md` §7):
    `d = x + (0.5/scale)/86400`, then day, hour, minute and second peeled from successive
    fractions and truncated; elapsed codes `floor((x / (1/86400) + 0.5/scale) / unit)`; a
    `General` token shows the biased seconds before truncation. Zero misses on 41,266 date and
    time rows. The 59.5-second rows separate `x / (1/86400)` (59.5) from `x * 86400`
    (59.49999999999999); the day-44927 rows separate a bias in days from a bias in seconds. The
    number display rule is likewise plain: the exact double rounded to 15 significant digits
    with ties down, zero misses on 69,690 numeric rows; the 16-digit step of item 8 was an
    artifact of item 16. After the recapture of item 16 the 210 pairs agree in every case.
16. **openpyxl writes floats with 16 significant digits (suite `parser_check`, 600 rows).**
    `openpyxl.compat.strings.safe_string` formats numbers with `"%.16g"`, so for about half of
    all floats the text in the file denotes a double one ulp from Python's (1,240 rows across
    the suites at the time, 456 of them in the 2023 capture, which used the same writer). The
    suspected parser error does not exist: `parser_check` writes 100 values both as the
    16-digit text and as `repr` and reads the loaded significand back bit-exactly through
    `=INT(B*2^P/2^30)` and `=B*2^P-INT(B*2^P/2^30)*2^30` (exact in IEEE arithmetic, shown
    exactly by `General`); all 200 loaded doubles are the correctly rounded parse of their
    text. Fix in `capture/build.py`: floats are written as `repr(value)` text with
    `data_type = "n"` (`Case.written` overrides the text), recorded as `numbers: repr` in the
    suite header; `tests/test_captures.py` compares suites without that header against
    `float("%.16g" % value)`, the double Excel actually loaded. Recaptured with exact texts:
    `datetime` (0 of 3,630 texts changed), `clock_boundaries` (4 of 186), `rounding_15digit`
    (18 of 135), `clock_formulas` (22 of 642). Known deviations went from 99 to 0.
17. **Negatives that round to zero (2026-10-24, suite `lone_zero`, 96 rows, 12 formats × 8
    values).** The model holds: Excel renders a negative as its magnitude plus a
    minus sign, and when the magnitude renders as zero it renders the positive magnitude again
    from section selection onward. `[<=0]0` shows `0.001` for -0.001, `[<=0.1]0` shows `0.3`
    for -0.3 (the re-rendered 0.3 fails the lone condition and falls to General), `[<=0.5]0`
    shows `0`, `[<>0]0` and `[=0]"zero";0` show `0`, `[<=0]0.0` shows `0.04` for -0.04;
    `[<0]0` and `[<=-0.2]0` show `0` because their sign was dropped before rendering, so there
    is nothing to re-render. One new fact: when both conditioned sections of a two-section
    format fail (`[<=0]0;[>1]0` with 0.3, or with -0.3 after the re-render), Excel shows a `#`
    run, not General; a lone failing condition still falls to General. The narrow
    `_lone_zero_admitting` rule is replaced by the general re-render step in `_render.py`.
18. **Fraction rules against exact texts (2026-10-24, `fraction` recaptured: 0 of 2,205 texts
    changed).** With every fraction row (2,205 plus the 5,310 of the 2023 capture) scored, the
    15-digit rounding of the value before the denominator search changes nothing and was removed.
    The plain rules for the fixed-denominator numerator do not fit: `1.15` under `# ?/10` shows
    `1/10` although `1.15 * 10` is exactly 11.5 in double arithmetic (so the double product rounded
    half-up gives `2/10`), while `1.515` under `# ??/100` shows `52/100` although the fraction
    part times 100 is 51.49999999999999 (so the fraction-part product gives `51/100`); `1.115`
    shows `12/100` and `0.15` shows `2/10`. Half-up on the exact product with the tolerance
    `5e-18 * D**2` fitted all 7,515 rows of the time (superseded by item 22). The whole part
    of a value with 15 or more integer digits is rounded half-down (`123451234512345.9` under
    `# ??/??` shows `123451234512346 29/32`); without it 12 rows fail.
19. **Section selection with a conditioned second section (2026-10-24, suite
    `selection_fallback`, 40 rows, 5 formats × 8 values).** `[<=0]0;[>1]0;0.0`: a value failing
    both conditions takes the third section (`0.3` shows `0.3`). `0;[>1]0`: the unconditioned
    first section takes positives only; `0` and negatives show a `#` run (`0;[>1]0;0.0` sends
    them to the third section with the sign kept: `-0.3`, `0.0`). `[>1]0;0;0.0`: after the
    failed condition the second section is the negative one (`-2` shows `2`, sign dropped) and
    the third takes zero and the positives that failed (`0.3` shows `0.3`, `0` shows `0.0`),
    whereas with two sections (`[>100]0.00;0`) the second is "everything else" with the sign
    kept. `[<=0]0;[>1]0;"else"` confirms the re-render of item 17 (`-0.3` shows `else`).
    `_render.py: select` was rewritten from this table.
20. **`capture2026` (2026-10-24, 89,680 rows, Excel 16.114).** All cases of the 2023 capture
    rebuilt with exact number texts and captured in 3.1 s of Excel time. `capture/diff_captures.py`
    compares two suites row by row and classifies a difference as `text16` (the older suite wrote
    16-digit texts and the value's double differs under them) or `drift` (Excel build or method).
    Result: **zero differences** in 89,680 rows. The 456 rows whose 2023 double was a neighbor
    of Python's displayed the same text, and Excel 16.114 agrees with the 2023 build on every
    case. Both files stay, full-tier only.
21. **Colors (2026-10-24, suite `colors`, 170 rows).** Bulk reads do not work:
    `color of font object of display format of every cell of range` returns `missing value` for
    every cell (`capture/probe_color_bulk.applescript`), so the capture template loops over the
    cells of column B with one Apple event each, about 10 ms per cell (170 cells in 4.5 s with
    the text capture). `Suite.colors` switches the loop on; the TSV stores `RRGGBB`; a cell with
    no section color reads `0,0,0`. Agreement on 155 rows: the eight names, all 56 indices
    (`src/xlformat/_palette.py: DEFAULT_PALETTE` is the legacy palette, confirmed value by
    value; `Color25` = `Color11` and other duplicates included), `[color 3]` and `[COLOR3]` as
    `Color3`, color in sections 2 to 4, color with conditions in either order (`[<0][Red]` is
    kept as `[Red][<0]`), `[Red]` alone and `[Red];[Blue]` (kept as `[Red]General;[Blue]General`),
    scientific, fraction, date, elapsed and percent bodies, and the re-rendered `-0.3` under
    `[Red][<=0]0;[Blue]0` in blue. Fifteen rows read as automatic black where the section has a
    color, in three patterns: a value of exactly zero rendered by section 1 or 2 (`[Red]0` with
    0; `[Red][<0]0;[Blue]0` with 0; `[Green][=0]"zero";[Red]0` with 0), while zero in the third
    section is colored (`0;0;[Green]0`); text or a boolean through a colored text section
    (`[Cyan]@`, `0;-0;0;[Magenta]@`); and a `#` run (`[Red]m/d/yy` with -1234.5). Checked by eye on the
    sheet (2026-10-24): B9 (zero under `[Red]0`) is red, B132 (`abc` under `[Cyan]@`) is cyan and
    B147 (the `#` run under `[Red]m/d/yy`) is red. **The DisplayFormat font color is unreliable
    for exactly these three cases**: zero rendered by section 1 or 2, text and booleans through a
    colored text section, and `#` runs; the sheet paints the section color and so does the
    package. The fifteen rows are skipped under `tests/test_captures.py:
    DISPLAY_FORMAT_COLOR_MISREPORTS`; a future color capture needs another source for them.
22. **Fixed-denominator numerators at half steps (2026-10-24, suite `fraction_ties`, 237 rows).**
    For D in 10, 100, 1000, 16, 7, 3 and whole parts 0, 1, 12, 1000, doubles whose fraction part
    times D sits a measured number of ulps below or above a half step (the tag records the
    distance; whole parts above zero reach only coarse distances, so the generator varies the
    numerator k to land near each target), plus exact binary ties for D in 2, 4, 8, 16, 32 (all
    round up). Exact texts. Models scored on these and on every other fixed-denominator row
    (7,770 in all): the current rule (half-up on the exact product with the tolerance
    `5e-18 · D²`) misses 32, all in this suite; half-up on the double product `x·D` misses 98;
    half-up on the fraction part times D misses 30; a tolerance of c ulp of `k + 0.5` misses 11
    at best (c = 2); comparing after rounding the scaled fraction to 15 digits misses 33. The
    threshold is relative, not a number of ulps: two rows in the same binade with the same
    deficit (`0.9499999999999997` and `0.8499999999999996` under `# ?/10`, both 3.55e-15 below
    the half) round opposite ways, up for 9.5 and down for 8.5. Half-up on the double product of
    the fraction part with a tolerance of `c · 2⁻⁵² · (k + 0.5)` for c between 1.7 and 1.8 misses
    exactly one row (`0.03124999999999999` under `# ??/16`, 1.5 ulp below 0.5, rounds down);
    division by `1/D` instead of multiplication changes nothing. Decision (time-boxed): the
    current tolerance stayed and the 32 rows were known deviations.
    **Resolved 2026-10-25 by a larger capture of half-step values**, included as
    `fraction_threshold` and `fraction_random`. The rule, with `W = floor(x)`
    and `f = x − W` exact: `f' = fl(f · (1 + 2⁻⁵²))`, `t = fl(f' · D)`, numerator `floor(t)`
    plus one when `t − floor(t) ≥ 0.5`. Only DBL_EPSILON fits (window 0.995 to 1.005 times
    `2⁻⁵²`: `0.7504999999999997` under `# ???/1000` rounds up, `0.09374999999999997` under
    `# ??/16` rounds down); `2⁻⁵³` misses 172 rows, `2⁻⁵¹` 377; nudging `x` instead of `f`
    misses 105, nudging after the product 52, dividing by `1/D` 29; phrasing the half-up as
    `floor(t + 0.5)` misses 17 rows at k = 0 where `t + 0.5` rounds up to 1. The nudge adds
    `round(m)` ulps of `f` for a mantissa `m`, one below 1.5 and two above, which is the
    non-monotone pattern no tolerance fitted. The two suites (Excel 16.114, `repr` texts, full
    tier only):
    `fraction_threshold` (710 rows: k = 0 ladders for D = 2 to 32, every k of D = 16, a D = 1000
    sweep, whole parts 0, 1 and 12, each value as text and as a formula `=(mh·2^27+ml)·2^e`
    pinning the double; all 353 pairs agree) and `fraction_random` (30,300 rows: 3,000 doubles
    log-uniform over [1e-3, 1e6] under ten fixed-denominator formats plus 300 doubles within ten
    ulps of a half step; the texts were predicted before the capture and matched on every row).
    Scored over 38,904 fixed-denominator rows of every suite: the new rule misses none, the
    tolerance missed 484. `_fraction.py: fixed_numerator` carries the rule; the 32 known
    deviations and `capture/fraction_models.py` are gone. The free search was scored against
    6,351,456 rows of unpublished 2023 captures: the continued-fraction walk misses none,
    SheetJS's mediant walk and `limit_denominator` miss about three million each.
23. **Width, character mode (2026-10-24; suites `width_fill`, 1,650 rows at widths 3, 6, 9, 12
    plus the clipboard, and `numeric_core_width`, 32,832 rows at widths 6, 10, 14).** The General
    reduction was derived from the 986 `width_general` rows: decimals drop one at a time with
    half-up rounding and trailing zeros stripped (`1234.5678` at width 6 is `1234.6`); the rounded
    integer is shown when it fits (`1234567.891` at width 7 is `1234568`, preferred over
    `1.2E+06`); when the integer part does not fit, scientific notation keeps as many mantissa
    digits as fit down to `dE+ee` (`12345678` at width 5 is `1E+07`); a value whose unconstrained
    form is scientific stays scientific (`1E+15` at every width) and, when not even `dE+ee` fits,
    shows `0` below width 5 if it is below one and `#` otherwise (`1E-100` is `0` at width 4 and
    six `#` at width 5); below one, a decimal whose digits all round away goes scientific if that
    fits and shows `0` (`-0`) otherwise (`0.000123` at width 5 is `1E-04`, at 4 is `0`; `-1.5E-07`
    at width 5 is `-0`), and a decimal that carried into a new power of ten (`0.00005` at width 6
    as `0.0001`) or shows fewer significant digits than the scientific form that fits
    (`0.0000002` against `1.5E-07` at width 9) loses to it, while `0.99999` at width 6 is `1`.
    Booleans overflow like numbers (`TRUE` at width 3 is `###`), text never does, except through a
    section with a `*` fill (`@*-` with `Hello World` at width 6 is `#`). A negative whose
    `General` reduces to zero drops its sign when literals surround it (`General"x"` with -0.001 at
    width 6 is `0x`, `( General )` is `( 0 )`) and keeps it when bare (`-0.05` at width 3 is
    `-0`). Excel's `#` count is glyph-measured (`#`, `##`, `###`, `####`, then six at width 5 and
    seven at 6 in Calibri 11); character mode shows `width` of them and the test equates `#` runs.
    Glyph slack and tightness: Excel fits `1,000.0` in width 6 and `-12345678` in 8.43 and
    overflows `10000%` at width 6 because `,` `.` space and `-` are narrower than a digit and `%`
    is wider; the test skips such rows (2,224 + 11 + 16) as Excel-mode evidence. Result: every
    other width row passes (986 + 1,639 + 30,608).
24. **Glyph advances (2026-10-24, suites `measure11` and `measure12`, 523 texts each).** Every
    printable ASCII character in runs of 1, 2, 5 and 10, the digit pairs, month and day names in
    full and abbreviated, `AM`, `PM`, `E+07`, `####` and a dozen mixed strings, each autofitted
    in its own column (CAPTURING.md item 7 method) with the column width read back in units and in
    points (`width of range`), in an openpyxl workbook (Calibri 11) and one whose Normal style is
    Calibri 12 (`Suite.font_size`). Points are whole numbers; units and points obey
    `points = 6 · units + 5` at both sizes, so the maximum digit width is 6 points for Calibri 11
    and 12 alike. The autofit width is the floored sum of the glyph advances plus 7 points; a
    glyph's advance is the ten-run less 7 over 10. Advances are hinted per size: digits 6 at both,
    `.` `,` `:` `i` 3 at both, `-` `(` `)` 3 then 4, `+` `E` `#` `a` 5 then 6, `/` 4 then 5, `%` 8
    then 9.1, `W` 10.1 then 11.1, `M` `m` 9.1 then 10.1; a size-independent em fraction cannot
    give 6 for the digit at both sizes, so `_glyphs_calibri.py` holds both measured tables.
    `capture/glyph_fit.py` fits and writes them: 518 of 519 texts exact at each size (the
    exception at 11 and the two at 12 are the long day-name strings, one point off). The space
    is ambiguous in the autofit data (2.0 to 2.4 fit the mixed strings equally); the accounting
    fills decide it (item 25). A blank text autofits to the default width and is ignored.
25. **Excel mode against every width row (2026-10-24).** Column budget `6 · w` points with `w`
    the snapped width (`round(6w + 5) − 5` points); fit when the measured text is at most the
    budget: `1,000.0` (36) fits width 6, `-12345678` (51) fits 8.5, `10000%` (38) overflows 6.
    The `#` run holds `floor(budget / advance('#'))` characters: 1, 2, 3, 4, 6, 7 at widths 1 to
    6 in Calibri 11, all 5,254 overflow rows. The `*` fill repeats `floor((budget − text) /
    advance)` times: 35 dashes at width 30 in Calibri 12; the accounting format's fills at widths
    6, 10 and 14 pin the space advance at 2 points (12 fill spaces around 11 points of text in
    36) and show `_x` pads and `?` blanks measuring as spaces. `General` alone does not use the
    glyph measure: it reduces by characters against the width in units, a leading minus counting
    half (`-1234.6` and `1,000.0` both measure 36 points at width 6, Excel reduces the first and
    keeps the second; `-12345678` shows at 8.5; `-2E-07` needs 6, `-1.5E-07` needs 8); all 1,107
    shown General rows fit that budget and none is reducible further. Among literals the General
    token's budget is the units less the literals' and the sign's advances in digit units (`( `
    and ` )` cost 1.67: `( 1234 )` shows at width 6, `( 12345 )` is `#`, `( 0.125 )` becomes
    `( 0.13 )`), and the scientific form of a value whose integer part does not fit is admitted
    without the sign's cost (`-1E+05x` at width 6) while a value that rounds to zero pays it
    (`-0.001` shows `0x`). Text overflows only through a section whose `*` fill follows it
    (`@*-` with `Hello World` at width 6 is `#`, `*-@` shows the text). Result: 35,174 width rows
    in `width_general`, `width_fill` and `numeric_core_width` pass in Excel mode with no skip, the
    2,251 glyph-slack rows included; character mode is unchanged.

26. **Calibri at 9, 10, 14 and 16 points (2026-10-25; suites `measure9`, `measure10`,
    `measure14`, `measure16`, 523 autofit texts each; `space_fill9` to `space_fill16`, six suites
    of 200 rows, one per measured size including 11 and 12; `width_general14`, 1,044 rows).**
    The item-24 method at four more Normal-style sizes (`Suite.font_size`), each one Excel run
    of about 19 s for the autofit of 523 columns. The autofit padding is 7 points at every size
    (`2 · points(5 digits) − points(10 digits)`, the same for all ten digits), the maximum digit
    width is 5, 5, 6, 6, 7 and 8 points at 9, 10, 11, 12, 14 and 16, and the column obeys
    `points = MDW · w + 5` at every size (`column_points` headers: width 6 is 35, 35, 41, 41, 47
    and 53 points; 8.43 at 14 is 64). `capture/glyph_fit.py` now loops over every
    `measure*.measure.tsv`, reads the size from the header and derives the padding from the digit
    runs; 515 of 515 space-less texts are exact at every size, and of the four spaced texts the
    two long day-name strings are one point off at every size, as before. Hinting per size
    again: `-` is 3, 3, 3, 4, 4, 5; `%` 6, 7, 8, 9.1, 10, 11; `W` 8.1, 9.1, 10.1, 11.1, 12.1,
    14.1; `.` 2, 3, 3, 3, 4, 4; so the tables at 9 and 10 share the digit width and differ
    elsewhere. **The space** cannot come from autofit (item 24), so each size has a `space_fill`
    suite: the two accounting formats, `0* ` and `* 0` over ten values at widths 6, 10, 14 and
    20 plus the clipboard. The fit script renders those 160 width rows through `xlformat` in Excel
    mode with every candidate advance from 1.0 to 8.0 in tenths and keeps the one that matches:
    a single candidate matches all 160 rows at every size, 2, 2, 2, 3, 3 and 4 points (the 11- and
    12-point values agree with the item-25 reading from `numeric_core_width`). `_measure._advances`
    uses any measured table and scales the nearest one, larger on a tie, for other sizes.
    **Validation**: `width_general14` passes in Excel mode on all 986 width rows with the Calibri
    14 table; General's character budget holds at the new size, including one row that differs
    from Calibri 11: `-12345678` at width 8.43 fits at 11 (the column snaps to 8.5 units, 56
    points) and reduces to `-1.2E+07` at 14 (64 points is 8.43 units of 7, below the 8.5 the sign
    and eight digits cost). Nothing in the model changed for the new sizes.

27. **General at every whole-point width (2026-10-28; suites `general_points9`, `11`, `12`, `14`,
    `16`, the 58 `width_general` values at every point from MDW + 5 to 100 points, 25,984 width
    rows; `general_literals_points`, the same values under `General"x"`, `( General )`,
    `"$"General`, `General" units"` and `General;-General` at every point from 11 to 100 at
    Calibri 11, 26,100 width rows).** The earlier General suites use whole dialog widths, where
    a character count and a glyph measure agree; columns read from files are mostly fractional
    (an openpyxl `8.43` is 7.67 units). The 0.1 model missed 158 bare rows and 637 among
    literals, all now matched:
    - When no reduced form fits by characters, the shorter of the rounded integer and the
      one-digit scientific form shows if its glyph measure fits the points budget: `1E+06` is
      28 points and shows at 4.67 units (28-point budget), `-1E+06` at 5.17, `1E+100` at 5.67,
      `1E-100` (a value below one: always the scientific form) at 5.33. On a tie the format
      `General` keeps the integer (`12345` at 4.67 is `#####`: 30 points) while a General
      section takes the scientific form (`( 1E+04 )` at 6.33, `1E+04` under `General;-General`
      at 4.67). This also explains 0.1's "scientific form admitted without charging the sign"
      among literals (`-1E+05x` at width 6: 36 points of 36), which is removed.
    - A leading minus costs its own advance in digits: 3 / 6 at Calibri 11 (0.1's half unit),
      4 / 6 at 12, 4 / 7 at 14, 5 / 8 at 16; at x.57 units the cost at 14 ties the width, so
      cost and budget are compared rounded to nine places (as is `8.17 − 25 / 6` = 4, which fits
      `3.14 units`).
    - Among literals, with less than one character left, a value from 0.0001 to under 10 shows
      `0` (`( 0 )` for -1.5 and -0.05 at 2.67 units, where the sign takes the last half unit;
      `0` for -1.5 under `General;-General` at width 1) and any other value shows nothing
      (-12345678 and -1.5E-07 are `###`; only -1.5E-07 decides the lower bound, which is
      General's existing 0.0001 threshold). A General that reduced to zero drops the sign and
      a `-` literal just before it; bare General keeps `-0`.
    - Every General text passes the glyph test as well (`0 units` is `#` at width 1).
    Character mode: 2 rows of `general_literals_points` where Excel's literals cost less than a
    character each leave room that character mode does not have (`( 0 )` for -1.5E-07 at width
    5) are skipped as glyph slack.

28. **The column model (2026-10-28; suites `column_file9` to `column_file16`, 138 widths
    written into the file as `<col width>` by openpyxl, 19 cases each, at the six measured
    sizes; `column_default_base10_11`, `column_default_base10_14`, `column_default_dcw12_11`,
    `column_default_dcw843_14`, a column without `<col>` under those `<sheetFormatPr>`
    settings; `column_narrow11`, `column_narrow14`, dialog widths below one unit set by the
    script; `limits`, fills and `#` runs near 255 characters).** The script reads back each
    column's `column width` and `width` without setting it (method `file_width`, or
    `file_default` for no `<col>`; the header's `width_source: file`, `sheet_format`).
    - Points of a file width: `trunc((256 · w' + trunc(128 / MDW)) / 256 · MDW)` with
      `w' = trunc(256 · w) / 256`, ECMA-376's pixel formula with MDW in points: 828 of 828
      columns at Calibri 9, 10, 11, 12, 14 and 16, including widths placed just past a half to
      separate it from rounding half down (which misses 6 to 8 per size) and the formula
      without the 1/256 step (one miss at 14). openpyxl's `8.43` is 51 points at Calibri 11,
      `60` is 360, `3.7109375` is 22.
    - A column without `<col>`: `baseColWidth · MDW + 5` points when `defaultColWidth` is
      absent (openpyxl writes `baseColWidth="8"` only): 53 points at Calibri 11 and 12, the
      dialog showing 8.0, not 8.43; `baseColWidth="10"` is 65 points. A present
      `defaultColWidth` is an ordinary file width (12 is 72 points, 8.43 is 51).
    - The dialog shows `(points − 5) / MDW`, and `points / (MDW + 5)` below one unit (3 points
      is 0.27 at Calibri 11), to two decimals half up (1.125 shows as 1.13 at 16).
    - Below one unit a width set in the dialog is `round(w · (MDW + 5))` points (0.09 is 1
      point at Calibri 11). A column too narrow for one `#` shows nothing for a number, a date,
      a boolean, a literal or a text through `@*-`; the General format shows one `#`, and a
      one-character result even where its glyph does not fit (`1` for 1 and 0.5 in a 1-point
      column; 12 and -1 are `#`).
    - A `#` run is at most 255 characters (1529 points at width 254 hold 304), and a text
      that would exceed 255 characters with its `*` fill shows a `#` run instead, text too
      (`* 0` with 1 at width 86: 255 spaces and the digit make 256; `@*-` and `*-@` with `abc`
      at width 130). A number whose text exceeds 255 characters is 255 `#` without a width
      (`1E+300` under `0`, 301 digits). Read with `string value`; whether the screen shows the
      same is not yet checked.

29. **Cell font different from the Normal font (2026-10-28; suites `cell_font`, 132 cases in
    Calibri 9, 10, 12, 14 and 16 cells and Normal-style cells of a Calibri 11 workbook at every
    whole point from 11 to 100; `general_cell_font`, the `width_general` values in those cell
    fonts at the same widths; `cell_font_numeric`, every `numeric_core` case in Calibri 9 and 14
    cells at widths 6, 10 and 14, 65,664 rows; `general_zero_sign`, 8 formats with 5 negatives
    at every point from 11 to 50 at Calibri 11).** A case's `font` is set on its cells and its
    tag ends in the font's label (`[Calibri 14]`; `[normal]` for the Normal style). Probe first
    (Calibri 9, 14, 16, 11 bold, italic and bold italic cells, 13 widths): budget and units from
    the Normal font with the text measured in the cell font matched every non-General row;
    everything in the cell font or everything in the Normal font failed 111 to 172 rows per
    size. Rules, all rows matched:
    - The column (unit, points, budget) comes from the Normal font; fit, `#` and fill counts
      measure the text in the cell font: eight digits fit width 8 at Calibri 11 and are `######`
      (seven-point `#`, 48 points) in a Calibri 14 cell.
    - General counts characters against the budget in the cell font's digits (36 / 5 = 7.2 at
      Calibri 9: `1234568` at width 6) and a leading minus by the cell font's advance (5 / 8 in a
      Calibri 16 cell: `-12345678` at width 10 is `-1E+07`, the last row the half unit missed).
    - Only the plain General format, and the empty format, keep `-0` and charge the minus inside
      the General text; `[Red]General`, `General;General`, `@`, `@" x"` and `[>100]0.00`
      (reached through a failed condition) render the magnitude after a sign piece, so -0.05 at
      1.5 units is `-0` under `General` and `0` under the others, and -1.5 at width 1 is `0`
      under them (less than one character left, General's 0.0001-to-10 rule). In a Calibri 14
      cell at width 1 a value below 0.0001 gets no form under plain General either (`1E-05` is
      `#`; 6 points of a 7-point digit).
    - Autofit measures in the cell font with the same 7-point padding, so style tables can be
      measured in a workbook whose Normal style is the regular face (Calibri 11 bold: 515 of 515
      space-less texts exact). A Calibri 14 cell in a Calibri 11 workbook autofits one point
      wider than in a Calibri 14 workbook on 14 ten-runs of glyphs with a tenth (readback through
      other column units), so tables are measured with the Normal style in the same family and
      size.
    Character mode skips cell-font rows (characters of another font than the column's).

30. **Bold, italic, bold italic and Aptos Narrow (2026-10-28; 42 `measure_<family>_<style><size>`
    suites of 523 autofit texts and 42 `space_fill_...` suites of 200 rows, for Calibri bold,
    italic and bold italic and Aptos Narrow in all four styles at 9, 10, 11, 12, 14 and 16;
    validation suites `cell_font_styles` and `general_cell_font_styles`, styled and Aptos Narrow
    cells in a Calibri 11 workbook at every point from 11 to 100, and `column_file_aptos_narrow11`,
    `general_points_aptos_narrow11` and `numeric_core_width_aptos_narrow11` in an Aptos Narrow 11
    workbook).** The Normal style is the regular face of the family at the size (written without
    the theme scheme for Aptos Narrow, so Excel does not substitute the theme's Calibri; the
    header reads it back), and the cells carry the style, as item 29 found best. Results:
    - Every one of the 48 faces fits its 515 space-less texts exactly with `floor(sum) + PAD`,
      and one space advance matches all 160 of its fill rows. Bold has the regular digit width
      and padding at every size, with wider letters (Calibri 11: `A` 7 against 6, `e` 6 against
      5). Aptos Narrow's digit widths equal Calibri's at every size (5, 5, 6, 6, 7, 8); its space
      is 2, 2, 2, 2, 3, 3.
    - Italic faces autofit with a padding of 9, 11 or 13 points, and with the plain advances no
      space advance matched more than 107 of 160 fill rows; an italic cell held exactly one fill
      space fewer than an upright one at every width (Calibri 11, `0* `: 14, 26, 38, 56 against
      15, 27, 39, 57). Taking `(PAD − 7) / 2` points (the overhang: the slant reaching past the
      advances on each side) off the cell's budget matches all 160 rows of every italic face,
      the overhang unique among 0 to 6. Overhangs: Calibri 1, 2, 2, 2, 2, 3 from 9 to 16; Aptos
      Narrow 1, 1, 2, 2, 2, 3. General in an italic cell loses the overhang too (`1E-04` shows
      `0` in a Calibri 11 italic cell at width 5: 28 points, 4.67 characters).
    - The regular Calibri tables regenerated by the generalized `capture/glyph_fit.py` are
      identical to 0.1's.

31. **Default alignment (2026-10-28; `capture/alignment.py`, `tests/captures/alignment.layout.tsv`,
    196 rows).** Neither `horizontal alignment` (which reports `general`) nor `string value`
    says where Excel puts a value, so this uses Excel's PDF export, the method of sheetpeek's
    layout capture: one case per row in a column 40 file units wide (240 points at Calibri 11),
    gridlines printed, 100 % scale, so a PDF point is a sheet point; pdfminer.six (the
    `capture` dependency group, never a package dependency) gives each row's characters, and
    the run's place between the gridlines gives the alignment: left when it starts within 6
    points of the left gridline (left text starts 3 points in), right when it ends within 6 of
    the right one, center when its middle is within 2 points of the column's. No clipboard is
    used. Seven values (12345, 1.5, -1, 0, a date serial, TRUE, `abc`) under 22 formats and the
    seven classic errors under 6.
    - A number under a format with only a text section (`@`, `@" units"`, `"x"@`, `[Red]@`,
      `@@`, `_(@_)`, `*-@`, `@*-`) is left-aligned and shows General, the section's literals
      not applied; under formats with numeric sections (`0;0;0;@`, `General;General;General;@`)
      it is right-aligned, and through a failed condition (`[>100]0`) too. Booleans stay
      centered and text left under every format; dates right.
    - Errors are centered and show their own text under every format, `;;;` included.
    - The printed text is the text at the column width: 47 `#` for -1 under `m/d/yy`, fills
      across the column; the PDF draws no pad spaces. All 154 non-error rows match `text` at
      that width (spaces aside), and all 134 whose alignment can be read (text present, no fill,
      not a `#` run) match `align`.

32. **Error values (2026-10-28; suite `error_width`, 1,638 rows).** Formulas produce the seven
    classic codes (`=1/0`, `=NA()`, `=SQRT(-1)`, `="a"+1`, `=XYZZY()`, `=#REF!`,
    `=SUM(A1:A2 B1:B2)`) under `General`, `0`, `@`, `[Red]0`, `;;;`, `0;-0;0;@`, `"x"0`,
    `0.00*-` and `[Blue]General`, in Normal-style and Calibri 14 cells of a Calibri 11 workbook,
    width-less and at widths 1 to 12, with each width-less cell's display-format color. A case's
    value is a `CellError`, written with `value_type` `error`. An error shows its own text under
    every format (`;;;` included) in the automatic color (also under `[Red]0`), and at a width
    overflows to `#` like a number, by glyph width: `#DIV/0!` (35 points) shows at width 6
    (36) and is six `#` at 5; `#N/A` shows at 4; in a Calibri 14 cell `#DIV/0!` needs width 8
    and a 6-point budget holds no 7-point `#`. Every row matches. `#SPILL!` and `#CALC!` need
    dynamic-array metadata openpyxl does not write and are not captured; the rule is the same
    for every code, only the text differs.

## 3. The pipeline

```
capture/cases.py              suites: tagged Case(value, format) lists plus the widths to capture at,
                             a measure flag, the date system and the fast-tier cap
capture/build.py              openpyxl -> <Excel container>/xlformat/<suite>.xlsx and
                             capture/build/<suite>.manifest.json; sheet "cases": A = row index,
                             one value column per width; sheet "measure" for autofit
capture/capture_template.applescript   filled per suite into capture/build/<suite>.capture.applescript
capture/capture.py            runs osascript, parses the output, assigns tiers, writes the TSV
capture/tiers.py, retier.py   fast-tier rule and the command that reapplies it without recapture
capture/schema.py             the TSV schema (frozen, see below), reader and writer
capture/convert_2023.py       the 2023 clipboard capture into the schema (capture2023.tsv.gz)
capture/rounding_models.py    model comparison for the 15-digit rule
capture/glyph_fit.py          glyph tables from the measure and space_fill suites (items 26, 30)
capture/alignment.py          default alignment through Excel's PDF export (item 31)
tests/captures/*.tsv[.gz]      the suites; tests/conftest.py and tests/test_captures.py load them
```

How a capture works: `uv run python -m capture <suite>` builds the workbook, writes the
script, saves the text clipboard, runs the script, restores the clipboard, parses the output and
writes `tests/captures/<suite>.tsv` (gzipped when the file would exceed about 50 KB). The script
opens the workbook from Excel's container by path, switches on `date 1904` when the suite asks,
records Excel's version, the standard font and the workbook's Normal font, then for each
capture column either copies the column to the clipboard (width-less rendering, read back with
`pbpaste`, base64-encoded in transit) or sets the Excel column width, reads it back (the
`actual_width` column) and reads `string value of every cell` in one Apple event. With a measure
sheet it autofits every measure column and reads the widths back. It closes the workbook without
saving and restores alerts on every exit path. A 2,000-row suite takes about one second of
Excel time.

Choices that hold:

- **Values are written as numbers, never as dates**; the package's date-to-serial conversion
  is unit-tested against Excel-confirmed pairs. **Floats are written as `repr` text** (shortest
  exact decimal), not openpyxl's `"%.16g"`, and the suite header says `numbers: repr`; for suites
  captured before 2026-10-24 the test compares against the 16-digit double Excel loaded (§2
  item 16). A `Case.formula` writes a formula instead (Excel evaluates it in IEEE arithmetic);
  a `Case.written` sets the exact text.
- **Column widths are set by the script in Excel's units** and read back, because openpyxl's
  width unit differs slightly from Excel's.
- **Width-less rows come from the clipboard, width rows from `string value`.** The clipboard
  is Excel's own width-less rendering (`*` fills emit nothing, `#` runs are 255 long); `string
  value` is `Range.Text` and reflects the column. The 2023 capture used the clipboard, and the
  new clipboard capture reproduces it exactly.
- **Width suites in two flavors** (character mode with digit-only output, Excel mode with
  everything) plus autofit measurements for glyph tables; see `ARCHITECTURE.md` §8.
- **Invalid formats are captured rows or unit tests**: strings are written with openpyxl and
  opened; a refused file is a `FormatError` (unit test), a rendered one is a row to match.
- **Color** is read through `display format` one cell at a time (section 2 item 21; bulk reads
  return `missing value`); only suites with `Suite.colors` fill the column, and fifteen rows
  the API misreports are skipped by the test.
- **The 2023 capture stays**, gzipped, as a full-tier-only suite; the diff between captures is
  evidence about Excel builds and about the method.
- **Header metadata** per suite: Excel build, standard font, workbook font, locale, capture
  date, generator (the `capture.cases` function), capture method per column, tier rule. A regeneration that
  changes any expected text is a reviewed diff.

**Schema (frozen 2026-10-23).** Columns: `suite tag value_type value_repr format width
actual_width date1904 expected_text expected_color tier`; `#`-prefixed header lines carry the
metadata. Changing the schema rewrites every suite and its history, so it is a reviewed
decision, not a convenience.

`TEXT()` formulas are not used anywhere: they differ from cell display for `*`, `_x` and `####`
and never produce overflow.

The TSV carries `actual_width` (the width Excel reports after setting it; `8.43` snaps to
`8.5`). Tiers are recomputable: `uv run python -m capture.retier <suite>|--all` reassigns the
`tier` column from `capture/tiers.py` (first row per output shape per format and width, at most
15 per group; `capture2023`, `capture2023_sample` and `capture2026` are full-tier only);
captures call the same function.
