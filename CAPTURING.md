# Capturing Excel's output: regenerating and extending the suites

Section 2 records probe results in the order they were found; a later item that overturns an
earlier one notes it in both places. Section 3 describes the capture pipeline.

## 1. The 2023 capture

`tests/captures/capture2023.tsv.gz` (89,680 rows) is a 2023 clipboard capture made with an earlier
prototype that is not in this repository: openpyxl wrote one row per (value, format) pair with
the format applied to column A as a real number format, columns B and C carry `"<type> <repr>"` and
`repr(format)`; the workbook was opened in Excel for Mac, the cells selected and copied, and the
clipboard text saved. The clipboard text is the displayed text with padding spaces but without
column-width effects: `1234567.911` under `General` is in the file although it cannot fit a
column of the default width. So the clipboard is not `Range.Text`, and width needs a different
capture.

## 2. Driving Excel (Excel for Mac with AppleScript, Excel for Windows by capture jobs), and what each probe found

Excel 16.113.4's scripting dictionary (`sdef "/Applications/Microsoft Excel.app"`) has what the
capture needs on `range`:

| property | meaning |
|---|---|
| `string value` (read-only, `any`) | "Returns the text for the range": the `Range.Text` equivalent |
| `value`, `formula`, `number format` | read/write |
| `column width` | read/write, in characters, per column |
| `horizontal alignment`, `display format` | alignment and the effective display format object (its font color reflects number-format colors, item 5) |

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
8. **Rounding at 15 digits.** Workbook `probe_halfdown.xlsx`, formats `0`, `General`,
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
     2023 row, fails 8 of the 48 probe rows (all values with 15 or more integer digits);
   - ties down at 15 digits, or truncation: fail both sets;
   - round the exact value to 16 significant digits, then to 15 with ties down: fits all 48
     probe rows, fails 23 2023 rows such as `0.1234512345123455` (exactly
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
9. **Formats Excel refuses in a file.** 21 formats the dialog rejects were each written
   with openpyxl into their own workbook and opened with alerts off. Seventeen made Excel refuse
   to open the file (`open workbook` fails with "Parameter error (-50)"): `0.00e+00`,
   `0.00e-00`, `##0.0e+0`, `0.00E+00;0.00e+00`, `0;0;0;0;0`, `[Red`, `0.00"`, `[Red][Blue]0`,
   `[>1][<5]0`, `@;@`, `0;@;0`, `yyyy 0`, `0 yyyy`, `0.00E+`, `# ?/`, `[hh]:mm:[ss]`,
   `General General`. These are `FormatError`. **Confirmed by hand**: after relaunching Excel,
   double-clicking one of these workbooks brought up the repair dialog ("We found a problem
   with some content in 'inv_01.xlsx'. Do you want us to try to recover as much as we can?"),
   so `-50` is Excel refusing the file with the dialog suppressed. What the recovered cell shows
   was observed later: General, the other cells keeping their formats (item 33). Four opened
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
   a lone condition that fails renders like an invalid value, not as General (the earlier prototype,
   not in this repository, rendered General here; captured width-less in `numeric_core` as `[>1]`, `[>100]0.00` and `[<0]0.0`).
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
    text never does. The exact reduction rule was derived from this file (item 23).
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
    `_lone_zero_admitting` rule (since removed) is replaced by the general re-render step in `_render.py`.
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
    DISPLAY_FORMAT_COLOR_MISREPORTS`; a future color capture needs another source for them. Two
    more joined the list with item 60: text under `[Red]General` and text under a lone `[Red]`,
    which parses as that General. The screen has shown neither; the package colors both by the
    rule the screen shows for `[Blue]General`.
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
    division by `1/D` instead of multiplication changes nothing. Decision: the
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
    deviations and `capture/fraction_models.py` (since removed) are gone. The free search was scored against
    6,351,456 rows of 2023 captures that are not in this repository: the continued-fraction walk misses none,
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
    half (refined in items 27 and 29: the minus costs its own advance in digits, which is half
    only at Calibri 11, and a glyph test decides the last resort) (`-1234.6` and `1,000.0` both
    measure 36 points at width 6, Excel reduces the first and keeps the second; `-12345678` shows at 8.5; `-2E-07` needs 6, `-1.5E-07` needs 8); all 1,107
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
    (an openpyxl `8.43` is 7.67 units). The earlier model missed 158 bare rows and 637 among
    literals, all now matched:
    - When no reduced form fits by characters, the shorter of the rounded integer and the
      one-digit scientific form shows if its glyph measure fits the points budget: `1E+06` is
      28 points and shows at 4.67 units (28-point budget), `-1E+06` at 5.17, `1E+100` at 5.67,
      `1E-100` (a value below one: always the scientific form) at 5.33. On a tie the format
      `General` keeps the integer (`12345` at 4.67 is `#####`: 30 points) while a General
      section takes the scientific form (`( 1E+04 )` at 6.33, `1E+04` under `General;-General`
      at 4.67). This also explains the earlier model's "scientific form admitted without charging the sign"
      among literals (`-1E+05x` at width 6: 36 points of 36), which is removed.
    - A leading minus costs its own advance in digits: 3 / 6 at Calibri 11 (the earlier model's half unit),
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

28. **The column model (2026-10-28; suites `column_file9` to `column_file16`, 137 widths
    written into the file as `<col width>` by openpyxl and a column without one, 19 cases each, at the six measured
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
      identical to the earlier tables.

31. **Default alignment (2026-10-28; `capture/alignment.py`, `tests/captures/alignment.layout.tsv`,
    196 rows).** Neither `horizontal alignment` (which reports `general`) nor `string value`
    says where Excel puts a value, so this uses Excel's PDF export, the method of sheetpeek's
    [`capture/layout.py`](https://github.com/sheetmill/sheetpeek/blob/main/capture/layout.py): one case per row in a column 40 file units wide (240 points at Calibri 11),
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

33. **What the repair does (2026-10-28; `capture/repair.py`, `tests/captures/repair.layout.tsv`,
    248 rows).** The one interactive capture: two workbooks opened with alerts on, the repair
    dialogs accepted by hand. The first holds all 29 refused formats of item 9, one per row,
    each between rows of valid custom formats (`0.000`, `[Red]0.00`, `"x"0`, `#,##0.0`,
    `m/d/yy`); the second a single refused format between two controls. Four values per row
    (1234.5678, -0.5, 0, `abc`), no column widths. After the repair every cell with a refused
    format reads back with the format `General` and shows General in its 8-unit column (120 of
    120 cells match the package's `General`), and every control keeps its format and shows it
    (128 of 128); the single-format workbook behaves the same. So the repair resets only the
    offending cells, and catching `FormatError` and rendering General reproduces it.
    The dialogs, per workbook: "We found a problem with some content ... Do you want us to try
    to recover as much as we can?" (yes), then a notice offering to view the repair details or
    delete them. The details are a `recoveryLog` XML file whose only entry, for both workbooks
    (29 refused formats or one), is `Repaired Records: Format from /xl/styles.xml part
    (Styles)`. A second run of the single-format workbook read back the same cells.

34. **General with long exponents, and subnormal numbers (2026-11-03; suites
    `general_exponent3`, 31 values under `General`, `[Red]General`, `General"x"`, `( General )`,
    `General;-General` and `@` at every point from 11 to 100 at Calibri 11, 16,926 rows;
    `general_exponent_sweep`, five mantissas at exponents from 5 to 307 of either sign under
    `General` and `General"x"` at every point from 11 to 90, 51,840 rows; `subnormal`, nine values
    under eight formats, 72 rows).** The values in earlier suites with three-digit exponents were
    round (1e100, 1.5e300), so none showed this.
    - The mantissa keeps five decimals up to an exponent magnitude of 98 and four from 99,
      read from the value before the mantissa rounds: `1.23456789e98` is `1.23457E+98`,
      `1.23456789e99` is `1.2346E+99` (ten characters, so the rule is not "what fits 11"),
      `1.23456789e-99` is `1.2346E-99`, `2.2250738585072014e-308` is `2.2251E-308`. `9.99995e98`
      stays `9.99995E+98` and `9.99995e-99` rounds to `1E-98`.
    - Reducing at a width, such an exponent is charged three digits, 99 included: `1.2E-99` at
      8.67 units, where `1.23E-99` would fit by its own eight characters.
    - A negative whose scientific form has no room shows `-0` while `-0` fits and a
      five-character form does not, then the `#` run: -2.2250738585072014e-308 is `#` to 1.33
      units, `-0` from 1.5 to 5.33, `######` at 5.5 and 5.67, `-2E-308` from 5.83.
    - Below the smallest normal double every value is zero under every format, with no sign:
      `2.225073858507201e-308` shows `0` under `General`, `0.00E+00` under `0.00E+00`, `zero`
      under `0;"neg";"zero"`, and its negative `1/0/00` under `m/d/yy` where
      -2.2250738585072014e-308 shows a `#` run. A value written below that limit is read as 0.
    Character mode: 5 rows of `( General )` where Excel's narrow literals leave room for the
    longer form are skipped as glyph slack.

35. **What a bare letter means (2026-11-03; `capture/letters.py`, 351 candidates checked one
    workbook each, 73 refused; suites `letters`, `letters_upper`, the rendered ones with eight
    values, width-less and at widths 6 and 12, 6,504 rows; `letters1904`, ten formats under the
    1904 system, 80 rows).** Every letter of either case alone, doubled, after `0` and after
    `yyyy `, and the forms of the year code. Excel refuses a workbook holding 259 of the rendered
    formats while either half of them opens, hence two suites.
    - Literal: `a c f i j k l o p q r t u v w x z` in either case (`0c` shows `1235c`, `-1c`).
    - Refused wherever it stands, outside quotes, escapes, `_x` and `*x`: `n` and `N`.
    - Refused: `E` in any run without a sign and digits (`E`, `EE`, `Ee`, `eE`, `yyyy E+`), while
      `e` and `ee` are the year.
    - `b` and `bb` show the last two digits of the Buddhist year (Gregorian + 543) and longer runs
      all four, in either case: 12.5 under `abc` is `a43c`, under `bbbb` it is `2443`; also under
      the 1904 system. After any `[$...]` tag it is the plain year (`[$-409]bbbb`, `[$€]b`,
      `[$-41E]b`), before one it is not (`b[$-409]`). A date code: digits with it are refused
      (`0 b`, `b0`), and a negative shows `#`.
    - **The clipboard gives the plain year** for all these width-less rows (`03` where the sheet
      shows `46`), under both date systems; `string value` at any width and `TEXT` show the
      Buddhist year. The package follows the sheet and the capture test compares these rows with
      the plain year code.
    - A format whose last character is a lone code `B` is refused: `B`, `bB`, `BbB`, `yyyy B`,
      `0;B`, `[Red]B`. `BB`, `BBB`, `Bb`, `bBB`, `B;0`, `B_)` and `B yyyy` render.
    - `g` in any run and case is a date code that shows nothing in en-US (`yyyy g` is `1903 `; a
      negative shows `#`; `0g` and `@g` are refused).
    - `b1` and `b2` (either case) print nothing, before every other code of the section: `b1`
      shows nothing and `-` for a negative, `b1 x` shows ` x`, `b1@` and `b1General` show what
      `@` and `General` do. After another code, or with digit placeholders, they are refused
      (`yyyy b1`, `"x"b1`, `b1 0`, `b10`, `b1b2`), and so is any other digit after `b` (`b3`,
      `bb1`). `b1bbbb` is the Buddhist year.
    - Left out of the suite as localization, which is out of scope (locale tags are parsed and
      ignored): `b2yyyy` is the Hijri year (1321 for 1234.5678), `b2bbbb` that year plus 543,
      `[$-1070000]yyyy` the Buddhist year and `[$-D07041E]yyyy` the same in Thai digits.
    - Found on the way: a clock section reached as the negative section shows a `#` run
      (`0;yy`, `0;h:mm`, `0;m/d/yy`, `x;yy` with -0.5), and an elapsed one the magnitude (`0;[h]`
      shows `12`). `General` in a date section reduces at a width as among literals (`General
      yyyy` with 45000.75 is `0.499999803 2023` width-less, `0.5 2023` at 12, `0 2023` at 6).

36. **Values too large for fraction and elapsed arithmetic (2026-11-03; suite `huge_values`, 20
    formats with values from 1e15 to the largest double of either sign, width-less and at width
    12, 360 rows).** A fraction's whole part and an elapsed count print 15 significant digits and
    then zeros (`# ?/?` with 1e100 is `1` and 100 zeros, `[h]` is `24` and 100 zeros); past 255
    characters they are `#` runs. An elapsed count past the largest double is 0, judged per unit:
    1e304 under `[s]` is `0` and under `[ss]` `00`, under `[h]` (2.4e305) and `[m]` it is a `#`
    run; the largest double is `0` under all of them, and its negative `-0`. `?/?%` with the
    largest double is `0/1%`, `%` having scaled it past the limit, and `?/10` is `0/10`, the
    product with the denominator having done so.

37. **The empty text, and the numerator's limits (2026-11-03; suites `empty_text`, 19 formats in
    cells holding `=""`, width-less and at widths 6 and 12, 57 rows; `fraction_percent`, 15 formats
    and 32 values, width-less and at width 12, 960 rows).** The `TEXT` capture of item 38 showed
    both, and these show the cell does the same.
    - A real empty text is ordinary text: `@" units"` shows ` units`, `"<"@">"` shows `<>`,
      `_(@_)` two spaces, a fourth section `"text"` its literal, and `@*-` and `*-@` a run of
      dashes at a width. `@` and `General` show nothing. Every earlier `""` row was a blank cell,
      as openpyxl writes the empty string so, and the capture test reads such a row as `None`.
    - A fixed denominator's numerator shows up to 32767 and then a `#` run, with or without a
      percent and whatever the placeholders: 3276.7 under `?/10` is `32767/10` and 3276.8 is `#`,
      32.767 under `?/10%` is `32767/10%` and 32.768 is `#`. A free denominator's reaches at least
      1073741824/5 and 2147483647/1 is `#`.

38. **The `TEXT` function (2026-11-03; suites `text_function`, every format of every other suite
    with the empty format and the formats Excel refuses in a file, under 29 values, 24,737 rows;
    `text_function1904`, 24 date formats in a 1904 workbook, 696 rows).** Each cell holds
    `=TEXT(inputs!$A$n,inputs!$B$n)`, its value and format read from a second sheet, so no literal
    is rounded to 15 digits, no format needs quoting, and a refused format only makes the cell an
    error instead of the file unopenable (`Case.text_function`; the tag starts with `text:`, the
    schema is unchanged). The result is text in a General cell, read from the clipboard, so the
    clipboard's plain year of item 35 does not arise. Every row is the width-less rendering except:
    - A blank value is 0 (`0.00` shows `0.00`, `m/d/yy` shows `1/0/00`).
    - An error value passes through, under a refused format too (`#N/A` under `0n`).
    - A format Excel refuses in a file gives `#VALUE!`, and so does every value whose rendering is
      a `#` run (a negative date, 1e300 under `0`). A literal `#` is not one: `1` under `"#"` is `#`.
    - The empty format gives the empty text for a number and `-` for a negative one
      (-0.0004 included), and leaves text and booleans as they are.
    Text reading as a number is not part of the corpus: `TEXT("12","0.00")` converts first, which a
    caller of `text_function` does itself. It found the crash of item 36 and the rules of item 37.

39. **General that is not the General format, in narrow columns (2026-11-03; suites
    `general_text_narrow`, 15 values under 8 formats at every point from 1 to 30 at Calibri 11,
    3,600 rows; `general_sign_narrow`, 71 negatives under `@` and `[Red]General` at 1, 1.33 and 1.5
    units, 426 rows).** A probe of merged cells showed single columns behaving this way too.
    - A column too narrow for one `#` shows nothing for a number through `@`, `@*-`, `*-@`, `_(@_)`,
      `@" x"`, `[Red]General` or `General;General`. Only the plain `General` shows one `#`.
    - At 1 to 1.33 units, where a minus leaves less than a character, the value decides: `0` from
      0.0001 to under 0.5 (-0.49) and from 1 to under 10,000 with a leading digit under 5 (-1, -4.5,
      -49.5, -100, -1000); nothing, so a `#` run, for -0.00009, -0.5 to -0.99, -5 to -9.9, -50, -999,
      -9999 and from -10000 on. At 1.5 units the form fits (`-1`, `-9`) until `-9.5` rounds to two
      digits. The earlier rule, "from 0.0001 to under 10", had been fitted to -1.5 and -0.05.

40. **Merged cells (2026-11-03; suites `merged11`, `merged14`, 33 blocks of one to five columns
    written into the file as `<col width>`, wide and narrow, mixed, with a hidden column, 90 cases
    each, Calibri 11 and 14, 2,970 rows each; `Suite.merged`, the width column holding the block's
    total points).** `string value` of the first cell of a merged range formats it for the whole
    range. Every row matches one column of the summed points through
    `column_metrics_from_points`: the padding of 5 is charged once, General counts characters
    against that width, and the narrow rules apply to the sum (two 1-point columns behave as one
    of 2 points, and six of them as one of 6). The blocks include single columns of the same point
    widths as controls, and a hidden column adds nothing (`(0, 5)` and `(5, 0)` are 30 points).

41. **Fit and fill use whole advances (2026-11-03; suites `advance_tenths11`, `advance_tenths12`,
    runs of 1 to 30 glyphs and `*` fills of the glyphs that carry a tenth in the tables (`W @ M m`
    at Calibri 11, `W % w` at 12) and of a whole one (`H`) at every point up to 400 and every 37 to
    1,500, 41,481 rows, in about 11 minutes).** The tables hold 10.1 for `W` at Calibri 11 because
    the autofit of ten of them is 101 points. Decisions do not use that tenth: in a 111-point column
    (budget 106) ten `W` and a digit show, where 101 + 6 would overflow, and `*W0` holds ten `W`,
    not nine. Every row matches the sum of the floored advances; only autofit keeps the tenth. Excel
    draws them whole as well, 10 points apart in a PDF export at Calibri 11 (the `WWWWWWWWWWWW` rows of
    sheetpeek's [`tests/captures/layout.tsv`](https://github.com/sheetmill/sheetpeek/blob/main/tests/captures/layout.tsv)). Character mode skips these suites: the glyphs are wider than digits.

42. **A point and zeros after a date or time code (2026-11-03; `capture/elapsed_forms.py`, 74 forms
    one workbook each, 18 refused; suite `elapsed_decimals`, 59 forms under 13 values, width-less
    and at widths 8 and 30, 2,301 rows).** `[s].00` raised `FormatError` although Excel renders it.
    - The zeros are fractions of the second, whatever the unit: 1234.5678 under `[h].00` is
      `29629.92`, under `yyyy.00` `1903.92`, under `s x.00` `37 x.92`, under `AM/PM.0` `PM.9`, and
      `s.00.0` is `37.92.9`. Literals may stand between (`[s]_).00`, `[s]".".00`). `m` before the
      fraction is a minute (`m.00` is `37.92`). A negative under an elapsed code with a fraction is
      a `#` run, where `[s]` shows `-43200`; `[s].` keeps the sign, its point being a literal.
    - Refused: more than three zeros (`[s].0000`), another placeholder (`[s].#`, `[s].0#`), an
      escaped point (`[s]\.00`), a point and zeros before any code (`.00 s`), and beside an elapsed
      code `%`, `AM/PM` and `/` before a digit.
    - `[h]h` and `[s]s` are stored as `[h][h]` and `[s][s]` and show the count twice (`3636.0`).
    - Not modeled, left out of the suite: `d [s].00` and `yyyy [s].00` show numbers no rule here
      explains (`0 4363200.50` for half a second).
    - A batch stopped with `AppleEvent timed out` (-1712) on the workbook holding `[s]*-.00`. Excel
      did not answer scripts for some time; when looked at later no dialog showed, that workbook
      was open and alerts were on, and it answered once the workbook was closed by hand. Whether
      the format is the cause is not established. `capture/probe_file_many.applescript` now limits each workbook to 30 seconds and stops at the
      first timeout, naming the format. The form has no capture.

43. **Every suite again, on Microsoft 365 (2026-11-12; Excel 16.115 (26110837), Microsoft 365
    Subscription, update channel Current; every suite but the 2023 capture).** The suites before
    this were captured on Excel 16.114 under Office Home & Business 2021; both the build and the
    edition changed, so a difference could have come from either. None did: every row of all 166
    suites (597,582 rows, about 35 minutes of Excel) has the same text, color and reported column
    width, the autofit widths of the `measure*` suites are the same, the glyph tables refit from
    them are unchanged, the alignment layout of the PDF export is the same on all 196 rows, and
    so is what the repair leaves on all 248 rows of `repair.layout.tsv`.
    - Excel's standard font is now Aptos Narrow 12, where it was Calibri 12. The suites set the
      workbook's Normal font (Calibri 11 unless a suite says otherwise), which is what widths
      follow, so no row moves; the header's `standard_font` records it.
    - Setting a column's width is faster: `advance_tenths11` and `advance_tenths12` took about 55
      seconds each, where they took about five minutes.
    - Refitting showed a fault in `capture/glyph_fit.py`: it scored every candidate space advance
      with the first table, a cached column measure surviving the table's change, so every
      candidate tied and the lowest won. With that cache cleared each face matches all 160 of its
      fill rows, and the tables come out as committed.
    - `text_function` grew to 1,010 formats (29,290 rows): its corpus is every format of every
      suite at build time, so it takes in the suites added since its first capture, and it now
      also holds every refused form of `capture/letters.py`, `capture/elapsed_forms.py` and the
      parser tests. Each still gives `#VALUE!`, so no refusal changed and none needed a file.
    - **Refusals after a build change** are re-checked this way: one workbook of `TEXT` cells
      settles that a format is still refused; only a format `TEXT` now accepts is opened in a file.
      Files hold at most about 120 custom formats: Excel takes between 200 and 250 per workbook
      (item 35's workbook of 259 was refused while either half opened), so batches of a thousand
      formats, as a formula probe might use, would fail on the count alone.
    - **The repair dialog** comes up after an AppleScript open has returned, so the workbook is
      not there to read until a person has answered, and a request sent while a person answers
      a dialog can crash Excel. `capture/repair.py` therefore opens each case through Launch
      Services and exits; once the dialog is answered (Yes) and `read` is run, it finds the workbook
      by name (a repaired workbook keeps its file's name, its window's caption ending
      "- Repaired") and reads it. A second open and read of the same case gave the same cells.
      Excel was restarted before this capture, since one capture can leave Excel altered for the next.
    - **Headers** now record the full build (`excel_build`), the edition and update channel
      (`edition`; `capture/edition.py`, which refuses to run until the edition is stated in
      `capture/build/excel_edition.txt` or `XLFORMAT_EXCEL_EDITION`) and whether the display was
      kept awake (`display`). Requests go through `capture/excel_watch.py`: a window of Excel's
      own (the About window, a dialog) stops a run before its first request, a request past its
      time limit saves evidence under `capture/build/excel_hangs/` and closes its workbook once
      Excel answers, and the long captures run under `caffeinate -d -i`.

44. **Row heights (2026-11-14; `capture/row_heights.py`, 48 workbooks, one per Normal font:
    Calibri and Aptos Narrow at 9, 10, 11, 12, 14 and 16 points in four styles, 147 rows each,
    7,056 heights and 48 standard heights in `row_height.layout.tsv.gz`, 7,488 since (below); five minutes;
    Excel 16.115 under Microsoft 365, restarted first).** Excel fits each row when it opens the file.
    - **The font height** F of a face, the height of a row in that Normal font, does not depend on
      bold or italic: Calibri 9 to 16 give 12, 14, 15, 16, 19 and 21 points, Aptos Narrow 13, 14,
      15, 16, 19 and 22. The sheet's standard height and a row without an element are the Normal
      font's F (the sheet keeps openpyxl's `<sheetFormatPr>`, without `customHeight`).
    - **A row is as tall as its tallest font:** the larger of the Normal font's F and that of every
      cell font in the row, an empty cell styled with a font included, in all 48 workbooks.
    - **Wrapped text** (one to six lines joined by line feeds, wrap on): in the Normal font, k lines
      are k(F + 1) points, so even one wrapped line is a point taller than the row. In another cell
      font, k lines are the larger of the Normal font's F and k(F_cell + 1) + c, where c depends
      on the Normal font and the cell font, not on their styles, and is p(normal) - p(cell) or 0,
      whichever is larger, with one whole number p per face: 0 for Calibri 9; 1 for Calibri 10, 11
      and 12 and Aptos Narrow 9 to 12; 2 for Calibri 14 and 16 and Aptos Narrow 14; 3 for Aptos
      Narrow 16. The rule matches all 3,456 wrapped rows and all 144 pairs of faces; c is 0 for
      102 of them. p behaves like a part of the Normal font's line that a smaller font's line
      does not fill, the descent being the likely candidate.
    - **Wrapped text beside another font** (2026-11-15, eight more rows a workbook, 7,488 heights in
      all; the 7,104 earlier ones repeated exactly): the other cell's font counts as the Normal
      font's does, both for the row's height and for p. k lines in face f beside unwrapped cells
      are the largest of every font's F and k(F_f + 1) + max(0, P - p_f), P the largest p of the
      Normal font and the row's other cell fonts; all 384 such rows match. Several cells with
      wrapped text in one row are not captured.
    - **18, 20 and 24 points** (2026-11-15, 72 workbooks of 222 rows, 15,984 heights; the 7,488
      earlier ones repeated exactly): F is 24, 26 and 31 at Calibri and 24, 27 and 32 at Aptos
      Narrow, in every style, and p is F - size - 3 as item 47 found (3, 3, 4 and 3, 4, 5); the
      rules above match every row.

45. **A sheet that hides zeros (2026-11-15; probe suites `hidden_zeros`, on a sheet with
    `<sheetView showZeros="0">`, and `hidden_zeros_control`, the same 425 cases on a sheet that
    shows zeros: 25 formats under 17 values, width-less and at width 12, 850 rows each; Excel
    16.115 under Microsoft 365).** The control matches the package on every row, and the hiding
    sheet matches `show_zeros=False`. The case `=0.1+0.2-0.3` holds 0, the value Excel stores for
    it (its first capture, as a probe, recorded the Python double and gave the same texts). On the
    hiding sheet:
    - **A cell whose value is exactly 0 shows nothing, under a format without a zero section**
      (fewer than three sections): `General`, `0`, `0.00`, `0%`, `0.0E+00`, `# ?/?`, `m/d/yy`,
      `h:mm`, `[h]:mm`, `@`, `0;-0`, `0;"neg"`, `"x"0`, `[Red]0`, and conditional formats whose
      condition takes the zero (`[=0]"z";0`, `[<1]"small";0`), width-less and at a width alike.
    - **A format with a zero section shows it as usual:** `0;-0;"zero"`, `0;-0;;@`,
      `General;General;General`, `#,##0;(#,##0);"-"` and the accounting format are unchanged.
    - **Only the stored value counts:** 0.4 under `0` still shows `0`, and 1e-10, the text "0" and
      the booleans are unchanged. -0.0 and 5e-324 are loaded as 0 and hidden, and so are formulas
      giving 0, `=0.1+0.2-0.3` included, which Excel calculates as exactly 0.
    One rule matches all 1,700 rows: render the value as usual, except that a number equal to 0
    under a format of fewer than three sections becomes the empty text.

46. **Vertical placement in the print layout (2026-11-15; `capture/vertical_pdf.py`, a probe:
    two workbooks, Normal font Calibri 11 and Aptos Narrow 11, 168 rows each of "Hxg" in the 12
    regular faces aligned top, center and bottom in rows of the face's F, F + 10 and 2F, rows Excel
    fits, two and three wrapped lines, and the styles at 11 and 16; Excel's PDF export read with
    pdfminer.six, baselines from the text origins; written to `capture/build/probes/`).**
    - **Placement depends on the face and the row's height only.** Every row of a stated height is
      placed the same under both Normal fonts, and bold and italic place as regular.
    - **Two constants per face:** with top alignment the baseline is a fixed distance below the
      row's top (Calibri 9 to 16: 9, 10, 10, 11, 13, 15 print points; Aptos Narrow: 8, 9, 10, 11,
      13, 15), with bottom alignment a fixed distance above its bottom (Calibri: 4, 4, 4, 4, 5, 5;
      Aptos Narrow: 4, 4, 4, 4, 5, 6), at every row height tried; centered text sits midway between
      the two placements, to half a point.
    - **The print layout is not the screen:** rows print smaller than their height (15 points as
      14, 25 as 23, 30 as 28) and wrapped lines closer (14 points apart at Calibri 11, where the
      sheet gives 16), so these distances are print points. The screen gives the sheet's.

47. **Vertical placement on screen (2026-11-15; `capture/screen_vertical.py`, a probe: the cases
    of item 46 with column A filled magenta and cyan in turn, Excel's window captured at zoom 100
    on a 2x display, 10 views, 336 rows in `vertical.layout.tsv`, every measured distance in
    pixels; the images stay in `capture/build/`).** Measured from each
    row's top in sheet points, the baseline being the bottom of "H":
    - **Top-aligned text has its baseline the face's size below the row's top** (Calibri 9 at 9
      points, Aptos Narrow 16 at 16), at every row height. Item 54 corrects this and the gap G
      below by 1 point: the origin used here was 2 pixels below each row's own top.
    - **Bottom-aligned text has its baseline a gap G above the row's bottom:** Calibri 9 to 16 4,
      4, 4, 4, 5, 5; Aptos Narrow 4, 4, 4, 4, 5, 6. G is F - size except at Calibri 9 (4, where
      F - size is 3), and the row model's p is F - size - 3 for every face.
    - **Centered text sits midway**, to half a point; **wrapped lines are F apart** (Calibri 11:
      15 points), the last one placed as a bottom-aligned line.
    - **The Normal font and the style move nothing**, with one exception that follows the row
      model: wrapped text in a face whose p is below the Normal font's sits c = p(normal) - p(face)
      higher, the extra height of item 44 lying below the last line (Calibri 9 under Calibri 11:
      1 point). All 336 rows match these rules to half a point.
    - The print layout of item 46 gives the same gaps and order with smaller distances; the screen
      is the sheet's measure.
    - **The ink** of "H" above the baseline and of "g" below it depends on the family and the
      size only (Calibri 11: 7 and 2 points; Aptos Narrow 16: 10.5 and 3), in every style.
    - **18, 20 and 24 points** (2026-11-15, 480 rows; the 336 earlier ones repeated, the band
      columns aside, which depend on where a row falls in a view): the baseline is the size below
      the top, G is F - size (Calibri 6, 6, 7; Aptos Narrow 6, 7, 8), and the ink is 11.5/3,
      13/3.5, 15.5/4.5 (Calibri) and 12/3.5, 13.5/3.5, 16/4.5 (Aptos Narrow).
    - Traps of the method: Excel draws no gridline between cells of one fill, so the marker
      column alternates two colors; the screen draws cyan as about (117, 251, 253); and italic
      glyphs touch, so a single ink run gives H's foot from its leftmost columns and its top
      from the run's highest ink.

48. **18, 20 and 24 points (2026-11-15; the `measure*` and `space_fill*` suites at those sizes in
    both families and four styles, 48 suites; `column_file18` to `24`, `column_file_aptos_narrow18`
    to `24` and `general_points18` to `24`, 35,612 rows; Excel 16.115 under Microsoft 365).**
    - **A column's padding grows with the digit width:** the default column (8 units) is 79, 87
      and 103 points at Calibri 18, 20 and 24, eight digits plus 7, where it is eight plus 5 up to
      16 points. `2 * ceil(MDW / 4) + 1` gives 5 for a maximum digit width of 5 to 8 and 7 for 9
      to 12, and with it every row of the new suites matches: dialog widths, budgets, `#` runs,
      fills and General. Autofit pads by the column's padding plus 2, so the regular faces at the
      new sizes have no overhang (their autofit padding is 9), and the italic overhang is 3, 3 and
      4 points at 18, 20 and 24.
    - **The glyph tables** at the new sizes fit every autofit text (515 of 515 per face) and every
      fill row (160 of 160); the digit is 9, 10 and 12 points and the space 4, 5, 5 (Calibri) and
      3, 4, 4 (Aptos Narrow). The tables at 9 to 16 come out unchanged.
    - **One file width** converts to a point more than ECMA-376's formula gives: 10.7109 at a digit
      width of 12 is 129 points, 128 by the formula; the 136 other widths at 24 points follow it, and
      rounding the width to 1/256 or `128 / MDW` instead of truncating breaks other rows. A known
      deviation (`QUIRKS.md`). Overturned by item 53: the width written was 10.7109375, which the
      formula converts to 129 points as Excel does; the suite file recorded it as "10.7109".
    - Two faults of the fitting script surfaced: it wrote an italic face's overhang only when not
      zero, so a style listing only the new sizes made the sizes below scale from them; it now writes
      every measured size. And it took the autofit padding beyond 7 as overhang, which held only
      while the column's padding was 5.

49. **The built-in number format ids (2026-11-16; `capture/builtin_formats.py`, a probe: every id
    from 0 to 163 as a style naming the id with no format string, as Excel writes a typed date,
    under a date and time, 1234.5678, -1234.5678, 0.25, 0, 1.5 and a text; Excel 16.115 under
    Microsoft 365, en-US; written to `capture/build/probes/`).** Excel opens every id; `number
    format` and `number format local` agree for all of them, and the package renders every one of
    the 1,148 texts from Excel's own format string.
    - 0 to 22, 37 to 40, 45, 46, 48 and 49 are the strings ECMA-376 lists, except **14, which is
      `m/d/yy`** (a typed 1 January 2024 shows `1/1/24`), and 5 to 8, the dollar formats, which
      read `$#,##0_);($#,##0)` and the like with the dollar unquoted.
    - 41 to 44 are the accounting formats, the dollar unquoted and the parentheses plain; 47 is
      `mm:ss.0`.
    - 23 to 26 are `General`; 27 to 31 and 36 are `m/d/yy`; 32 to 35 are `h:mm:ss` (ids defined for
      East Asian locales, shown in en-US as these).
    - 50 to 58 are `m/d/yy`, 59 to 81 repeat 1 to 8, 9, 10, 12 and 13, then `m/d/yy` twice, 15
      to 17, 20, 21, 22, 45, 46 and 47; 82 to 163 are `General`.
    - **The sources:** ECMA-376 Part 1 §18.8.30 lists 14 as `mm-dd-yy`, 22 as `m/d/yy h:mm` and 47 as
      `mmss.0`; Microsoft's [MS-OI29500] §2.1.712 says Excel defines 14 as `m/d/yyyy`, 22 as
      `m/d/yyyy h:mm`, 37 to 40 with `_)` and 47 as `mm:ss.0`, for en-US. Excel for Mac agrees with
      Microsoft at 37 to 40 and 47 and differs at 14 and 22 (`m/d/yy`, `m/d/yy h:mm`); neither
      source says an id follows the system's settings.
    - openpyxl's table differs at 14 (`mm-dd-yy`), 47 (`mmss.0`) and 44 (its sections lack the
      semicolons); its 5 to 8 and 41 to 43 differ only in quoting and escapes, rendering the same.
    - **macOS's settings** (2026-11-16; ids 5 to 8, 14 to 22 and 37 to 47 captured three more
      times, Excel restarted before each, one change each): (a) the United States region with a
      custom short date `y-MM-dd`: id 14 shows `2024-01-01` and 22 `2024-01-01 13:30`, their `number
      format` still `m/d/yy` and `m/d/yy h:mm` and their `number format local` `yyyy-mm-dd` and
      `yyyy-mm-dd h:mm`, and 15 and 16 pad the day (`01-Jan-24`, `01-Jan`); nothing else changes.
      (b) The United Kingdom region with its default formats (`en_US@rg=gbzzzz`): nothing changes,
      dates and the dollar alike. (c) The United States defaults restored: every id as first
      captured. What a cell under 14, 15, 16 or 22 shows depends on the viewer's custom short date.

50. **Section colors under a workbook palette (2026-11-16; `capture/screen_palette.py`, a probe:
    a workbook whose `<indexedColors>` holds 64 colors far from the default ones, and a control with
    the default palette; one row per color name and per `[Color1]` to `[Color56]`, bold 20-point
    text on a brown fill; window captures converted to sRGB; each cell's text color its most
    frequent pixel away from the fill, at least 2,080 pixels a cell; Excel 16.115 under Microsoft
    365).** Every measured color is within 3.2 of one candidate (Euclidean distance in sRGB):
    - **`[ColorN]` is drawn in the workbook's palette entry N + 7**, all 56 of them.
    - **The eight names (`[Black]` to `[Cyan]`) are drawn in their standard colors**, `FF0000` for
      `[Red]` and so on, whatever entries 0 to 7 of the palette hold, as the screen shows where
      Excel's scripting reports the palette's color.
    - With the default palette every cell is the default color.
51. **Line breaks and tabs in text (2026-11-16; suites `line_breaks` and `line_breaks_aptos_narrow`:
    `first`/`second`, `0`/`0` and `ab`/`cd` joined by 1, 2, 5 or 10 carriage returns, line feeds,
    CR LF pairs or tabs, as formula cells so that each character is the one intended (the XML
    parser turns a stored CR into LF), beside the joined text and a stored LF; under `@`, `@* `
    and `@*0` at widths 8, 12, 20 and 30, with each case's autofit width; Calibri 11 and Aptos
    Narrow 12 workbooks; Excel 16.115 under Microsoft 365).**
    - **A line feed and a tab take the space's advance; a carriage return takes none.** So CR LF
      measures one space. Every fill count and every `#` in the 1,296 rows matches that, where
      measuring them as a digit failed 335 rows in Aptos Narrow alone.
    - Autofit agrees: in Calibri 11 `firstsecond` fits 57 points and each added LF, CR LF or tab
      adds about 2.4 (81 for ten); in Aptos Narrow 12, 61 points and about 2.2 for each (83 for
      ten). Any number of CRs leaves Aptos Narrow's autofit at 61. In Calibri 11 it is one point
      narrower than the joined text (56, for every count). No decision uses autofit, so this is not
      modeled.
    - The capture runner read osascript's output in Python's text mode, which turns CR and CR LF
      into LF. It now decodes the bytes itself. `text_overflow`, recaptured as a control,
      came back with every row unchanged.
52. **Column-width edges (2026-11-17; 20 suites, 8,098 rows: `column_default_base6_*`,
    `base10_*` and `base12_*` at 9, 11, 14, 16, 18 and 24 points where not captured before, a column
    without `<col>` under that `baseColWidth` beside file widths 8, 8.43 and 10; `column_boundary11`,
    `16` and `24`, 120 file widths each, 50 written to four decimals just below a 1/256 step, 50 just
    above and 20 on a step; `long_fills`, six fills at widths 40 to 255 in a Calibri 9 workbook;
    Excel 16.115 under Microsoft 365).** Every row matches the model as it stands: a column without
    `<col>` is `baseColWidth · MDW + P` points at every base width and size; a file width is
    truncated to 1/256 before ECMA-376's formula on both sides of a step; a fill whose text would pass
    255 characters shows a `#` run.
53. **The width that was not a deviation (2026-11-17; `column_boundary11`, `16` and `24` recaptured,
    every expected text and read-back width unchanged).** The capture wrote a width's text with six
    significant digits (`f"{width:g}"`), so a file width of 10.7109375 (a 1/256 step) was recorded
    as "10.7109", and the test converted that text: 128 points at a digit width of 12, where Excel's
    129 is what ECMA-376's formula gives for the width written. The known deviation of item 48 was
    this, not Excel. Width texts are now exact (`repr` where six digits lose the value; header
    `widths: exact`), and for older suite files the test takes the 1/256 step whose six-digit text
    the row holds. The three `column_boundary` suites, whose widths just off a step could share a
    six-digit text with the step itself, were recaptured with exact texts.
54. **The row's top on screen (2026-11-17; no new capture: item 47's images and
    `vertical.layout.tsv`, and the window images of sheetpeek's
    [`text_style`](https://github.com/sheetmill/sheetpeek/blob/main/tests/captures/text_style.tsv)
    capture, group `align`: top- and bottom-aligned "Hxg" in 60-point rows; both Excel 16.115; the
    window images are in neither repository, and this file's `band_shift` column carries the
    result).** Item 47 took each row's top from the view's first marker band
    plus the row heights. That band starts below the column header's separator, while every other
    band starts on the top pixel of its own top gridline (the file's `band_shift` is -2 pixels for
    all 464 such rows): a row's fill covers the gridline above it, so a row runs from the top of that
    gridline to the top of the next. Measured against the gridline pixels directly in sheetpeek's
    images, top-aligned Calibri 11 has its baseline 12 points below that top and bottom-aligned 3
    above the bottom; item 47's frame was 1 point low throughout. Corrected: top-aligned text has
    its baseline the face's size plus 1 below the row's top, and the gap G is one less (Calibri 9 to
    24: 3, 3, 3, 3, 4, 4, 5, 5, 6; Aptos Narrow 3, 3, 3, 3, 4, 5, 5, 6, 7). Line pitch, centering,
    cap height, descender and the Normal-font exception are distances within a row and stay. The
    test applies the 2-pixel origin to the file's rows; all 480 placements match.

55. **Arial (2026-11-19; 72 `measure_arial[_<style>]<size>` and `space_fill_...` suites in four
    styles at 9, 10, 11, 12, 14, 16, 18, 20 and 24; `column_file_arial9` to `24`,
    `general_points_arial10` and `11`, `numeric_core_width_arial10` in Arial workbooks;
    `cell_font_arial` and `general_cell_font_arial`, Arial cells in a Calibri 11 workbook at every
    point from 11 to 100; `row_height.layout.tsv.gz` recaptured with Arial as a third family, 108
    workbooks of 324 rows; 121,000 cells and 35,000 heights in 32 minutes; Excel 16.115 under
    Microsoft 365, which bundles Arial's four faces).** The Normal style is written without the
    theme scheme, as for Aptos Narrow (item 30), and the headers read Arial back.
    - **Widths:** every face fits its 515 space-less texts with `floor(sum) + PAD` and all 160
      fill rows with one space advance, as Calibri's and Aptos Narrow's do. The digit is 5, 6, 6,
      7, 8, 9, 10, 11 and 13 points from 9 to 24 (wider than Calibri's from 10), the space 3, 3, 3,
      3, 4, 4, 5, 6, 7, and the italic overhang 1, 1, 2, 2, 2, 2, 3, 3, 4. A digit of 13 gives
      a column padding of 9, as `2 * ceil(MDW / 4) + 1` predicts: the default column of an Arial
      24 workbook is 8 digits plus 9.
    - **General by glyphs:** Arial's `E` (7 at 10 points) and `+` (6) are at least a digit wide,
      where Calibri's are narrower. A General form that fits by characters can then overflow by
      glyphs: `1234567.891` in an Arial 10 file column of 5.875 (5 units, 30 points of text)
      shows `#####`, where `1E+06` measures 31; also at 5.75 and 5.875 in Arial 24, and `1E+100` at
      width 6 in `numeric_core_width_arial10`. So the glyph test applies to the General format as
      well, except for the one character it keeps with less than one left (`'1'` in a 1-point
      column, `column_narrow`). Every earlier suite agrees with this.
    - **Row heights:** the Arial faces are 12, 13, 14, 16, 18, 20, 23, 25 and 30 points from 9 to
      24, in every style. All 15,984 heights of the earlier capture came back unchanged. Of the
      rows mixing families, 11 pairs of faces, each Arial with a Calibri or Aptos Narrow face
      just under it, are a point taller than the taller face (Arial 11 with Calibri 10, 14 points
      each: 15; Arial 24 with Aptos Narrow 20: 31); Calibri with Aptos Narrow never is. A row takes
      its fonts' largest ascent plus largest descent, rounded up, where a face's ascent is its
      height less its descent. With the fonts' own `hhea` descents rounded every way, 60 rows or
      more miss; a descent per face in quarter points matches every row, and the table kept is the
      one nearest the fonts' descents (Arial 2 to 5.5 points from 9 to 24, within 0.71 of
      `434 / 2048` of the size). A wrapped cell's font counts too: one wrapped line of Aptos
      Narrow 9 under Arial 11 is 15 points. Arial's line insets are 0, 0, 0, 1, 1, 1, 2, 2, 3.
    - **Placement** (2026-11-19, `capture/screen_vertical.py` with Arial 10 as a third Normal font
      and Arial among the faces, 1,080 rows; the 480 earlier rows reproduced exactly, the band
      columns aside): top-aligned Arial has its baseline the size plus 1 below the row's top, as
      the other families do. Bottom-aligned, Arial 9 to 24 sit 2, 2, 2, 3, 3, 3, 4, 4, 5 points
      up, but Arial 9 to 11 sit 3 up in the Calibri 11 and Aptos Narrow 11 workbooks: a
      bottom-aligned baseline is at least the Normal font's distance up. Calibri 9 shows the same
      in the Arial 10 workbook (2 points, where every earlier workbook's Normal font gave 3, so
      item 54's 3 for Calibri 9 was Calibri 11's floor). The ink of "H" and "g" is 6.5/2, 7.5/2,
      8/2.5, 9/2.5, 10/3, 11.5/3, 13/4, 14.5/4, 17.5/5; at 11 and 16 points the descender is half a
      point shorter or longer by the alignment (the baseline's pixel position), in every workbook.
      Four rows were read where a view's last row lay under the sheet-tab bar; the script now
      starts each view at the previous view's last row, and the test lists the four.

56. **Padding and fills in angled text (2026-11-19; `capture/angled_formats.py`: 16 formats of
    "Total", 1234.5 and 12345678901, with `_)` padding, `*` fills of spaces and dashes, the
    built-in accounting format or neither, turned 15, 30, 45, 60, 75 and 90 degrees and -45 and
    -90, in Calibri 11 at the default width; 128 rows Excel fits on opening, heights read through
    AppleScript, in `angled_formats.layout.tsv`).** A fitted row of angled text grows with the
    text's width (`W·|sin θ| + S·|cos θ| + 7`), so the rows show which text Excel turns. At every
    angle, formats whose width-less texts are equally wide get the same row, and the rows rise
    with that width only: `_)` adds its space, a fill adds nothing (`0.0* `, `* 0.0` and `0.0*-`
    fit as `0.0` does, where a fill expanded to the 53-point column would be far wider), and
    12345678901, a `#` run upright in that column, is turned whole. At 90 and -90 degrees every
    row is exactly the width-less text's width plus 7 (`0.0` 40 points, `0.0_)` 42), and each
    negative angle fits as its positive one. So angled text is the width-less rendering, as
    `format_value` gives it without a width (`tests/test_angled_formats.py`).

57. **A face's height at any size (2026-11-19; `capture/font_heights.py`: one workbook per face,
    Calibri, Arial, Aptos Narrow, Times New Roman, Courier New and Georgia, its Normal style the face
    at 1 point; a number in the face at every size from 1 to 72 in half steps, and the face at 9,
    11, 14 and 20 beside Calibri at 9 to 24; 1,074 rows in `font_height.layout.tsv`; then `wrapped`:
    one to three wrapped lines of every face at 9, 11, 14 and 20 under a Calibri 20 and a Calibri 9
    Normal font, 144 rows in `wrapped_height.layout.tsv`; both read through AppleScript, Excel
    16.115).** The rule tested, `F(S) = floor(a · ceil(S) + 1.5)` with a slope `a` per face, was
    written down with its predictions before the run (slopes Calibri 1.2689, Arial 1.2120, Aptos
    Narrow 1.2820, Times New Roman 1.2155, Courier New 1.3013).
    - **From 7.5 to 72 points** every face fits one slope at every whole and half size, a half size
      taking the next whole one. The windows the rows allow: Calibri 1.2685–1.2692, Aptos Narrow
      1.2812–1.2818, Arial 1.2119–1.2125, Times New Roman 1.2159–1.2164 (three rows refuted the
      predicted 1.2155), Courier New 1.3000–1.3014, Georgia 1.2188–1.2192 (no prediction). The
      package takes each window's middle. The slope is not the font's ascent plus descent.
    - **Below 7.5 points** no slope fits: 1 point gives 4, 1.5 to 3 give 7, 3.5 and 4 give 8, 4.5
      and 5 give 9, 5.5 gives 10, in every face; 6 to 7 differ by face (Calibri 9, 10, 10; Arial
      10, 11, 11). The package keeps these captured heights.
    - **Rows mixing faces:** a row is the largest ascent `ceil(S) + 1` plus the largest descent
      `B = F − ceil(S) − 1` of its fonts: all 144 rows beside Calibri predicted before the run, and
      all 35,100 rows of `row_height.layout.tsv` with it, replacing the descent per face that item
      55 fitted. `B` is item 55's bottom offset (`vertical_metrics`), and the line inset is `B − 2`
      at all 27 faces measured before; under a Calibri 20 Normal font (inset 3) wrapped text in each
      new face sits higher by `3 − (B − 2)` where that is positive: all 144 wrapped rows predicted
      before the run.

58. **Underlined rows (2026-11-19; `capture/font_heights.py underlines`: one workbook per face, its
    Normal style Calibri at 1 point, the face at every whole size from 6 to 72 under no underline,
    single, double, single accounting and double accounting, one line, and three wrapped lines of
    "H" with no underline and with single accounting; 2,814 rows in `underline_height.layout.tsv`;
    font numbers read by `capture/font_tables.py`).** The prediction written before the first run,
    that an accounting underline adds 3 points to a row and to each wrapped line, held only up to
    about 30 points.
    - **Single and double underlines change no row**, wrapped or not, in any face (1,206 rows).
    - **An accounting underline adds `E` points to the row and to each wrapped line** (402 of 402
      wrapped rows add `E` a line). For Calibri, Arial, Aptos Narrow and Georgia,
      `E = 3 · max(1, floor(S · underlineThickness / unitsPerEm))` at every size (268 of 268; the
      thickness from the font's `post` table: Calibri 134, Arial 150, Aptos Narrow 50, Georgia 101
      of 2048). Times New Roman follows it except at 37 to 40 points (4, not 3); Courier New grows
      by single points (3 at 6 to 17, then 4, 5, 6, 7, 9, 10, 11 at 72), which a term from the
      accounting line's position (`round(S · usWinDescent / unitsPerEm)`) does not reproduce
      either. Double accounting adds what single does, except in Arial at 28 to 35 and from 41
      points, where it adds 3 less, and 6 less from 69 (9 at 72, where single adds 15).
    - **A workbook may hold no more than about 1,400 rows** for this route: in a first run with all
      faces in one workbook, every row from 1,406 on read 16, the default height of a row Excel
      had not fitted. Each face now has a workbook of its own.

59. **Placement of more faces and sizes (2026-11-19; `capture/screen_vertical.py` with Times New
    Roman, Courier New and Georgia at 9 to 24 points, and Calibri and Aptos Narrow at 13, 15, 17
    and 22, appended to item 55's cases: 780 rows a Normal font, 2,340 in `vertical.layout.tsv`).**
    The 1,080 earlier rows reproduce; the four item 55 lists as read under the sheet-tab bar now
    read clean, since each view starts at the previous view's last row. Every one of the 1,260
    new rows matches the placement predicted before the run from the row model (top baseline
    `ceil(S) + 1`, bottom `B = F − ceil(S) − 1`, at least the Normal font's, line inset `B − 2`).
    The ink, per face and size (the same in every row of it): Times New Roman 6/2, 7/2, 7.5/2.5,
    8/2.5, 9.5/3, 11/3, 12/4, 13.5/4, 16/5 from 9 to 24; Courier New 5.5/1.5, 6/2, 6.5/2, 7/2.5,
    8/2.5, 9.5/3, 10.5/3.5, 11.5/4, 14/4.5; Georgia 6.5/2, 7/2, 8/2.5, 8.5/2.5, 10/3, 11.5/3.5,
    12.5/4, 14/4, 17/5; Calibri 8.5/2.5, 9.5/2.5, 11/3, 14/4 and Aptos Narrow 8.5/2.5, 10/3,
    11.5/3, 14.5/4 at 13, 15, 17 and 22. The descender is half a point either way by alignment at
    Times New Roman 11, 16 and 20 and Georgia 11 and 20, as at Arial 11 and 16.

60. **Rules and text color on screen (2026-11-19; `capture/screen_rules.py`: "HHHH" in the six
    faces at 9, 11, 14, 20 and 36 points without a rule and under each underline and a
    strikethrough, at zoom 100 and 200; "Total" and 12 in Calibri 11 under nine formats; 396 rows
    in `rules.layout.tsv`; predictions written before the run in `capture/build/`).** Positions
    are in screen points at the screen size `P = S · zoom / 100`, from the font's tables:
    - **Thickness:** an underline's is `max(1, floor(P · post.underlineThickness / unitsPerEm))`
      (119 of 119 single and accounting rules, where `round`, predicted, gave a thicker line at
      zoom 200), a strikethrough's `max(1, round(P · OS/2.yStrikeoutSize / unitsPerEm))` (58 of
      60; the two others are a one-pixel read at the top of a 36-point row at zoom 200, the strike
      itself measured as predicted below it).
    - **Position:** the single underline's top `floor(P · −underlinePosition / unitsPerEm)` below
      the baseline in 51 of 59 rows (six of the misses are Courier New's, whose underline sits
      higher than its table says); the strikethrough's top `round(P · yStrikeoutPosition /
      unitsPerEm)` above it in 57 of 60 (one real miss, Calibri 9 at zoom 200); the accounting
      rule's top `round(P · usWinDescent / unitsPerEm)` below it in 58 of 60 (Aptos Narrow 9 and 20
      at zoom 100, a point lower).
    - **Doubles:** the second line lies two thicknesses below the first in 100 of 120 rows, and
      the first line of a double underline sits higher than a single one's in 24 of 60. Settled by
      item 62; read with item 62's baselines and its cut lines, every row of this capture but
      Calibri 9's single underline at zoom 100 (joined to the stems) matches `rule_metrics`.
    - **Text color:** text under `[Blue]General` shows blue and under `[Blue]@` blue, through a
      four-section format's `[Red]@` red; under `[Red]0.0`, two or three numeric sections, a fourth
      section `@` without a color, or a condition, text stays black. The numbers' colors match the
      package's in all nine formats. The `colors` suite's read-back gave black for text under
      `[Red]General`: one of the read-back's known misreports of colored text.
    - **Display (2026-11-21):** captured again in a macOS virtual machine (item 61), twice (the
      second time with the capture recording the display's profile in the header): all 360 rule
      rows match the first capture's, and the 20 colored ones of the 36 text-color rows read the specified
      colors (Excel's blue as `53 53 255` and `36 36 255` where the first capture read `50 50 246`
      and `33 33 245`, its red `255 37 37` where it read `235 70 57`). The first capture's display
      has the Display P3 profile, and its window captures are converted to sRGB, which pulls
      saturated colors a few units off. The VM's display profile ("Apple Virtual") has sRGB's primaries
      (its red, green and blue colorants equal the system sRGB profile's to four decimals).
      `rules.layout.tsv` is the VM's capture since, its header recording `display_profile`;
      every color reads as the same named color as before.

61. **Times New Roman's accounting rows at 36 to 41 points (2026-11-21; `capture/font_heights.py
    accounting`, Times New Roman and Georgia at every half point from 36 to 41, plain and under
    single and double accounting, one line and three wrapped, 132 rows in
    `accounting_height.layout.tsv`; the first capture of this repository run in a macOS virtual
    machine with Excel 16.115 (26110837), after the `wrapped` part of item 57 reran there with
    all 144 data lines identical to its committed capture).** All 132 rows were predicted before
    the run: Times New Roman's accounting underline adds 4 points from 36.5 to 40, single and
    double alike and per wrapped line, where item 58's rule gives 3, and 6 from 40.5 as the rule
    says; Georgia, with nearly the same thickness (101 of 2048 against 100), follows the rule. A
    half size takes the next whole one. No rule explains the 4: in both faces the accounting line's
    top (`round(S · usWinDescent / unitsPerEm)`) equals the face's descent `B` at every size from 7
    to 72, so its position adds nothing there, and the thickness does not tell the two faces apart.
    The package keeps the four sizes as captured.

62. **Where a double underline's lines go (2026-11-21; `capture/screen_rules.py rules_double`:
    "HHHH" in Calibri, Arial, Aptos Narrow, Times New Roman and Georgia at every size from 9 to
    36 points and from 40 to 72 in steps of 4, under each underline and a strikethrough, at zoom
    100 and 200, then 20 small rows; 1,850 rows in `rules_double.layout.tsv`, captured in the
    macOS virtual machine, display profile with sRGB's primaries).** Predictions were written
    before the run (`capture/build/`), with a third of the 370 face, size and zoom combinations
    held out (seed 20261121): the placement as it stood (a double's first line at the single
    underline's top) and the stepped double below, from the font's tables. Read with each size's
    baselines from rows no rule touches: the strikethrough row for the plain rows, the single
    accounting row for the accounting rows. The first run, with column B 30 characters wide, read
    no accounting underline where "HHHH" overflowed the cell (24 rows from 56 points): the
    underline spans the cell, so it was shorter than the text and not taken as a rule. The run
    used here has a 60-character column and differs from the first only in those rows and in 8
    strikethrough rows that gained the line described below.
    - **Screen size alone:** wherever the same screen size `P = size · zoom / 100` occurs at both
      zooms (95 face and screen sizes), the lines are equal.
    - **Single and accounting:** a single accounting underline has the single's thickness; a
      double accounting underline's first line is at the single accounting line's top, and its
      lines have the double underline's thickness and gap. The font tables give the single
      underline in 357 of 369 rows, the single accounting underline in 335 of 370 (Aptos Narrow's
      a point lower at 20 screen sizes, Calibri's from 80), the strikethrough in 361 of 370. From
      80 screen points Calibri's lines fall on half screen points.
    - **The double's first line** starts at the single underline's top and steps up by its line
      thickness until the double (three thicknesses) ends at or above the accounting underline's
      top, as Calibri 11 at zoom 100 shows (single at 1 point, double at 0 and 2, accounting at
      3); the second line lies two thicknesses below (366 of 369 rows; three thicknesses in
      Times New Roman at screen sizes 38 and 40). The first line can rise above the baseline:
      Arial 9 at zoom 100 a point above, as Calibri 9 at zoom 100 in item 60.
    - **The double's thickness** is `max(1, floor(P · min(underlineThickness, yStrikeoutSize) /
      unitsPerEm))`, 232 of 246 selection rows, against 215 for the single's thickness. Arial
      alone tells them apart: its strikeout size is 102 units, its underline 150, and its doubles
      are thinner than its single underline.
    - **Scores, the double's two lines exact:** the placement as it stood, 238 of 369 readable
      rows; the stepped double from the tables, as predicted, 282; the same with the captured
      single, thickness and accounting lines, 298. The rule adopted (the stepped double with the
      double's own thickness and no stop at the baseline, chosen on the selection rows) scores
      198 of 246 selection rows and, scored once, 98 of 123 held out (88 for the prediction).
      Exact in Aptos Narrow but for 9 points and in Georgia but for 128 and 144 screen points.
    - **Not explained:** in Arial at most screen sizes from 29, in Times New Roman at about half
      of those from 38 and in Calibri from 80, the first line is drawn one to three screen points
      higher than the rule gives, and in Arial at some sizes the lines are a screen point thicker
      than `floor` gives. A stricter fit test, a fit in the font's own units, a fit against the
      row's bottom and a step of the single's thickness were tried; none explains them. The
      package returns the drawn lines at every captured screen size (`capture/rule_table.py`
      writes `src/xlformat/_rules_drawn.py`) and the rule elsewhere; `QUIRKS.md` lists the misses.
    - **A double accounting underline outgrows its row at zoom 200:** at zoom 100 every double
      accounting underline ends 2 pixels above its row's bottom (the row is fitted to it); at
      zoom 200 the second line reaches that point and runs on into the row below in 55 of 185
      rows, the next row showing its end at its top. The rows' heights are sheet points, fitted
      at the sheet's size, and the lines are placed at the screen size.

63. **Courier New's rules and screen sizes that are not whole (2026-11-21; `capture/screen_rules.py
    rules_courier`: Courier New at every size from 6 to 72 points and every half size from 36.5
    to 71.5, Times New Roman from 36 to 41 in half steps, Calibri at 9, 14, 20, 36 and 48 as
    controls, each under every underline and a strikethrough, zoom 100 and 200, 1,190 rows;
    `rules_zoom`: Courier New and Calibri at 20 sizes from 8 to 70.5, zoom 100 and 150, 400
    rows; both in a macOS virtual machine).** Predictions and their scorer were committed before
    the runs (`capture/courier_predict.py`, `capture/predictions/`): Courier New's double with
    its second line at the single's table position `u = floor(P · 477 / 2048)` and its first a
    gap above it, the gap a table by screen size (2 up to 17, then 3 from 18, ... 9 at 65 to
    72), one point lower at `P` = 12; the single at `u` while the gap is two thicknesses, on the
    double's first line above that; accounting at `round(P · 615 / 2048)`, a double accounting
    underline's second line the gap below. For a fractional `P`, four candidates: the nearest
    whole size (a half up), `floor`, `ceil`, and the formulas fed the fraction.
    - **Courier New at whole screen sizes:** 488 of 490 rows as predicted; the two others,
      Courier New 6 at zoom 100, show no line to the reader. Zoom 200 at a size equals zoom 100
      at twice that size in 155 of 155 pairs. At the odd screen sizes from 73 to 143, measured
      here, the gap rises from 10 to 19, one value at each size. The double is drawn a point
      lower at `P` = 12 at both zooms, the single not. The predictions' gap table came, up to
      72, from Courier New's accounting rows (item 58), and above that from an earlier screen
      reading at even sizes; every one of the 66 accounting rows readable on screen grows by
      the drawn gap plus the line's thickness, single and double accounting alike.
    - **Times New Roman, not predicted:** from 36.5 to 40 points at zoom 100 (screen sizes 37 to
      40; 36.5 rounds to 37) its double behaves as Courier New's: second line at the single's
      table position, first line three thicknesses above it, the single on the double's first
      line, a double accounting underline's lines three thicknesses apart. That is the 4 points
      its accounting rows add there (item 61) and the three-thickness gaps of item 62.
    - **One law for accounting rows:** in every face, an accounting row grows by its line's
      thickness plus the gap of a double drawn in that line (`rules_double` and
      `underline_height`): three thicknesses where the gap is two, Arial's double accounting
      with its thinner double lines (the 3, 6 and 9 points of item 58), Times New Roman's 4 at
      37 to 40, Courier New's gap and thickness. Every captured row height is unchanged.
    - **Screen sizes that are not whole:** at zoom 100's half sizes (Courier New 36.5 to 71.5)
      the nearest whole size and `ceil` fit 180 of 180, the fraction 147, `floor` 129. In
      `rules_zoom` the nearest whole size fits best: Courier New 170 of 180 fractional rows,
      Calibri 156 (`ceil` 162 and 143, the fraction 145 and 141, `floor` 127 and 126). With the
      drawn lines at that whole size in place of the tables, 23 of 400 rows differ by a screen
      point (`QUIRKS.md`). At screen size 34.5 Courier New draws a point lower at zoom 150 than
      at zoom 100; Calibri draws alike.
    - **The reader at other zooms:** at zoom 50 (one pixel a sheet point) it stopped at its first
      view of large rows, so `rules_zoom` runs at 150. There Excel rounds each row to whole
      pixels, so the reader takes the known 3 pixels a point and each row's own band, and the
      zoom-150 rows were read again from the run's screenshots after that change; the zoom-100
      rows read identically both times. The scorer's cut of a line at the row's bottom, from the
      row's height, was corrected after the run for the same rounding: equal lines had counted as
      different.
64. **A cell font whose padding passes the Normal font's (2026-11-22; `capture/cases.py
    cell_font_large`: the cell-font values in Calibri 16, 18, 20, 24, Aptos Narrow 18, 20, 24
    and Arial 14, 16, 18, 20, 24 cells of a Calibri 11 workbook, at every point from 11 to 100,
    26,026 rows, in a macOS virtual machine).** A column's padding is `2·ceil(MDW/4)+1`, 5
    points up to a digit of 8, so no earlier suite had a cell font whose padding differed from
    the Normal font's. Predictions and their scorer were committed before the run
    (`capture/padding_predict.py`, `capture/predictions/cell_font_large.tsv.gz`) with two
    candidates: the cell's budget the column's (the model then), and the budget less
    `max(0, P(cell) − P(normal))` points, General's units less that over the cell font's digit.
    They differ on 3,906 of the 25,740 rows with a width; the Normal font's rows and the digit-8
    controls (Calibri 16, Arial 14) are not among them. The second candidate fits 25,740 of
    25,740 rows; the first misses all 3,906 where they differ (a cell one character short:
    `###` where Excel shows `##`). Calibri and Aptos Narrow from 18 points and Arial from 16
    lose 2 points, Arial 24 (digit 13, padding 9) 4.
65. **Excel for Windows by capture jobs (2026-11-22; `capture/job.py`, `tests/captures/windows/`).**
    A group of suites runs as one job in sheetmill's capture-job format (version 1), one
    workbook per suite, each built as for the Mac with fixed zip times and document dates. The
    steps are the Mac capture script's: `copy_text` for a width-less column, `set_size` and a
    `read` of each cell's text for a column at a width (with its width read back in units and
    points), `autofit` and a width read for a measure sheet, and the display format's font color
    per cell for a suite that reads colors. Each suite file in `tests/captures/windows/` has the
    Mac schema and holds the job's result header whole (`job_header`); the manifest is beside
    them as `<group>.job.json`, and `tests/test_windows_captures.py` writes each job again from
    its `job_id` and requires the same manifest and input checksums.
    - **Glyph tables, 96 dpi** (groups `tables_calibri`, `tables_aptos_narrow`, `tables_arial`;
      Excel 16.0.20527.20196, Microsoft 365, Current Channel): every `measure*` and `space_fill*`
      suite of the three families, four styles, nine sizes, 216 suite files.
66. **Windows rules at 96 dpi (2026-11-22; group `rules`, the build of item 65).** 50 suites,
    119,776 rows: columns without a width under base widths 6 to 12 (`column_default_*`), file
    widths on either side of a 1/256 step (`column_boundary*`), fills past 255 characters
    (`long_fills`), file widths at every size (`column_file*`), cell fonts against a Calibri 11
    Normal font (`cell_font_styles`, `cell_font_large`, `cell_font_arial`) and error values
    (`error_width`). The package's own Windows rendering of every row other than the error values'
    was committed before the run. The package missed 358 of those 116,768 width rows, all at a half
    width (2.5, 4.5, ...), each one character or `#` short: Windows rounds a column's half pixel up,
    where the package rounded it to even (36.5 pixels is 37 at a digit of 7 and width 4.5). With
    that rounding every row of the 50 suites matches, among them the 28 `General` rows under a
    larger cell font. The Mac rounds a half point up too (item 67). The group confirms on Windows:
    the default column rounded up to a multiple of 8 pixels at every base width and size; a fill
    past 255 characters showing 255; every 1/256 boundary width truncated, as on the Mac; and a cell
    font losing the padding it has beyond the Normal font's (item 64).
67. **A column of a half point on the Mac (2026-11-22; `capture/cases.py column_half_points_*`:
    the cell-font values in eight Normal fonts, Calibri 9, 11, 14, 18 and 20 and Arial 16, 20
    and 24, at every width from 1 to 20 whose column `MDW · w + P` is a half point, 4,180 rows
    and 190 columns, in a macOS virtual machine).** No earlier Mac suite set a width giving a
    half point. Predictions and their scorer were committed before the run
    (`capture/half_predict.py`, `capture/predictions/column_half_points.tsv`) with two
    candidates, the package's rounding to even and rounding up, as Windows rounds a half pixel
    (item 66). They differ on 96 columns' points and 143 rows' texts. Rounding up fits all 96
    columns and all 143 rows, to even none of them; the 94 columns and 4,037 rows where the two
    agree fit both. One rule for both platforms: a half rounds up (12.5 points is 13 at Calibri
    9 and width 1.5).
68. **Windows validation at 96 dpi (2026-11-22; groups `validation_numbers` and
    `validation_general`, the build of item 65).** Rows the Windows tables were not fitted on: 32
    suites and 337,117 rows, among them `numeric_core_width` in Calibri 11, Aptos Narrow 11 and
    Arial 10, `cell_font_numeric`, `General` at every point from 9 to 24 and among literals,
    `General` in cell fonts, narrow columns, fills, letters and elapsed decimals. The package's
    Windows rendering of every row was committed before the runs. Every one of the 331,786 rows at
    a width is as predicted. Of the 5,331 width-less rows, read through the clipboard, 189 differ,
    all in `letters`, `letters_upper` and `elapsed_decimals`: the Mac's two known deviations of the
    clipboard (a Buddhist-year code's plain year, an elapsed code after a date code), the same on
    Windows. In character mode 23 rows of Arial 10 and `width_fill` show a fill one character
    shorter than the width, because the text's glyphs are wider than a digit; they are
    glyph-slack rows (`tests/test_captures.py`). The merged-cell suites, whose rows hold the
    block's width in the Mac's points, are captured on Windows by their own group (item 71).
69. **Windows row heights at 96 dpi (2026-11-22; group `rows`, the build of item 65).** The
    workbooks of the Mac's row capture (item 44), one per Normal font (Calibri, Aptos Narrow and
    Arial, four styles, nine sizes), each row's height read by the job: 34,992 rows in
    `tests/captures/windows/row_height.layout.tsv.gz`, without the Mac's standard-height lines (the
    job format reads no sheet property; the row without an element gives the same height). In
    pixels, a face has a height H and a descent D, and a row is `max(lines · H − D) + max D` over
    its cells, the Normal font a one-line cell among them: `capture/windows_rows.py` takes H from
    each face's own row and searches D, and the table fits all 34,992 rows, each D the only value
    that does. Calibri and Aptos Narrow are the same in every style; Arial's bold and italic differ
    from its regular at 11, 12, 14, 20 and 24 points (bold Arial 11 is 20 pixels, regular 19).
70. **Windows built-in formats (2026-11-22; group `builtin`, the build of item 65).** The Mac's
    probe of item 49 as a job: ids 0 to 163 in four workbooks, each id's format read as stored and
    as shown locally and seven values' texts at a column width of 40, in the Windows
    `tests/captures/windows/builtin_formats.layout.tsv`. No group was refused. Every format string
    is the Mac's except the 20 ids whose Mac string is the short date: 14, 27 to 31, 36, 50 to 58,
    71 and 72 are `m/d/yyyy`, and 22 and 78 `m/d/yyyy h:mm`, the en-US Windows short date; the
    stored and local strings agree for every id. The package's Windows rendering of each format
    Windows reports gives all 1,148 texts.
71. **Windows merged cells (2026-11-22; group `merged`, the build of item 65).** `merged11` and
    `merged14` as a job, 5,940 rows: each block's anchor column read as any other, and every
    column of the block read in row 1, its width the sum of their points as Excel for Windows
    measured them (a hidden column reads 0). The package's Windows rendering, committed before
    the run with each block's points predicted from its columns' file widths, holds for every
    row, block widths included: a merged block formats as one column of its columns' pixels
    with one padding, as on the Mac (item 40).
72. **Windows TEXT (2026-11-22; group `text_function`, the build of item 65).** The
    `text_function` suite as a job: Excel's `TEXT` over every format of the other suites, 29,580
    cells computed when the workbook opens, read through the clipboard. The prediction committed
    before the run was the package's `text_function`, which follows the Mac: every row is as
    predicted, so `TEXT` gives the same text on both platforms in en-US.
73. **Window pictures on Excel for Windows: the pilot (2026-11-23; group `screen_pilot`,
    `capture/screen_windows.py`, the build of item 65, ClearType font smoothing as the header
    records it).** Calibri and Arial at 9, 11 and 20 points under no rule, each underline and a
    strikethrough at zoom 75, 100 and 200, and Calibri and Arial 11 placed top, center and bottom
    in rows of the face's height H, H + 13 and H + 14 pixels and 2H, fitted and wrapped, at zoom
    100: 9 pictures, in `screen_pilot_rules.layout.tsv` and `screen_pilot_vertical.layout.tsv`.
    - **The frame.** The column headings' border covers the top pixel of a view's first row, so
      that row is never whole; row 1 is every first view's first row and is not measured (the
      groups after the pilot put a filler there). Every other band of every view begins where
      the rows above it end, each row's height drawn as whole pixels rounded half up and the
      rows stacked: at zoom 100 and 200 the pixels are exact, at 75 a row of 16.5 pixels is 17.
      The reader anchors each view on those bands and stops when one disagrees.
    - **A light blue ribbon icon** (155, 222, 255) passes the reader's cyan marker test, so a
      marker column is now one whose fill runs down 8 pixels or more.
    - **Placement, in the 27 rows measured (baseline: the H stems' last ink row, from the row's
      first pixel row).** Top-aligned text has its baseline at H − D, centered text at
      ⌊(R − H)/2⌋ + H − D in a row of R pixels, and bottom-aligned text at R − max(D, D'), D'
      the Normal font's descent: Arial 11 (D 4) in a Calibri 11 workbook (D 5) sits a pixel
      higher than its own D would put it, as the Mac's rule of item 55 does in points. These
      are the pilot's rows only; `screen_vertical` decides the rules.
    - **Rules, findings only.** At zoom 75 and 100 every line, the strikethrough's too, lies a
      pixel lower against the measured baseline than the Mac's lines of item 60 read in device
      pixels at the ppem rounded to a whole pixel; at zoom 200 the lines are thicker than the
      Mac's floor gives (2 or 3 pixels). Neither is applied to `rule_metrics` until
      `screen_rules` measures it across the faces and sizes. Two rows read as the Mac's do at
      their edges: a double accounting underline whose second line would pass the row's last
      pixel shows one line (zoom 75), and a line on a row's first pixel row is the row above's.
74. **Text placement on Excel for Windows (2026-11-23; group `screen_vertical`, the build of item
    65).** "Hxg" in Calibri, Aptos Narrow and Arial at the nine table sizes, fitted, wrapped to two
    and three lines, and placed top, center and bottom in rows of H, H + 13 and H + 14 pixels and
    2H, with bold, italic and bold italic at 11, 12 and 20 points, in a Calibri 11 workbook; and
    the three faces at 9, 11 and 20 points in workbooks of Calibri 9 and 20, Arial 9 and 11 and
    Aptos Narrow 11: 693 rows from 33 pictures at zoom 100, in `screen_vertical.layout.tsv`.
    Every band after a view's first lies at the view's one offset from the rows (760 of 760). In
    pixels from the row's first pixel row, H and D the cell face's height and descent (item 69)
    and D' the Normal font's descent, the baseline of every row is:
    - **top:** H − D; **center:** ⌊(R − H)/2⌋ + H − D in a row of R pixels; **bottom:**
      R − max(D, D');
    - **wrapped (bottom-aligned):** lines H apart, the last one D' − D higher still where the
      face's descent is smaller than the Normal font's (Calibri 9 and 10, Aptos Narrow 9, Arial 9
      to 12 in the Calibri 11 workbook: a pixel), as the Mac's wrapped lines sit by the difference
      of their insets (item 55).
    These are the Mac's placement rules in device pixels. The reader now also drops a marker run
    under 4 pixels: with the Arial 9 Normal font, column A's marker column lay under the same
    light blue ribbon icon as in item 73, above the cells.

75. **Rules on Excel for Windows (2026-11-23; groups `screen_rules` and `screen_rules_zoom`, the
    build of item 65).** "HHHH" in Calibri, Aptos Narrow and Arial at every size from 9 to 36 and
    from 40 to 72 in steps of 4 under each underline and a strikethrough at zoom 100 and 200
    (1,110 rows from 144 pictures), and Calibri and Arial at the nine table sizes at zoom 75, 90,
    100, 110, 125 and 150 (540 rows from 27): `screen_rules.layout.tsv`,
    `screen_rules_zoom.layout.tsv`. Every view's bands lie at its one offset. Measured from the
    baseline (the H stems' last ink row), the lines follow no candidate closely: the Mac's rule
    formulas of item 60 read in device pixels at the ppem rounded to a whole pixel give 62 of
    1,650 rows; every line a pixel lower, 958; that with the thicknesses rounded, 1,133 (at zoom
    200 a single underline is often two pixels where the floor gives one). No rule is taken into
    `rule_metrics` for Windows from these rows; they are what a Windows rule or table is fitted
    to. A double accounting underline's second line that ends on the row's last pixel row is cut
    there (item 78 draws it whole in taller rows): its thickness in these rows is no evidence.
76. **Wrapped text under a larger Normal font on Excel for Windows (2026-11-23; group
    `screen_wrapped`, the build of item 65).** Bottom-aligned cells wrapped to two and three lines
    in Calibri 11, Arial 11, Aptos Narrow 14 and Arial 18 in a Calibri 20 workbook and Arial 9
    and Calibri 9 in an Arial 14 one, faces whose descent D is smaller than the Normal font's
    D', with Calibri 24, Calibri 14 and Arial 16 beside them as controls: 18 rows from 2
    pictures, `screen_wrapped.layout.tsv`. Written as a prediction from item 74's wrapped rows
    before the run, the last line's baseline at R − max(D, D') − max(0, D' − D), lines H apart,
    holds for all 18; without the second term, for the 6 controls only. Wrapped lines sit the
    difference of the descents higher, as the Mac's do by the difference of their insets.

77. **Rules on Excel for Windows at new sizes and zooms (2026-11-23; group `screen_rules_fresh`,
    the build of item 65).** Calibri, Aptos Narrow and Arial at 37, 38, 39, 41, 43, 45, 49, 53
    and 57 points under each underline and a strikethrough at zoom 80, 120, 140 and 175, none of
    which items 73 and 75 hold: 540 rows from 87 pictures, `screen_rules_fresh.layout.tsv`.
    Rules written from the rows of item 75 and frozen before this run hold for every row, with P
    the ppem rounded to a whole pixel (size · 4/3 · zoom / 100), em the face's units per em and
    every rounding half up, positions in pixels below the baseline (the H stems' last ink row):
    - thickness of a single underline, a single accounting underline and a strikethrough:
      max(1, round(P · thickness / em)), from `post.underlineThickness` and
      `OS/2.yStrikeoutSize`;
    - single underline: round(P · −`post.underlinePosition` / em) + 1;
    - strikethrough: round(P · `OS/2.yStrikeoutPosition` / em) − 1 above the baseline;
    - accounting underline (single, and a double's first line): the hinted descent + 1, the
      descent being Arial's `VDMX` yMin at P and round(P · `OS/2.usWinDescent` / em) for
      Calibri and Aptos Narrow, which have no `VDMX`;
    - a double underline's two lines (both kinds): twice a line's thickness apart, a double
      accounting line as thick as a double's;
    - Aptos Narrow's double underline: the single's top and thickness, stepped up a thickness at
      a time until it ends at or above the accounting underline's top, the Mac's rule of item 62.
    A double accounting underline's second line past the row's last pixel is not measured (10
    rows), and one ending on the row's last pixel row is cut there (item 78): its thickness is
    no evidence. The font numbers are the Mac's copies of the faces (Calibri 6.20, Arial 6.80i,
    Aptos Narrow 2.01), where Windows draws Calibri 6.27 and Arial 7.07; they agree at the ppems
    measured here (39 to 133). Calibri's and Arial's double underline (its top and thickness)
    follows none of these rules.

78. **Rules by ppem on Excel for Windows, a row holding every rule (2026-11-23; group
    `screen_rules_table`, `capture/screen_table.py`, the build of item 65).** "HH" in Calibri and
    Arial, a row a face and size, its cells under a single underline, a double, a strikethrough,
    a single accounting and a double accounting underline side by side (three to five to a row,
    as fit the window at the zoom), one workbook per zoom at 16 zooms from 15 to 400: 4,370
    cells from 372 pictures, `screen_rules_table.layout.tsv`. Each cell gives its own baseline
    (the H stems' last ink row) and the row's, from its single underline's cell, since a double's
    first line can lie on the stems' last row (at ppem 9 and 10 in items 73 and 75).
    - Zoom 15: column A is about 5 pixels wide and the reader finds no marker in it; its 10
      cells are written unmeasured.
    - In the workbooks of zoom 150 and 200, which define 1,431 fonts each, every cell from F201
      on (Arial 29's double accounting, then Arial from 29.5 to 72 points) is drawn in the
      default font, without its underline: F201 is the first cell whose font is number 1,000
      in the file's list (counting from 0), and the cells before it in its row are drawn in
      their own fonts. Rows from 202 on keep 15 points, fitted to that font: 172 rows, a gap
      in this capture. The other 14 workbooks define at most 301 fonts and are drawn as built.
      Observed on Excel for Windows: in that workbook, cells whose font is number 1,000 or later
      in the font list were drawn in the default font.
    - Against items 73 and 75 at the sizes and zooms both hold (zoom 75, 100, 150 and 200,
      measured from the baseline, a double from its single's): every top, thickness and gap
      of the single, strikethrough, single accounting and double underlines agrees (90 of 90
      rows), and the double accounting's tops and first lines too. The double accounting's
      second line, which ended on or past the row's last pixel row there, is drawn whole here,
      in rows about 2 points taller.

79. **The cells of item 78 drawn in the default font, again (2026-11-23; group
    `screen_rules_table_arial`, the build of item 65).** Arial from 29 to 72 points under the
    five rules at zoom 150 and 200, in four workbooks of at most 220 fonts: 870 cells from 57
    pictures, `screen_rules_table_arial.layout.tsv`, every row fitted. They agree with items 75
    and 78 at every size, zoom and rule both hold, but for the double accounting's second line
    cut there by the row's last pixel row. Over items 78 and 79 together (the cells of this item
    in place of those item 78 drew plainly), with P the ppem:
    - item 77's rules hold for every thickness (every measured cell) and every double's gap of
      two thicknesses; the single underline's, the strikethrough's and the accounting
      underlines' tops hold in Arial at P 11 to 384 and in Calibri at P 11 to 170. In Calibri
      from P 171 (171 to 192, and every eighth P to 384), every one of those lines is a pixel
      higher, measured from the H stems' last ink row; below P 11 both faces miss by a pixel
      here and there (P 9 and 10 as in item 75's rows);
    - an equal P draws equal lines: 5,976 face, rule and part keys drawn at two zooms or more
      agree but for 58, each the thickness of Calibri's double accounting second line where the
      row's last pixel row cuts it (72 of the 73 cells where that line is thinner than the first
      end on the row's last pixel row).
    At zoom 15 nothing is measured (item 78), and 14 double gaps, all at P 14 or less, are not.

80. **Rules at the large ppems items 78 and 79 left out, and Aptos Narrow (2026-11-23; group
    `screen_rules_table_high`, the build of item 65).** Calibri and Arial at every ppem P from
    193 to 383 not a multiple of 8, once each (both doubles; Calibri's single, strikethrough and
    single accounting underlines too), and Aptos Narrow under every rule at every P to 11 (zoom
    25 and 75) and every eighth P from 200 to 384 (two zooms each, 384 at zoom 400): 1,521 cells
    from 555 pictures in 23 workbooks of at most 161 fonts, `screen_rules_table_high.layout.tsv`.
    At zoom 25 Excel draws no vertical gridlines; the reader places those cells from the columns'
    points. Written from items 78 and 79 before this run, these hold for every measured cell:
    - Calibri from P 171: item 77's tops of the single, strikethrough and accounting underlines
      a pixel higher, their thicknesses as item 77's (168 cells of each);
    - Aptos Narrow from P 200: item 77's rules for every line and the Mac's stepping rule for its
      double (47 cells of each part);
    - Arial: the double accounting's top as item 77's accounting top, but at P 255, a pixel
      higher (167 of 168); and in both faces a double's gap of two thicknesses.
    Aptos Narrow below P 12 misses item 77's tops here and there, at zoom 25 and 75 alike, as
    the other faces do (item 79); its accounting lines at zoom 25, in rows of about 5 pixels, are
    mostly past the row. With items 78 and 79, Calibri's and Arial's double underlines (both
    kinds) are now measured at every P from 1 to 384, and 6,416 face, rule and part keys drawn
    at two zooms or more agree but for 67, each a line cut by its row's last pixel row (58 of
    item 79's, 9 of Aptos Narrow's double accounting at zoom 25). Such a line is not cut where
    Excel draws it: it goes on into the row below, under that row's top gridline. Of the 53
    double accounting second lines of item 78 that end on their row's last pixel row thinner
    than their first, 35 lack one pixel row, the row below's gridline row; of the 18 that lack
    two or three, 13 show the rest on the row below's next pixel rows (5 do not, unexplained).
    At zoom 25, which draws no gridlines, Aptos Narrow's lines show on the next row's first
    pixel rows.

81. **Text insets on screen, both platforms (2026-11-23; `capture/screen_insets.py`; Windows
    group `screen_insets`, the build of item 65; the Mac in a macOS virtual machine, 16.115).** "HH" in
    Calibri, Arial and Aptos Narrow at every size from 9 to 36 points and every fourth from 40
    to 72, left- and right-aligned under a single underline (which runs from the text's origin
    to its end) and under a single accounting underline (which spans the cell), in workbooks of
    two Normal fonts (Calibri 11, Aptos Narrow 20), at zoom 100 and 150: 444 rows each,
    `windows/screen_insets.layout.tsv` and `insets.layout.tsv`. The clear space between each
    gridline and the text is `ceil(D / 4)` screen units (pixels on Windows, screen points on the
    Mac), `D` the width of the cell font's digit at its screen size (Windows' ppem, the Mac's
    rounded screen size), `round(screen size * digit advance / em)`; on Windows Calibri and Aptos
    Narrow hint the digit at ppem 72 to 37 pixels, a pixel wider. Every row of both platforms
    holds it, the right side as the left, a seeded third of the face, size and Normal font keys
    held out from the fit and scored after it (148 of 148 on each platform). The Normal font
    moves nothing. An accounting underline's cell places its text as a plain one (the Mac's
    gridline is two device pixels wide, one screen point; counted from its first pixel column,
    the accounting cell's text and line, in column D, start where the plain one's do, in column B
    from the end of column A's fill). Counted from the cell's left boundary, its gridline
    included, the left inset is a screen unit more than the clear space. In the
    Mac's first face row of each workbook (Calibri 9 and 40) the right accounting cell's ink
    reads up to 25 pixels past its line, unexplained; its line agrees. The same rule for the
    Normal font is the column padding's, two insets and the gridline; for Calibri 11 on the Mac,
    3 points from the left boundary and 2 before the right, as Excel's print layout places it
    (`capture/alignment.py`). The table pictures of item 78 read one pixel wide on column B's left (the
    gridline there is under column A's fill); this capture's reader takes that edge from the
    fill.

82. **Glyph advances per face, from autofit widths (2026-11-23 and 24; `capture/faces.py`,
    `capture/face_fit.py`; Excel for Mac 16.115 in a macOS virtual machine; files in
    `tests/captures/faces/`).** Three groups of autofit captures (the measure method, one
    workbook per suite, each style a cell font over the Normal style), for 27 families: Calibri,
    Arial, Aptos Narrow, Aptos, Arial Black, Arial Narrow, Book Antiqua, Bookman Old Style,
    Calibri Light, Cambria, Candara, Century Gothic, Comic Sans MS, Consolas, Courier, Courier
    New, Garamond, Georgia, Haettenschweiler, Helvetica, Impact, Lucida Console, Rockwell,
    Tahoma, Times New Roman, Trebuchet MS and Verdana, in regular, bold, italic and bold italic.
    - **Check** (`faces_c10_<family>`, `faces_c13_<family>`; 58 workbooks, 59,624 columns, 35
      minutes): each of 218 characters (space, printable ASCII, the cp1252 code page's others)
      ten times, a digit five-run, twelve texts mixing ASCII and other characters
      (`1.234,56 €`, `naïve résumé`), and long texts (14 to 31 digits, a sentence cut at 20 to
      55 characters). MS Sans Serif, which the Mac lacks, autofits as Times New Roman at every
      text of both sizes (1,028 of 1,028 each). Segoe UI, also absent, autofits as none of the
      families (the closest, Tahoma, matches 104 of its 257 ten-runs at 10 points).
    - **Recovery** (`faces_r383_<family>`, `faces_r397_<family>`, `faces_r409_<family>`; 84
      workbooks, 79,968 columns, 46 minutes): every ASCII character `g` as `é + g·N + é`, every
      other as `g·N`, `é` alone `M` times, every digit's five- and ten-run; N and M the longest
      runs under 13 digit widths of the Normal font, from the 13-point check. A text holding a
      non-ASCII character is summed unrounded, so a width W gives `S · (2t + N·g)` within half a
      point of `W - P` (P the padding). For each face the paddings the digit runs allow, the
      padding rule and the three sizes leave one solution: every character pins to one multiple
      of `1/E` (E 2048; Consolas and Lucida Console 1024; Haettenschweiler 1000) but three, off
      every grid, which the face lacks (Haettenschweiler's `·` and `Ž`, Lucida Console's `€`).
      The space of an italic monospace face (Consolas, Courier, Courier New, Lucida Console)
      comes out half a unit past the grid there, while the check's spaced texts take the
      advance the face's other characters share (Consolas 10 italic: the sentence texts, 4 to
      10 points narrower with the space off the grid); it is given that advance. Segoe UI has no
      solution: its digit runs give 0.5379 em at 409 points, 0.5390 at 397 and 0.5379 at 383,
      not one ratio, so its widths are not proportional to the size.
    - **Padding** (`faces_p6` to `faces_p30`, `faces_p32` to `faces_p192`, `faces_phalf`,
      `faces_p1to5`; 38 workbooks, 10,080 columns, 6 minutes; Normal Calibri at the size): `0`
      five and ten times in every face. The five-run less five whole digits is the padding,
      `2·ceil(d/4) + 3` for an upright face (d its widest digit) at every size from 6 to 192,
      more by twice the overhang for an italic one; the ten-run is ten digits, the padding and
      the bonus below, which a wide digit earns against Calibri's (Verdana bold 12: 90 points
      are 15 of Calibri's 6-point digits, one point). A half size autofits exactly as the whole
      size it rounds to, half up (Calibri 8.5 as 9, Garamond 9.5 as 10). Under a Calibri 6
      Normal font every width of sizes 1 to 5 is a multiple of 5 points, so that workbook gives
      no padding.
    - **Switch** (`faces_s193` to `faces_s320`; 128 workbooks, 5,279 columns, 5 minutes; Normal
      Calibri at the size): in Calibri, Arial, Times New Roman, Verdana, Candara and Georgia,
      `0` and the six ASCII characters whose per-character rounding and whole-run rounding
      disagree most, five times each. Calibri rounds each character through 285 points and
      sums the run unrounded from 286 (228 of 228 deciding rows); the other five round each
      character at every size to 320. No row fits neither. Where the rounding stops is a
      property of the face, not one size. Sizes above 192 points raise `LookupError`: this
      group measures autofit only, and no capture shows fit, `#` or fill decisions above 192
      points, where the rounding of each character could stop as autofit's does (Calibri from
      286).

    The rule, on every check row but the runs of spaces, no-break spaces and soft hyphens
    (56,896 of 56,896): an ASCII character's advance is `u · S / E` rounded half up to whole
    points, and an ASCII text the sum; a text holding any other character sums every advance
    unrounded, one character at a time in double precision, and rounds the total half up
(ten Haettenschweiler 10 `‰`, 5.05 points each, add up to 50.49999999999999, and Excel's
width is 50 plus the padding, not 51). The autofit width adds the padding, twice the
    italic overhang and a bonus of `max(0, floor(W / D / 5) - 2)` points, W the text's width
    plus one overhang and D the Normal font's digit: a point at 15 digit widths and one more
    every 5. The bonus is item 41's tenth (ten `W` at Calibri 11 measure 101). Every glyph entry
    and overhang of the 108 per-size tables of items 24, 26, 30 and 48 follows from the
    recovered advances. Fit, `#` and fill decisions round every character, non-ASCII ones too:
    the euro rows of `numeric_core_width_arial10` and `cell_font_numeric` (`€ 99.00 ` in Arial
    10 shows `######` at width 6) fit only so.

## 3. The pipeline

```
capture/cases.py              suites: tagged Case(value, format) lists plus the widths to capture at,
                             a measure flag, the date system and the fast-tier cap
capture/build.py              openpyxl -> <Excel container>/xlformat/<suite>.xlsx and
                             capture/build/<suite>.manifest.json; sheet "cases": A = row index,
                             one value column per width; sheet "measure" for autofit
capture/elapsed_forms.py      where a point and zeros are fractions of a second (item 42)
capture/letters.py            which bare letters Excel refuses or reads as codes (item 35)
capture/capture_template.applescript   filled per suite into capture/build/<suite>.capture.applescript
capture/capture.py            runs osascript, parses the output, assigns tiers, writes the TSV
capture/tiers.py, retier.py   fast-tier rule and the command that reapplies it without recapture
capture/schema.py             the TSV schema (frozen, see below), reader and writer
capture/convert_2023.py       the 2023 clipboard capture into the schema (capture2023.tsv.gz)
capture/rounding_models.py    model comparison for the 15-digit rule
capture/glyph_fit.py          glyph tables from the measure and space_fill suites (items 26, 30)
capture/alignment.py          default alignment through Excel's PDF export (item 31)
capture/row_heights.py        row heights per Normal font, cell font and wrapped line (item 44)
capture/vertical_pdf.py       vertical placement in the PDF export, a probe (item 46)
capture/screen_vertical.py    vertical placement on screen: window captures at zoom 100, a probe
capture/builtin_formats.py    what each built-in number format id shows, a probe
capture/screen_palette.py     section colors under a workbook palette, on screen, a probe
capture/angled_formats.py     the rows Excel fits for padded and filled angled text (item 56)
capture/font_heights.py       row heights of six faces at every size, wrapped, underlined (57, 58, 61)
capture/font_tables.py        the numbers read from the font files (post, OS/2, head, hhea)
capture/screen_rules.py       underlines, strikethrough and text color on screen (items 60, 62, 63)
capture/rule_table.py         src/xlformat/_rules_drawn.py, the drawn lines, from items 62 and 63
capture/courier_predict.py    item 63's predictions (capture/predictions/) and their scorer
capture/padding_predict.py    item 64's predictions for cell_font_large, and their scorer
capture/half_predict.py       item 67's predictions for column_half_points_*, and their scorer
capture/job.py                Excel for Windows: suites as capture jobs (manifest and workbooks),
                             and the suite files a job's result gives (items 65 to 72)
capture/windows_rows.py       Windows row heights fitted into src/xlformat/_rows_windows.py (item 69)
capture/screen_windows.py     window pictures on Excel for Windows as capture jobs (screen_* groups:
                             the Mac's screen suites for the faces with Windows tables), and their reader
capture/screen_table.py       the rules table on Excel for Windows: each rule by ppem, the rules side by
                             side in a row (group screen_rules_table), and its reader
capture/screen_insets.py      where a cell's text starts and ends by its font, on screen: the Mac's
                             run (in a macOS virtual machine) and the reader of both platforms
                             (group screen_insets)
capture/windows_screen_tables.py  Excel for Windows' rules and ink into src/xlformat/_screen_windows.py
                             (items 74, 78 to 80)
capture/excel_watch.py        a time limit on each request to Excel, and evidence when one passes
capture/edition.py            the Excel edition and update channel, for each capture's header
capture/invalid.py            which candidate format strings Excel refuses to open (item 9)
capture/repair.py             what Excel's repair does to refused formats (item 33): `open`
                             through Launch Services, a person answers the dialog, then
                             `read` finds the workbook by name; `write` gathers the TSV
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
saving and restores alerts on every exit path. A suite takes about a second of Excel time per
10,000 rows read, and about 0.75 s more for each column whose width the script sets (the testing
page gives the same rates).

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
  return `missing value`); only suites with `Suite.colors` fill the column, and seventeen color rows
  the API misreports are skipped by the test.
- **The 2023 capture stays**, gzipped, as a full-tier-only suite; the diff between captures is
  evidence about Excel builds and about the method.
- **Header metadata** per suite: Excel version and full build, edition and update channel,
  whether the display was kept awake, standard font, workbook font, locale, capture
  date, generator (the `capture.cases` function), capture method per column, tier rule. A regeneration that
  changes any expected text is a reviewed diff.

**Schema (frozen 2026-10-23).** Columns: `suite tag value_type value_repr format width
actual_width date1904 expected_text expected_color tier`; `#`-prefixed header lines carry the
metadata. Changing the schema rewrites every suite and its history, so it is a reviewed
decision, not a convenience.

`TEXT()` formulas are used only by the suites of its own (§2 item 38): its result differs from
cell display for `####`, and it has no width.

The TSV carries `actual_width` (the width Excel reports after setting it; `8.43` snaps to
`8.5`). Tiers are recomputable: `uv run python -m capture.retier <suite>|--all` reassigns the
`tier` column from `capture/tiers.py` (first row per output shape per format and width, at most
15 per group; `capture2023`, `capture2023_sample` and `capture2026` are full-tier only);
captures call the same function.
