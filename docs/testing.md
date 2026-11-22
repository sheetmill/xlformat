# How the output is tested

Excel's documentation of number formats is incomplete and, in places, wrong, so xlformat treats
Excel itself as the specification. Every expected output in the test suite was captured from
Excel 16 for Mac, and every rule in the code cites the captured rows that pinned it.

## Capturing from Excel

A suite is a list of cases, each a value, a format string and the column widths to read it at,
defined in `capture/cases.py`. Capturing a suite takes four steps:

1. **Build.** openpyxl writes the cases into a workbook, one row per case and one column per
   width, with the format applied as the cell's real number format.
2. **Open.** An AppleScript opens the workbook in Excel for Mac, inside Excel's sandbox
   container so that no permission dialog appears, with alerts switched off and restored on
   every exit path.
3. **Read.** Width-less text comes from copying the column to the clipboard, which is Excel's
   own rendering without a width. Text at a width comes from setting the column width and
   reading each cell's displayed text, which reflects `#` overflow, `*` fills and `General`
   digit reduction. Section colors are read from each cell's display format, and autofit
   widths from Excel's own measurement of a text.
4. **Write.** The texts are written to `tests/captures/<suite>.tsv` with a header recording the
   Excel build, the fonts, the locale, the date and the generator function.

The test suite then renders every row with xlformat and compares the text exactly. A default
`pytest` run checks every captured row in seconds. Width rows are checked twice: in Excel
mode, which must match exactly, and in character mode, which skips the rows where Excel fitted
a text by glyph width.

## Getting the input right

A test against Excel is only as good as the value Excel was given. openpyxl writes floats with
16 significant digits, which for about half of all doubles names a neighboring double, and
three early rules were fitted to those neighbors before this was found. The builder now writes
the shortest exact text of each double. A check suite reads the loaded value back bit for bit
through formulas, confirming that Excel's parser is correctly rounded. Cases that must pin a
computed double are written as formulas, which Excel evaluates in IEEE arithmetic.

## Evidence over guesses

Each behavior was probed before it was coded, usually with a small suite built for one
question: how a tie at the fifteenth digit rounds, which section a value reaches when two
conditions fail, how many points a hyphen takes at 12 points. Some findings came from
measurements with many digits rather than many rows. A `General` token inside a date format
leaks Excel's intermediate seconds to nine digits, which fixed the exact order of Excel's
floating-point operations for times. The fill count of an accounting format fixed the width of
a space, which autofit cannot resolve.

The fixed-denominator fraction rule came from a larger capture of half-step values, included as
`fraction_threshold` and `fraction_random`. The first wrote each value both as text and as a
formula and pinned the rule; the second holds 30,300 random rows whose texts were predicted
before Excel was asked, and every prediction matched.

## The suites

| suite | rows | what it covers |
|---|---|---|
| `capture2023`, `capture2026`, `capture2023_sample` | 89,680, 89,680, 1,993 | 152 formats across 584 values, captured in 2023 and again in 2026; no differences |
| `numeric_core`, `numeric_core_width`, `numeric_core_width_aptos_narrow11`, `numeric_core_width_arial10` | 109,440 | number sections, conditions, grouping, scaling, at three widths, in Calibri, Aptos Narrow and Arial |
| `fraction`, `fraction_ties`, `fraction_threshold`, `fraction_random` | 33,452 | denominators, half steps, the nudge's bounds, random doubles |
| `datetime`, `date1904`, `clock_boundaries`, `clock_formulas` | 4,794 | date codes, the 1904 system, rounding at second boundaries |
| `scientific`, `rounding_15digit`, `parser_check` | 1,695 | exponents, the fifteen-digit rule, the parser |
| `width_general`, `width_general14`, `width_fill`, `text_overflow`, 108 `space_fill` suites | 25,386 | `General` reduction, fills, overflow, the space of every measured face |
| `general_points9` to `24`, `general_points_aptos_narrow11`, `general_points_arial10`, `11`, `general_literals_points`, `general_zero_sign` | 85,334 | `General` at every whole-point column width, among literals, negatives that reduce to zero |
| `column_file9` to `24`, `column_file_aptos_narrow11`, `18`, `20`, `24`, `column_file_arial9` to `24`, four `column_default_*`, `column_narrow11`, `column_narrow14`, `limits` | 59,144 | widths written into the file, columns without a width, columns below one unit, the 255-character limit |
| 16 more `column_default_*` (base widths 6, 10 and 12 at 9 to 24 points), `column_boundary11`, `16`, `24`, `long_fills` | 8,098 | a column without `<col>` under other base widths; file widths on either side of a 1/256 step; fills past 255 characters |
| `cell_font`, `general_cell_font`, `cell_font_numeric`, `cell_font_styles`, `general_cell_font_styles`, `cell_font_arial`, `general_cell_font_arial` | 161,760 | cells whose font differs from the Normal font, in size, style and family |
| `general_exponent3`, `general_exponent_sweep`, `subnormal` | 68,838 | `General` with long exponents, and doubles below the smallest normal one |
| `general_text_narrow`, `general_sign_narrow` | 4,026 | a number through `@` or `[Red]General` in narrow columns, and negatives with less than a character beside the sign |
| `merged11`, `merged14` | 5,940 | cells merged across columns, in Calibri 11 and 14 |
| `advance_tenths11`, `advance_tenths12` | 41,481 | long runs and fills of the glyphs that autofit measures with a tenth of a point |
| `error_width` | 1,638 | error values under formats and widths |
| `colors`, `lone_zero`, `conditions_rounding`, `selection_fallback` | 493 | section colors and section selection |
| `letters`, `letters_upper`, `letters1904` | 6,584 | what a bare letter means in a format, and the Buddhist-year code |
| `elapsed_decimals`, `huge_values`, `empty_text`, `fraction_percent` | 3,678 | fractions of a second after any date or time code, values too large for fraction and elapsed arithmetic, the empty text, a numerator's limits |
| `hidden_zeros`, `hidden_zeros_control` | 1,700 | a sheet that hides zeros (`show_zeros=False`) and the same cases on one that shows them: exact zeros, values rounding to zero, zero sections |
| `line_breaks`, `line_breaks_aptos_narrow` | 648, 648 | carriage returns, line feeds, CR LF and tabs in text under fills of spaces and digits, at four widths, with autofit widths beside |
| `text_function`, `text_function1904` | 29,986 | Excel's `TEXT` over every format of the other suites, the empty format and those Excel refuses in a file |
| `measure9` to `measure16` and 102 `measure_<family>_<style><size>` | 523 texts each | autofit widths from which the glyph tables are fitted: Calibri, Aptos Narrow and Arial, four styles, nine sizes |

Thirteen further files are not suites of the common schema: `alignment.layout.tsv` (196 rows, read
from Excel's PDF export) holds where Excel places each value, `repair.layout.tsv` (248 rows)
what Excel's repair does to formats it refuses, and `row_height.layout.tsv.gz` (35,100 rows) the
height Excel gives each row for its Normal font, its cell fonts and its wrapped lines, and
`vertical.layout.tsv` (2,340 rows, measured in window captures) where the text sits in the row,
`angled_formats.layout.tsv` (128 rows) the rows Excel fits for padded and filled text at eight angles,
`font_height.layout.tsv` (1,074 rows) and `wrapped_height.layout.tsv` (144 rows) the rows Excel fits
for six faces at every size from 1 to 72 points, beside Calibri and wrapped,
`underline_height.layout.tsv` (2,814 rows) the same under each underline style,
`accounting_height.layout.tsv` (132 rows) Times New Roman and Georgia under accounting underlines
at every half point from 36 to 41,
`rules.layout.tsv` (396 rows, from window captures) where underlines and strikethrough are drawn
and what color text takes, `rules_double.layout.tsv` (1,850 rows, from window captures) the same
in five faces at every size from 9 to 36 points and up to 72 in steps of 4,
`builtin_formats.layout.tsv` (164 rows) what each built-in number format id stands for, and
`palette.layout.tsv` (128 rows, from window captures) the colors drawn under a workbook palette. In all, 334 suite files (226 with rows, the
others autofit measurements) hold 836,116 rows. Every row passes. Fifteen color readings are
skipped because Excel's scripting interface reports automatic black where the sheet, checked by
eye, shows the section color, and width-less rows of a Buddhist-year code are compared with the
plain year, which is what Excel's clipboard gives where the sheet shows the Buddhist year.

The full record of every probe, including the hypotheses that turned out wrong, is
[CAPTURING.md](https://github.com/sheetmill/xlformat/blob/main/CAPTURING.md) in the repository. The
rules themselves are summarized in
[ARCHITECTURE.md](https://github.com/sheetmill/xlformat/blob/main/ARCHITECTURE.md).

## Regenerating

Regenerating a suite needs Excel for Mac and permission for the terminal to control it:

```bash
uv run python -m capture numeric_core
```

A suite takes about a second of Excel time per 10,000 rows, and about 0.75 s more for each column
whose width the script sets. The scripts never touch a workbook they
did not create.
