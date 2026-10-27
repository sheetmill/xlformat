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
   Excel build, the fonts, the locale, the date and the generator's commit.

The test suite then renders every row with xlformat and compares the text exactly. A default
`pytest` run checks every captured row in a few seconds. Width rows are checked twice: in Excel
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
| `numeric_core`, `numeric_core_width` | 10,944 and 32,832 | number sections, conditions, grouping, scaling, at three widths |
| `fraction`, `fraction_ties`, `fraction_threshold`, `fraction_random` | 33,452 | denominators, half steps, the nudge's bounds, random doubles |
| `datetime`, `date1904`, `clock_boundaries`, `clock_formulas` | 4,794 | date codes, the 1904 system, rounding at second boundaries |
| `scientific`, `rounding_15digit`, `parser_check` | 1,695 | exponents, the fifteen-digit rule, the parser |
| `width_general`, `width_general14`, `width_fill`, `text_overflow`, `space_fill9` to `16` | 4,986 | `General` reduction, fills, overflow, the space at six sizes |
| `colors`, `lone_zero`, `conditions_rounding`, `selection_fallback` | 493 | section colors and section selection |
| `measure9` to `measure16` | 523 texts each | autofit widths from which the glyph tables are fitted |

Every row passes. Fifteen color readings are skipped because Excel's scripting interface
reports automatic black where the sheet, checked by eye, shows the section color.

The full record of every probe, including the hypotheses that turned out wrong, is
[CAPTURING.md](https://github.com/sheetmill/xlformat/blob/main/CAPTURING.md) in the repository. The
rules themselves are summarized in
[ARCHITECTURE.md](https://github.com/sheetmill/xlformat/blob/main/ARCHITECTURE.md).

## Regenerating

Regenerating a suite needs Excel for Mac and permission for the terminal to control it:

```bash
uv run python -m capture numeric_core
```

A 10,000-row suite takes about a second of Excel time. The scripts never touch a workbook they
did not create.
