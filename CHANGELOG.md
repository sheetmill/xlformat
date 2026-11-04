# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- The formula engine is named `xlrecalc` throughout the documentation. xlformat keeps
  `ExcelError`; `xlrecalc` has its own `ErrorValue`, which compares equal to it. The
  `sheetmill` meta-package is released and installs all three packages.
- The empty string is ordinary text, as in Excel: `@" units"` with `""` shows ` units`,
  `"pos";"neg";"zero";"text"` shows `text` and `@*-` fills a column. It rendered nothing. A blank
  cell is `None`, which still does. Evidence: cells holding `=""` (suite `empty_text`).
- More formats raise `FormatError`, because Excel refuses to open them: `n` or `N` outside quotes
  (`0 n`), `E` without a sign and digits (`E`, `yyyy E+`), `b` or `g` with digit placeholders
  (`0b`, `g0`) or `@` (`@b`), a format ending in a lone code `B` (`0;B`), a `b1` or `b2` after
  other codes or with digit placeholders, and any other digit after `b` (`b3`). They rendered
  before. Evidence: one workbook per form, `capture/letters.py` (59 of the 73 refused forms
  captured there were new to the parser).
- A fraction's numerator shows a `#` run above 32767 with a fixed denominator, as in Excel (`?/10`
  with 3276.8, `?/10%` with 50), and from 2147483647 with a free one. Those values rendered the
  fraction. Evidence: suite `fraction_percent`.

### Added
- `text_function(value, format_string, *, date1904=False, quirks=None)` and
  `NumberFormat.text_function`, what Excel's `TEXT` function returns: a blank is 0, an error passes
  through, a refused format or a `#` run is `ExcelError.VALUE`, and the empty format gives the
  empty text for a number and `-` for a negative one.
- The Buddhist-year code: `b` and `bb` show the year's last two digits and `bbb` and longer all
  four, in either case (`abc` with 12.5 shows `a43c`); after a `[$...]` tag they are the plain
  year. The era code `g`, empty in en-US. The calendar prefixes `b1` and `b2`, which print
  nothing. Other calendars are localization, which stays out of scope.

### Deprecated
- `font=` on `file_column_metrics` and `column_metrics`, which names the workbook's Normal font
  there and the cell's font in the format functions: pass `normal_font=`. A second positional
  argument to `column_metrics` keeps working and means the Normal font. `font=` warns with a
  `DeprecationWarning` and goes away in 1.0; passing both raises `ValueError`.

### Fixed
- A huge value no longer raises under a fraction or elapsed format: `# ?/?` with 1e100 raised
  `decimal.InvalidOperation`, and `[h]` and `?/?%` with the largest double raised
  `OverflowError`. Their digits past 15 print as zeros, as in Excel.
- A clock section reached as the negative section shows a `#` run for a negative (`0;yy` with
  -0.5), as in Excel; it showed the magnitude's date.
- `General` in a date section reduces at a width like a General among literals.
- General keeps four mantissa decimals, not five, from an exponent magnitude of 99 (Excel shows
  `2.2251E-308` and `1.2346E+99`), and counts such an exponent as three digits when it reduces at
  a width. A negative with no room for a scientific form shows `-0` where that fits and a `#`
  run where it does not.
- A number below the smallest normal double is zero, as in Excel, which has no subnormal
  numbers (`1e-310` under General shows `0`).
- `ColumnMetrics.budget` is 0 for a column under 5 points instead of negative
  (`file_column_metrics(0)` gave -5). A column of no width has a `width` of 0, which
  `column_width` refuses with `ValueError`.

## [0.2.0] - 2026-10-28

### Added
- `file_column_metrics` converts a column width as a file stores it (`<col width>`, or the
  `<sheetFormatPr>` default for a column without one) into the column Excel for Mac lays out,
  and `column_metrics` does the same for a Column Width dialog number. Both return a
  `ColumnMetrics` with the width in dialog units (what `column_width` takes), the width in
  points, the budget of a number, the digit width, and the text edges (provisional values from
  Excel's print layout). A column openpyxl leaves without a width is 8 units in Excel for Mac,
  and one it writes as 8.43 is 7.67.
- `normal_font` on `format_value`, `format_cell` and `NumberFormat.format` / `format_cell`: the
  workbook's Normal font, which sets the column's units, when the cell's font (`font`) differs;
  it defaults to `font`. `Font(name, size, bold, italic)` names a font.
- Text measurement as Excel for Mac does it: `text_width`, `glyph_widths` (per character, for
  runs), `digit_width`, `overhang` (an italic face's slant) and `measured_fonts()`, all taking
  `font=` like the format functions.
- `ExcelError`, every error value current Excel documents (`#DIV/0!`, `#N/A`, ... `#SPILL!`,
  `#CALC!`, `#TIMEOUT!`), accepted as a value: the error's text under any format, centered, in
  the font's color, and a `#` run at a width too narrow for it. An unlisted text shaped like an
  error code is accepted too, so a newer Excel's error renders.
- Glyph tables for bold, italic and bold italic Calibri and for Aptos Narrow in all four styles,
  each at 9, 10, 11, 12, 14 and 16 points, measured in Excel for Mac. An italic cell's text area
  is a point or two smaller, as in Excel.

### Fixed
- Excel mode: a column too narrow for one `#` shows nothing for a number that does not fit
  (the General format still shows one `#`), and dialog widths below one unit have Excel's
  width in points.
- A `#` run is at most 255 characters, and a number or text that would exceed 255 characters,
  with its `*` fill or without a width, shows a `#` run, as in Excel.
- `FormattedCell.align` is `left` for a number under a format with only a text section, such
  as `@`, as Excel aligns it; it was `right`.
- With a width, a negative that `General` reduces to zero shows `0`, not `-0`, under any General
  other than the plain format (`[Red]General`, `@`, a General section among others).
- Excel mode: `General` at fractional column widths. When no shortened form fits by
  characters, Excel shows the shortest form if its glyphs fit (`1E+06` in a column of 4.67
  units); among literals, a nearly full column shows `0` for a small value and the `#` run
  otherwise; a leading minus costs its own glyph width (half a digit at Calibri 11 only).
  Checked against 52,084 new width rows at every whole-point column width.

### Changed
- Faster and smaller: the four-section accounting format compiles in 42.5 µs (61.9 in 0.1),
  Excel mode with a cell font renders in about 6 µs (10), and a compiled format retains about
  1.1 KiB (3). No rendering changed: every captured row renders byte for byte as before.
- Tested against 526,729 rows in 101 suites captured from Excel 16.114 for Mac (en-US), with
  the 2023 capture, plus Excel's alignment from its PDF export and what its repair does to
  formats it refuses: a refused format becomes `General` in its cell only. Every row passes; no
  known deviations.

### Known limits
- Excel mode bundles Calibri and Aptos Narrow; other sizes are scaled from the nearest
  measured size, and other fonts raise `LookupError`. Widths are those of Excel for Mac;
  Windows Excel is not measured.
- The 255-character limit is taken from Excel's scripting interface, and the text edges in
  `ColumnMetrics` from its print layout; neither is yet checked on screen.
- Where an error type shared with the other sheetmill packages lives is provisional.

## [0.1.0] - 2026-10-27

First release.

### Added
- `format_value(value, format_string)` returns the text Excel displays; `format_cell` returns
  a `FormattedCell` with the text, the width-less text, the section color, the default
  alignment, the `*` fill position and whether the value overflowed. `NumberFormat.parse`
  compiles a format once (cached) and exposes `format`, `format_cell`, `is_date`, `is_text`,
  `is_general` and the sections.
- Values: `int`, `float`, `Decimal`, `bool`, `str`, `None`, `date`, `datetime`, `time` and
  `timedelta`. The 1900 and 1904 date systems (`date1904=True`).
- Every section kind: numbers with grouping, scaling, percent and literals; `General`;
  scientific; fractions with free and fixed denominators; dates, times and elapsed time;
  text; up to four sections with conditions. Format strings Excel refuses to open from a file
  raise `FormatError`.
- Section colors by name and `Color1` to `Color56`; `DEFAULT_PALETTE` maps them to `RRGGBB`.
- Width in two modes. Character mode (`width`) counts characters for terminals and plain
  text. Excel mode (`column_width`, `font`) measures glyphs as Excel for Mac does, for `#`
  overflow, `*` fills and `General` digit reduction, with Calibri measured at 9, 10, 11, 12,
  14 and 16 points. A custom `measure` can replace the character count.
- Three switchable Excel quirks, all on by default: `leap_year_1900`,
  `fifteen_digit_precision` and `fraction_denominators`; `Quirks.STRICT` turns them off.
- Tested against 270,549 rows in 30 suites: 180,869 captured from Excel 16.114 for Mac
  (en-US), and the 89,680 rows of a 2023 capture, which Excel 16.114 reproduces exactly. Every
  row passes; no known deviations.

### Known limits
- Excel mode bundles Calibri only; other Calibri sizes are scaled from the nearest measured
  size, and other fonts raise `LookupError`.
- Expected outputs come from Excel for Mac. Month and day names are English; locale tags are parsed but
  not applied.
- Error values such as `#N/A`, conditional formatting and cell styles beyond the format string
  are not handled.

[Unreleased]: https://github.com/sheetmill/xlformat/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.2.0
[0.1.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.1.0
