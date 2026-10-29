# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
- Glyph tables for bold, italic and bold italic Calibri and for Aptos Narrow in all four styles,
  each at 9, 10, 11, 12, 14 and 16 points, measured in Excel for Mac. An italic cell's text area
  is a point or two smaller, as in Excel.

### Fixed
- Excel mode: a column too narrow for one `#` shows nothing for a number that does not fit
  (the General format still shows one `#`), and dialog widths below one unit have Excel's
  width in points.
- A `#` run is at most 255 characters, and a number or text that would exceed 255 characters,
  with its `*` fill or without a width, shows a `#` run, as in Excel.
- With a width, a negative that `General` reduces to zero shows `0`, not `-0`, under any General
  other than the plain format (`[Red]General`, `@`, a General section among others).
- Excel mode: `General` at fractional column widths. When no shortened form fits by
  characters, Excel shows the shortest form if its glyphs fit (`1E+06` in a column of 4.67
  units); among literals, a nearly full column shows `0` for a small value and the `#` run
  otherwise; a leading minus costs its own glyph width (half a digit at Calibri 11 only).
  Checked against 52,084 new width rows at every whole-point column width.

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

[Unreleased]: https://github.com/sheetmill/xlformat/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.1.0
