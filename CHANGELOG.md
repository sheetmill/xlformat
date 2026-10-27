# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
