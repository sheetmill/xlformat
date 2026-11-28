# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- **The default platform is Excel for Windows.** `platform=None` (every function that takes
  `platform`, and `NumberFormat`'s methods) now means `"windows"` (`DEFAULT_PLATFORM`), and
  `regional_settings=None` means `RegionalSettings.EN_US_WINDOWS`. Pass `platform="mac"` to keep
  Excel for Mac: with it every result is the same as before, apart from the rows added below.
- **Upgrade sheetpeek with this release.** sheetpeek 0.6.0 and earlier assume the Mac default and
  call `row_metrics` without a platform: with this release their Mac layouts take Windows' row
  heights, and a Helvetica cell raises `LookupError`.
- What a caller passing nothing now sees:
  - Excel mode (`column_width`): columns and text measured in whole pixels at 96 dpi, so `#` and
    fill counts can differ (`format_value(1234.5, "0.00*-", column_width=12)` is
    `'1234.50-------'`, on the Mac `'1234.50-----------'`), and a `*` fill that would pass 255
    characters stops at 255 instead of showing a `#` run. Without a width, and in character
    mode, the text is unchanged.
  - The measuring and column functions in points on Windows' pixel grid: `digit_width()` is 5.25
    (the Mac's 6.0), `text_width("Hello World")` 55.5 (55.0), `file_column_metrics().points`
    48.0 (53.0), `column_metrics(8.43).points` 48.0 (56.0); `autofit_width` takes the Windows
    bonus. `measured_fonts()` lists 25 families at 8 to 24 points, and `font_substitutes()` is
    empty, so text in Courier, Helvetica and MS Sans Serif raises `LookupError`.
  - Rows and placement: `row_metrics()` keeps a height of 15.0 but a wrapped line is 15 points
    (`line_height` 15.0, the Mac's 16.0). `row_metrics` and `vertical_metrics` cover every
    measured family and MS Sans Serif and Courier (below); Helvetica raises `LookupError`.
    `ink_metrics`, `rule_metrics` and `text_insets` cover Calibri, Aptos Narrow and Arial, and
    raise for other faces (on the Mac they cover more).
  - `builtin_format` gives the Windows short date for the 20 short-date ids: `builtin_format(14)`
    is `'m/d/yyyy'` (the Mac's `'m/d/yy'`), `builtin_format(22)` `'m/d/yyyy h:mm'`.
  - `text_function` is unchanged (`TEXT` gives the same text on both, CAPTURING item 72).

### Added
- `row_metrics` and `vertical_metrics` answer more faces (CAPTURING item 91). On Windows every
  measured family (Georgia and the 19 others that had none) and MS Sans Serif and Courier as the
  capture machine draws them, in four styles at every size from 1 to 72 points. Tahoma 11's height
  (`font_height`, the row it fits as the Normal font) is 14.25 points on Windows and 14 on the Mac,
  Book Antiqua 11's 16.5 and 15; under a Calibri 11 Normal font a one-line Tahoma 11 row is 15 on
  the Mac, captured, and 15 on Windows by the row model, not captured. On the Mac 15 more of the
  measured families at every size from 1 to 72 points: Aptos, Arial Narrow, Book Antiqua, Bookman
  Old Style, Calibri Light, Cambria, Century Gothic, Courier, Garamond, Helvetica, Lucida Console,
  Rockwell, Tahoma, Trebuchet MS and Verdana. Arial Black, Candara, Comic Sans MS, Consolas,
  Haettenschweiler and Impact have no Mac rows yet: their heights follow no one rule, and they still
  raise `LookupError` there. A face's name matches ignoring case (`"tahoma"` is Tahoma), as Excel's
  does on both platforms.

### Fixed
- `VerticalMetrics.cap_height` and `.descender` (deprecated) forward to the Mac's `ink_metrics`,
  the platform their `vertical_metrics` measured, whatever the default.
- `autofit_width`'s documentation names the two captured Mac ranges where the rule disagrees with
  Excel and the function still answers by it: from 1 to 5 points Excel's widths are multiples of
  5 points (172 of 1,080 captured texts differ) and at 383, 397 and 409 points 3,170 of 76,464
  differ, mostly by 1 or 2 points.

## [0.11.0] - 2026-11-25

### Added
- `autofit_width(text, font=None, *, normal_font=None, platform=None)`: the column width in
  points Excel's autofit gives one line of text in a cell's font, on the Mac and on Windows:
  the advances (on the Mac a text holding a non-ASCII glyph summed unrounded, as autofit does),
  the cell font's padding, twice the italic overhang and a long-text bonus counted in widths of
  the Normal font's digit, `max(0, floor(W / D / 5) - 2)` points on the Mac and
  `max(0, floor(W / D / 4) - 2)` pixels on Windows (CAPTURING item 90: every pixel to 45 digit
  widths in 13 Normal fonts). A carriage return, line feed or tab raises `ValueError`; an empty
  text is 0.0. `autofit_width("0" * 20)` is 129.0 points, and 112.5 on Windows.
- On Windows `row_metrics` and `vertical_metrics` answer Calibri, Aptos Narrow, Arial, Times New
  Roman and Courier New in four styles at every size from 1 to 72 points, half sizes included
  (CAPTURING item 89: every ppem from 1 to 96 captured, a size taking its ppem's height). 0.10.0
  raised `LookupError` below 6 points, above 40 points (Calibri: above 72) and at every half size
  except Calibri's from 6.5 to 39.5.

## [0.10.0] - 2026-11-24

### Added
- Excel for Windows measures 25 families, the Mac's 27 but Courier and Helvetica (which the
  capture machine has no font for), in four styles: every glyph of the 218 (space, ASCII and the
  cp1252 characters) in whole pixels at every ppem from 11 to 32 at 96 dpi, sizes 8 to 24 and the
  half sizes between, captured from Excel for Windows' autofit widths (CAPTURING item 85, which
  gives the autofit rule the tests check). A style a family has no font file for is the one
  Windows synthesizes, measured as drawn. `measured_fonts(platform="windows")` lists them at 8
  to 24; 0.9.0 measured Calibri, Aptos Narrow and Arial at 9 to 24 and raised `LookupError` for
  the other families. A face's advances are decoded one ppem at a time on first use.
- `row_metrics(platform="windows")` answers Calibri at 6 to 40 points and the half sizes
  between and at 41 to 72, and Aptos Narrow, Arial, Times New Roman and Courier New at 6 to 40,
  in four styles (CAPTURING item 86); it raised `LookupError` outside 9 to 24 points and for
  Times New Roman and Courier New.
- `SECURITY.md`, linked from the README: what xlformat's code covers, how to report a security
  problem privately through GitHub's private vulnerability reporting, where a problem in another
  sheetmill package goes, and which releases receive fixes (the latest).

### Changed
- On Windows a character outside ASCII measures by its own advance instead of the widest digit,
  and 8 points and every half size from 8.5 to 23.5 are measured instead of scaled from the
  nearest size. Sizes outside 8 to 24 points change too: they scale from the nearest measured
  ppem instead of the nearest measured size (ten digits of Calibri 25 measure 120.0 points, 127.5
  in 0.9.0; of Arial 6, 30.0 points, 37.5 in 0.9.0). Every committed rendering is unchanged.
- On the Mac an italic face's overhang is one rule per face, `floor(a * S + b)` points, captured
  at every whole size from 6 to 192 (CAPTURING item 87); 0.9.0 took the overhang between its
  captured sizes from the nearest one, scaled, and could be a point off (Verdana and Arial
  italic at 54 points: 11 and 10, now 10 and 9 as Excel shows). 2,126 of the 10,098 pairs of an
  italic face and a whole size from 6 to 192 (54 faces, 187 sizes) change. Below 6 and above 192
  the rule replaces the scaling, unverified either way. Text insets and the fit of italic text
  change at those sizes.
- Fit and `#` decisions above 192 points are captured in Calibri and Arial (194 to 318 points,
  CAPTURING item 88): Excel rounds each glyph, as the package does, also where Calibri's autofit
  sums a run unrounded (from 286 points). No result changes.

### Fixed
- `row_metrics`, `vertical_metrics`, `rule_metrics` and `ink_metrics` measure MS Sans Serif as
  Times New Roman on the Mac, as the other functions taking `font=` did; they raised
  `LookupError`.

## [0.9.0] - 2026-11-24

### Added
- Excel mode on the Mac measures 27 families: Calibri, Arial, Aptos Narrow, Aptos, Arial Black,
  Arial Narrow, Book Antiqua, Bookman Old Style, Calibri Light, Cambria, Candara, Century
  Gothic, Comic Sans MS, Consolas, Courier, Courier New, Garamond, Georgia, Haettenschweiler,
  Helvetica, Impact, Lucida Console, Rockwell, Tahoma, Times New Roman, Trebuchet MS and
  Verdana, each in four styles at every size, captured to 192 points and by the same rule,
  unverified, above (Excel allows 409; above 192 autofit stops rounding each glyph at a size
  that depends on the face, Calibri from 286). Each face is its units per em
  and one advance per character (space, ASCII and the cp1252 characters), recovered from Excel
  for Mac's own autofit widths (CAPTURING item 82), about 1 KB a face; a glyph's advance is
  `u * size / em` rounded half up to whole points.
- `font_substitutes(*, platform=None)`: names the platform has no face for and the face Excel
  measures them in. On the Mac, MS Sans Serif measures as Times New Roman, in every function
  that takes `font=`.

### Changed
- On the Mac, results change for Calibri, Aptos Narrow and Arial wherever 0.8.0 had no captured
  table. Every row captured from Excel renders as before.
  - A size other than 9, 10, 11, 12, 14, 16, 18, 20 and 24 points is measured by the rule, where
    0.8.0 scaled the nearest measured table and rounded each advance half to even. Calibri's and
    Aptos Narrow's digit at 13 points is 7 (0.8.0: 6; the 13-point check captures), so ten
    digits measure 70 points instead of 60 and a workbook whose Normal font is either at 13
    points has 7-point column units. On generated texts the difference reaches 16 points at 8
    to 36 points (the alphabet in Arial 13) and 42 points at 72 to 192.
  - A half size is measured as the whole size it rounds to, half up (10.5 as 11).
  - A character outside ASCII measured as the digit; now each cp1252 character measures by its
    own advance (`25°C ± 0.5` in Calibri 11: 46 points, 0.8.0 49), and a character beyond
    cp1252 as the widest digit.
  - Every size is answered, as in 0.8.0, past Excel's 409 points too, now by the rule where
    0.8.0 scaled: a Calibri 409 Normal font gives a 1,761-point default column (0.8.0: 1,735).
    A size under half a point measures 0 and a negative size a negative width, as in 0.8.0,
    and the column functions take a 1-point digit there where 0.8.0 raised
    `ZeroDivisionError`.
- `measured_fonts()` lists every measured family; its sizes are those at which every glyph's
  width was captured (Calibri, Aptos Narrow and Arial 9 to 14, 16, 18, 20, 24; the others 10
  and 13). `row_metrics`, `vertical_metrics`, `ink_metrics` and `rule_metrics` keep their own
  coverage, stated in each docstring (on the Mac Calibri, Aptos Narrow, Arial,
  Times New Roman, Courier New and Georgia), and raise `LookupError` for the
  other families.
- `text_insets` on the Mac covers every family `measured_fonts()` lists, from 1 to 72 points,
  by the widest digit of the cell's face (captured for Times New Roman, Verdana, Georgia,
  Courier New and Verdana italic as well, CAPTURING item 83), where it raised `LookupError` for
  all but Calibri, Aptos Narrow and Arial. Right-aligned italic text ends its overhang at the
  screen size further in (`right` and `accounting_right` grow by it), captured for Verdana,
  Calibri, Aptos Narrow and Arial italic (CAPTURING items 83 and 84); an accounting line still
  runs to the clear space. This changes the right insets of Calibri, Aptos Narrow and Arial
  italic from 0.8.0's: Calibri 11 italic 2 points to 4, Calibri 72 italic 9 to 22.
- The italic overhang on the Mac is measured at every whole size from 6 to 30 and at 32 to 192
  in steps, for every family.
- The bundled per-size Mac glyph tables are gone; `capture/glyph_fit.py` writes only Excel for
  Windows' tables, and `tests/test_glyph_tables.py` checks the Mac's measure files against the
  rule.

## [0.8.0] - 2026-11-23

### Added
- `text_insets(font, *, normal_font=None, zoom=100, platform=None)` returning `TextInsets(left,
  right, accounting_left, accounting_right)`: where a cell's text starts and ends, by the cell's
  font, on the Mac and on Windows (96 dpi), for Calibri, Aptos Narrow and Arial, in sheet points
  from the cell's boundaries: the left gridline and `ceil(D / 4)` screen units on the left, `ceil(D
  / 4)` on the right, `D` the cell font's digit width at its screen size; the Normal font does not
  move it.
- `vertical_metrics`, `ink_metrics` and `rule_metrics` work with `platform="windows"` (96 dpi),
  where they raised `LookupError`, for Calibri, Aptos Narrow and Arial, from window captures of
  Excel for Windows; results in points, whole pixels at the zoom.
  - `vertical_metrics`: placement in whole pixels from the face's height and descent, the
    sizes and styles of the Windows row tables.
  - `ink_metrics`: the ink of "H" and "g" at 9, 10, 11, 12, 14, 16, 18, 20 and 24 points, and
    bold, italic and bold italic at 11, 12 and 20.
  - `rule_metrics`: any whole zoom from 10 to 400 (a size and zoom under one pixel an em
    raises `LookupError`): the lines by a rule of the font's tables at the rounded ppem, with
    Calibri's and Arial's double underlines, Arial's accounting underline and every line up
    to 11 pixels an em from captured tables. A line may reach past its row: Excel draws it
    into the row below.

### Changed
- `ColumnMetrics.left_inset` and `right_inset` follow the Normal font, from the cell's
  boundaries as before: the left gridline and `ceil(digit_width / 4)` screen units of clear space
  on the left, the clear space on the right, where they were 3 and 2 points for every font and
  platform. Calibri 11 keeps 3 and 2 points on the Mac; on Windows it gives 2.25 and 1.5 points
  (3 and 2 pixels), and other Normal fonts give their own (Calibri 24: 4 and 3 points on the
  Mac).

## [0.7.0] - 2026-11-22

### Added
- `RowMetrics.accounting_height`: what the accounting underlines of a row's fonts add to its
  height, the difference from the same row without them (0 without one), as the captured rows
  with and without an accounting underline give it.
- Excel for Windows (16.0.20527.20196, Microsoft 365, en-US, 96 dpi): `platform=` on every
  function whose result can differ by platform, and `regional_settings=` on those that format or
  resolve built-in formats, with the frozen `Platform` (presets `Platform.MAC` and
  `Platform.WINDOWS`), `RegionalSettings` (presets `RegionalSettings.EN_US_MAC` and
  `RegionalSettings.EN_US_WINDOWS`) and `DEFAULT_PLATFORM`. `None` means the Mac and its en-US
  settings, so a call without them returns what it did in 0.6.0.
  - `platform="windows"` works in `format_value`, `format_cell`, `text_function`, the matching
    `NumberFormat` methods, `text_width`, `glyph_widths`, `digit_width`, `overhang`,
    `measured_fonts`, `column_metrics`, `column_metrics_from_points`, `file_column_metrics`,
    `row_metrics` and `builtin_format`, for Calibri, Aptos Narrow and Arial in four styles, with
    glyph tables in pixels captured from Excel for Windows; results stay in points. Both widths
    and rows cover the sizes 9, 10, 11, 12, 14, 16, 18, 20 and 24 points; rows without
    accounting underlines.
  - On Windows a column without a width rounds up to a multiple of 8 pixels, a fill stops at 255
    characters where the Mac shows a `#` run, and `builtin_format` gives `m/d/yyyy` for the 20
    ids whose en-US string is the short date (14, 27 to 31, 36, 50 to 58, 71, 72, and 22 and 78
    before ` h:mm`), from the regional settings' short date; only en-US is supported.
  - `vertical_metrics`, `ink_metrics` and `rule_metrics` raise `LookupError` for Windows, and
    every function raises it for Windows at a resolution other than 96 dpi.
- `rule_metrics` covers Courier New at screen sizes from 6 to 144 and raises `LookupError` outside
  them, where it raised for every size. Its double underline hangs up from the single
  underline's position, its gap captured at every screen size (no rule gives it), and its single
  sits on the double's first line once the gap is wider than two thicknesses.
- `row_metrics` gives Courier New's accounting rows from 6 to 72 points, where it raised
  `LookupError`: an accounting underline adds the face's double underline gap and its line
  thickness to the row and to each wrapped line. The same law gives every other face's
  accounting rows, which are unchanged.

### Changed
- A column whose width comes to a half point (Calibri 9 at width 1.5 is 12.5 points) is rounded
  up, as Excel for Mac and Excel for Windows both do; it was rounded to even, one point narrow
  at about half such widths.
- A cell font whose column padding `2·ceil(MDW/4)+1` is larger than the Normal font's (a digit
  of 9 points or more in a column of a smaller Normal font, such as Calibri 18 or Arial 16 cells
  in a Calibri 11 workbook) now loses the difference from its cell's budget in Excel mode, as
  Excel shows it: one `#` or one character fewer at some widths.
- `rule_metrics` at a screen size that is not whole (a half size, or a zoom such as 150) gives
  the lines at the nearest whole screen size, a half rounding up, as Excel draws them in most
  rows captured; 0.6.0 fed the fraction to the font's formulas, so values at such sizes can
  differ by a screen point.
- `rule_metrics` for Times New Roman at screen sizes 37 to 40 gives a double underline three
  line thicknesses wide, hanging up from the single's position, with the single on its first
  line, as captured; 0.6.0 gave this only at 38 and 40. At screen sizes 74, 75, 76 and 82 it
  gives the lines as now captured.

### Fixed
- `VerticalMetrics` has slots again, as in 0.5.0: in 0.6.0 its instances had a `__dict__`, so
  `vars()` worked on them; now it raises `TypeError`. Its six fields, equality, hashing, `repr`
  and pickling are unchanged.

## [0.6.0] - 2026-11-21

### Added
- `ink_metrics(font)` returning `InkMetrics(cap_height, descender)`: the measured ink of a face's
  "H" and "g", which raises `LookupError` for a face or size not measured. Measured: the six faces
  at 9 to 24 points, and Calibri and Aptos Narrow at 13, 15, 17 and 22.
- `Font(..., underline=None)` takes a file's underline name (`"single"`, `"double"`,
  `"singleAccounting"`, `"doubleAccounting"`). `row_metrics` reads it: an accounting underline
  adds `3 · max(1, floor(size · underline thickness))` points to the row and to each wrapped line.
  Times New Roman adds 4 points from 36.5 to 40 (captured); Courier New's accounting rows raise
  `LookupError`.
- `row_metrics` covers every size from 1 to 72 points, half sizes included, and three more faces:
  Times New Roman, Courier New and Georgia. A face's height is `floor(a · ceil(size) + 1.5)` with
  a slope per face from 7.5 points, and the captured heights below that.
- `rule_metrics(font, *, zoom=100)` returning `RuleMetrics`: where Excel draws underlines,
  accounting underlines and a strikethrough, for Calibri, Aptos Narrow, Arial, Times New Roman and
  Georgia (Courier New raises). It depends on the screen size `size · zoom / 100` alone: the
  lines as captured at the screen sizes measured, and the font's own tables elsewhere.
  `RuleMetrics.double_top` and `.double_thickness` give a double underline's first line, which
  moves up a line at a time until the double clears the accounting underline, and its lines'
  thickness, which in Arial is thinner than a single underline's.

### Changed
- `vertical_metrics` returns placement only (baselines, font height, line height, line pitch, line
  inset), by rule for every face and size `row_metrics` supports: six faces from 1 to 72 points.
  The ink fields moved to `ink_metrics`. `cap_height` and `descender` are no longer fields of
  `VerticalMetrics`: constructing one takes six values, and `dataclasses.fields`, `asdict`,
  `repr` and equality show the six; `dataclasses.replace` no longer accepts `cap_height` or
  `descender`, and its result, like a `VerticalMetrics` constructed directly, does not know its
  font, so its deprecated properties raise `LookupError`. For a face or size that placement
  covers but `ink_metrics` has not measured (Calibri 30, Times New Roman 30), `vertical_metrics`
  now returns where it raised `LookupError`, and reading `cap_height` or `descender` on the
  result raises `LookupError`.
- `row_metrics` takes a row of several fonts as the largest `size + 1` plus the largest
  `height − size − 1`, the rule behind the earlier tables, which every captured row height follows;
  its results for the faces and sizes measured before are unchanged.

### Fixed
- Text under a format of one colored General section, such as `[Blue]General`, takes that color,
  as Excel shows it; `FormattedCell.color` was `None`.

### Deprecated
- `VerticalMetrics.cap_height` and `.descender` warn and forward to `ink_metrics`; they are removed
  in 1.0.

## [0.5.0] - 2026-11-19

### Added
- Arial is measured: glyph tables for regular, bold, italic and bold italic at 9, 10, 11, 12, 14,
  16, 18, 20 and 24 points, so Excel mode, `text_width`, `column_metrics` and
  `file_column_metrics` take an Arial Normal font or cell font, and `row_metrics` gives Arial
  row heights (Arial 10: 13 points; its default column is 53 points), and `vertical_metrics`
  where its text sits in the row.
- `vertical_metrics(font, *, normal_font=None)`: given the workbook's Normal font,
  `baseline_from_bottom` includes Excel's rule that bottom-aligned text sits at least the Normal
  font's distance up. `None`, the default, means the cell's own font.
- Documentation: text turned to an angle shows the width-less rendering (its `_x` padding, no fill,
  never `#`), captured at 15, 30, 45, 60, 75 and 90 degrees and at -45 and -90 from the rows Excel
  fits (`angled_formats.layout.tsv`).

### Changed
- `vertical_metrics(("Calibri", 9)).baseline_from_bottom` without `normal_font` is 2 points, was 3:
  Calibri 9's own placement, as in an Arial 10 workbook. A caller passing the workbook's Normal
  font sees no change: with `normal_font=("Calibri", 11)` it is 3, as before.
- `row_metrics` takes a row of several faces as the largest ascent plus the largest descent,
  rounded up, rather than the tallest face. Faces of one family give the same heights as before;
  a row mixing Arial with Calibri or Aptos Narrow can be a point taller than its tallest face
  (Arial 11 beside Calibri 10: 15 points, where each alone is 14).

### Fixed
- `General` in a column too narrow for it by glyph width now shows the `#` run, as Excel does,
  where it had shown a form that fit by character count alone. Calibri and Aptos Narrow never
  reach this case; in Arial, whose `E` and `+` are wider than a digit, `1234567.891` at 5 units
  (Arial 10) is `#####`, not `1E+06`.

## [0.4.1] - 2026-11-17

### Fixed
- `vertical_metrics` placed text one point too low. `baseline_from_top` is now the face's size
  plus 1 (Calibri 11: 12 points, was 11) and `baseline_from_bottom` one point less at every size
  (Calibri 11: 3 points, was 4). The other fields are unchanged. The screen capture had taken each
  row's top 2 pixels below the top of the gridline above it, which is where a row begins.
- The known deviation listed in 0.4.0 for a file column width of 10.7109 at a 24-point Normal
  font is gone: the file held 10.7109375, which the expected outputs recorded with six digits, and
  Excel converts it as ECMA-376's formula does, as the package always has. No rendering changes.
  Expected-output files now record such widths in full, and new suites check file widths on
  either side of a 1/256 step, columns without a width under other base widths, and fills past
  255 characters.
- Documentation: the column padding and the digit widths are stated for all nine measured sizes;
  the text edges are described as measured in Excel's print layout at Calibri 11 (not as
  provisional); Microsoft's page on regional date and time formats is cited at its current
  address.

## [0.4.0] - 2026-11-16

### Added
- `section_rgb(color, palette=None)`: the `RRGGBB` Excel draws a section color in under a
  workbook's own palette: `[ColorN]` takes the palette's entry N + 7, while the eight color names
  keep their standard colors (captured from Excel's screen).
- `builtin_format(number_format_id)` and `BUILTIN_FORMATS`: the format string Excel uses for
  each built-in id from 0 to 163, as Excel for Mac reports it in en-US. Id 14, which Excel writes
  for typed dates, is `m/d/yy` (a typed date shows `1/1/24`); openpyxl's table has `mm-dd-yy` there
  and differs at 44 and 47 too.
- Glyph tables for Calibri and Aptos Narrow at 18, 20 and 24 points in four styles, measured
  in Excel. Text at those sizes is measured exactly, and the sizes between scale from the
  nearest of them instead of from 16 points.
- `vertical_metrics(font=None)` and `VerticalMetrics`: where Excel draws a face's text in a
  cell (the baseline from the top and from the bottom, the line pitch, the cap height and the
  descender), for Calibri and Aptos Narrow at nine sizes from 9 to 24 points (captured from
  Excel's screen).
- `show_zeros=False` on `format_value`, `format_cell` and the `NumberFormat` methods, for a
  sheet that hides zeros (`showZeros="0"`): a number equal to 0 shows nothing under a format
  without a zero section, as Excel shows it (captured from Excel).
- `row_metrics(normal_font=None, *, cell_fonts=(), wrapped=())` and `RowMetrics`: the height
  Excel gives a row without a stored height, from its Normal font, its cells' fonts and its
  wrapped lines, for Calibri and Aptos Narrow at nine sizes from 9 to 24 points (captured from
  Excel).

### Changed
- A line feed and a tab in text measure as a space, and a carriage return as nothing (CR LF
  is one space), as Excel fits and fills them; they measured as a digit before. Fits, `#`
  counts and fill counts of texts holding them change (captured from Excel).
- A column's padding grows with the Normal font's digit width: 5 points up to a digit width
  of 8 (Calibri 16) and 7 from 9 to 12 (Calibri 18 to 24), `2 * ceil(MDW / 4) + 1`. Columns,
  budgets and General's reduction in workbooks with a Normal font of 18 points or more change;
  the earlier sizes are as before.
- One more known deviation, listed in `QUIRKS.md`: a file column width of 10.7109 at a 24-point
  Normal font (a digit width of 12) is 129 points in Excel, one more than the 128 that ECMA-376's
  formula, which the package follows, gives.
- The expected outputs are captured from Excel for Mac 16.115 under Microsoft 365 (Excel's
  standard font there is Aptos Narrow 12). Every suite matched its earlier capture from Excel
  16.114 under Office Home & Business 2021, so no rendering changed.

## [0.3.0] - 2026-11-03

### Added
- `text_function(value, format_string, *, date1904=False, quirks=None)` and
  `NumberFormat.text_function`, what Excel's `TEXT` function returns: a blank is 0, an error passes
  through, a refused format or a `#` run is `ExcelError.VALUE`, and the empty format gives the
  empty text for a number and `-` for a negative one.
- `column_metrics_from_points(points, normal_font=None)`, the column of a width in points. A merged
  cell formats as one column of the sum of its columns' points, with a single padding of 5, and a
  hidden column counts as 0 (suites `merged11`, `merged14`).
- The Buddhist-year code: `b` and `bb` show the year's last two digits and `bbb` and longer all
  four, in either case (`abc` with 12.5 shows `a43c`); after a `[$...]` tag they are the plain
  year. The era code `g`, empty in en-US. The calendar prefixes `b1` and `b2`, which print
  nothing. Other calendars are localization, which stays out of scope. Without a width the code
  shows the Buddhist year, as the sheet and `TEXT` do; Excel's clipboard alone gives the plain year,
  a known deviation for that medium, not modeled.

### Changed
- The formula engine is named `xlrecalc` throughout the documentation. xlformat keeps
  `ExcelError`; `xlrecalc` has its own `ErrorValue`, which compares equal to it.
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
- `text_width` and `glyph_widths` give whole points, and fit and fill decisions use them, as Excel
  does: ten `W` at Calibri 11 measure 100, not 101, and `*W0` holds as many as whole points allow.
  Autofit alone counts the tenth some wide glyphs carry. Evidence: suites `advance_tenths11`,
  `advance_tenths12`.

### Deprecated
- `font=` on `file_column_metrics` and `column_metrics`, which names the workbook's Normal font
  there and the cell's font in the format functions: pass `normal_font=`. A second positional
  argument to `column_metrics` keeps working and means the Normal font. `font=` warns with a
  `DeprecationWarning` and goes away in 1.0; passing both raises `ValueError`.

### Fixed
- A point and one to three zeros after any date or time code are fractions of a second, as in
  Excel: `[s].00` raised `FormatError`, and so did `[h].00`, `yyyy.00` and `s x.00` (suite
  `elapsed_decimals`). Beside an elapsed code `%`, `AM/PM` and `/` before a digit now raise it.
- A number through `@`, `[Red]General` or `General;General` shows nothing in a column too narrow
  for one `#`; only the plain `General` shows one `#`. A negative through such a format with less
  than one character beside its sign shows `0` or `#` by its magnitude, as captured.
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

[Unreleased]: https://github.com/sheetmill/xlformat/compare/v0.11.0...HEAD
[0.11.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.11.0
[0.10.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.10.0
[0.9.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.9.0
[0.8.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.8.0
[0.7.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.7.0
[0.6.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.6.0
[0.5.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.5.0
[0.4.1]: https://github.com/sheetmill/xlformat/releases/tag/v0.4.1
[0.4.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.4.0
[0.3.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.3.0
[0.2.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.2.0
[0.1.0]: https://github.com/sheetmill/xlformat/releases/tag/v0.1.0
