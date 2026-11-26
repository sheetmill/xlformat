# Width

Excel decides what a cell shows from the column width: a `*` fill repeats to the edge, a number
or date that does not fit becomes a run of `#`, and `General` drops digits before it overflows.
xlformat applies the same rules in one of two modes.

## No width

Without a width, the result is what copy and paste from Excel gives: `*` fills add nothing,
`_x` pads add one space, `General` uses up to eleven characters, and a value a section cannot
show is 255 `#`.

## Character mode: `width`

`width` is a number of characters. It suits terminals and fixed-width text.

```python
>>> from xlformat import format_value
>>> format_value(1234567.891, "General", width=7)
'1234568'
>>> format_value(12345678, "General", width=5)
'1E+07'
>>> format_value(1234.5, "0.00*-", width=12)
'1234.50-----'
>>> format_value(-1, "m/d/yy", width=8)
'########'
>>> format_value("Hello World", "@", width=3)
'Hello World'
```

Text never overflows; clipping or spilling it into the next cell is the caller's job. Text
through a section whose `*` fill follows it is the exception, as in Excel. Character mode is
its own contract, not an approximation of Excel: a `#` run is always `width` long, and a value
fits when its character count does.

## Excel mode: `column_width` and `font`

`column_width` is the number from Excel's Column Width dialog. The text is measured with the
glyph widths of `font`, as Excel measures it on the platform asked for (Excel for Windows by
default, Excel for Mac with `platform="mac"`), so the result matches the sheet character for
character, including the number of `#` and fill characters. The rules are the same on both; the
Mac measures in whole points and Windows in whole pixels. The examples from here to
[Excel for Windows](#excel-for-windows) pass `platform="mac"`, whose whole points make the
arithmetic easy to follow; that section gives the Windows numbers.

```python
>>> format_value(1000, "#,##0.0", width=6)
'######'
>>> format_value(1000, "#,##0.0", column_width=6, platform="mac")
'1,000.0'
>>> format_value(1234.5, "0.00*-", column_width=12, platform="mac")
'1234.50-----------'
>>> format_value(1234.5, "0.00*-", column_width=12, font=("Calibri", 14), platform="mac")
'1234.50---------'
```

`font` is the cell's font: a name, a `(name, size)` pair or a `Font`; the default is
`("Calibri", 11)`, the Normal style openpyxl writes. On the Mac, 27 families (`measured_fonts()`)
are measured in regular, bold, italic and bold italic at any size (captured to 192 points; above
that, to Excel's 409, by the same rule, unverified), a half size as the whole size it rounds to,
half up (10.5 as 11), as Excel for Mac does; MS Sans Serif measures as Times New Roman
(`font_substitutes()`). Another font raises `LookupError`. An italic face's slant reaches past
its glyphs, so a cell in italic holds a point or two less than an upright one, as in Excel.

Excel sizes a column by the digit width of the workbook's Normal font and measures a cell's text
in the cell's own font. When the two differ, pass the Normal font as `normal_font`; it defaults
to `font`. A reader of a workbook passes the cell's font and the workbook's Normal font.

```python
>>> from xlformat import Font
>>> format_value(88888888, "0", column_width=8, font=("Calibri", 14), platform="mac")
'88888888'
>>> format_value(88888888, "0", column_width=8, font=("Calibri", 14), normal_font=Font("Calibri", 11), platform="mac")
'######'
```

The first column is 8 units of a 7-point digit (61 points), the second 8 units of a 6-point
digit (53 points), in which eight 7-point digits do not fit.

## Columns read from a file

A file stores a column's width as `<col width>`, which is not the dialog number, and a column
without a `<col>` entry takes its width from `<sheetFormatPr>`. `file_column_metrics` converts
either into the column Excel lays out; its `width` is what `column_width` takes.

```python
>>> from xlformat import file_column_metrics
>>> file_column_metrics(platform="mac")
ColumnMetrics(width=8.0, points=53.0, budget=48.0, left_inset=3.0, right_inset=2.0, digit_width=6.0)
>>> file_column_metrics(8.43, platform="mac").width
7.666666666666667
>>> format_value(12345678, "0", column_width=file_column_metrics(8.43, platform="mac").width, platform="mac")
'#########'
>>> format_value(12345678, "0", column_width=8.43, platform="mac")
'12345678'
```

A column openpyxl leaves without a width is 8 units in Excel for Mac (53 points at Calibri 11),
and one it writes as `8.43` is 7.67 units (51 points), narrower than the 8.43 of the dialog,
which is 56 points. `file_column_metrics` takes `base_col_width` and `default_col_width` from
`<sheetFormatPr>` and `normal_font`, the workbook's Normal font (`font` is its deprecated name
there). `column_metrics` gives the same record for a dialog width. `budget` is what a number
may take, 0 in a column under 5 points; a column of no width has a `width` of 0, which
`column_width` refuses with `ValueError`. `left_inset` is where left-aligned text in the Normal
font starts, from the cell's left boundary (its gridline and `ceil(digit_width / 4)` screen units
of clear space), and `right_inset` where right-aligned text ends, before the right boundary (the
clear space); a cell in another font has its own, from `text_insets`:

```python
>>> from xlformat import text_insets
>>> text_insets(("Calibri", 24), platform="mac").left
4.0
```

A merged cell formats as one column of the sum of its columns' points, with a single padding, and
a hidden column counts as 0. `column_metrics_from_points` makes that column:

```python
>>> from xlformat import column_metrics_from_points
>>> points = sum(file_column_metrics(w, platform="mac").points for w in (2, 2, 2))
>>> points
36.0
>>> merged = column_metrics_from_points(points, platform="mac")
>>> merged.width, merged.budget
(5.166666666666667, 31.0)
>>> format_value(12345, "0", column_width=merged.width, platform="mac")
'12345'
>>> format_value(12345678, "0", column_width=merged.width, platform="mac")
'######'
```

How Excel measures, as reproduced here:

- A column `w` units wide is `round(MDW · w + P)` points, MDW being the width of a digit in the
  Normal font: 5 points at 9 and 10, 6 at 11 and 12, 7 at 14, 8 at 16, 9 at 18, 10 at 20 and 12
  at 24, in Calibri and Aptos Narrow alike (Arial: 5, 6, 6, 7, 8, 9, 10, 11 and 13). `P` is the
  padding, `2 · ceil(MDW / 4) + 1`: 5 points for a digit width up to 8, 7 for 9 to 12, 9 for 13. The text gets `MDW · w`. Below one unit a column is
  `round(w · (MDW + P))` points.
- A `<col width>` of `w` is `trunc((256 · w + trunc(128 / MDW)) / 256 · MDW)` points, the formula
  of ECMA-376 with the digit width in points and `w` first truncated to 1/256.
- A text measures the sum of its glyphs' advances, each rounded half up to whole points from the
  face's units per em (`u · S / E`; on Windows whole pixels at the size's ppem), so a Calibri
  hyphen is 3 points at 11 and 4 at 12 while a digit is 6 at both.
- A number fits when its measure is at most the budget. Otherwise it becomes as many `#` as fit,
  none in a column too narrow for one (`General` still shows one) and at most 255.
- A `*` fill repeats as many times as fit in what is left. A number or text that would exceed 255
  characters with its fill shows a `#` run instead.
- `General` is the exception: Excel reduces it by characters against the width in units, one per
  character and a leading minus by its own width (half a digit at Calibri 11), not by glyph
  widths. When no shortened form fits, the shortest one shows if its glyphs fit.

## Rows

Excel fits a row without a stored height to what it holds when it opens a file. `row_metrics`
gives that height for the workbook's Normal font, the fonts of the row's other cells and the
cells with wrapped text, each as `(font, lines)` with `None` for the Normal font:

```python
>>> from xlformat import row_metrics
>>> row_metrics(platform="mac")
RowMetrics(height=15.0, font_height=15.0, line_height=16.0, accounting_height=0.0)
>>> row_metrics(cell_fonts=[("Calibri", 14)], platform="mac").height
19.0
>>> row_metrics(wrapped=[(None, 3)], platform="mac").height
48.0
>>> row_metrics(("Calibri", 14), wrapped=[(("Calibri", 9), 2)], platform="mac").height
28.0
>>> row_metrics(("Times New Roman", 12.5), platform="mac").height  # a half size takes the next whole one
17.0
```

A row holds the largest ascent and the largest descent of its fonts, and bold or italic changes
nothing. Faces of one family split their heights alike, so such a row is as tall as its tallest
font; Arial's descent is smaller, so Arial 11 beside Calibri 10, both 14 points, makes 15. Each wrapped line takes
its font's height plus one point, so a single wrapped line makes a Calibri 11 row 16 points; when
the Normal font or another cell's font in the row is larger, a smaller font's wrapped text takes up
to a few points more. The caller breaks the lines. Calibri, Aptos Narrow, Arial, Times New Roman,
Courier New and Georgia are measured at every size from 1 to 72 points, half sizes included; any
other face or size raises `LookupError`. Several cells with wrapped text in
one row are a combination not yet captured: the row takes the tallest of them.

Text turned to an angle is not fitted to the column: Excel shows the width-less rendering, with its
`_x` padding, without fills and never as `#`, so a caller formats angled cells without a width.
The capture turns text 15, 30, 45, 60, 75 and 90 degrees and -45 and -90.

`vertical_metrics` gives where Excel draws a face's text within the row, in sheet points, as
the screen shows it:

```python
>>> from xlformat import ink_metrics, vertical_metrics
>>> m = vertical_metrics(("Calibri", 11), normal_font=("Calibri", 11), platform="mac")
>>> m.baseline_from_top, m.baseline_from_bottom
(12.0, 3.0)
>>> ink = ink_metrics(("Calibri", 11), platform="mac")  # measured: the ink of "H" above and "g" below
>>> ink.cap_height, ink.descender
(7.0, 2.0)
>>> vertical_metrics(("Calibri", 9), platform="mac").baseline_from_bottom  # its own: a Calibri 9 Normal font
2.0
>>> vertical_metrics(("Calibri", 9), normal_font=("Calibri", 11), platform="mac").baseline_from_bottom
3.0
>>> m.line_pitch, m.line_height
(15.0, 16.0)
```

Bottom-aligned text, Excel's default, has its baseline `baseline_from_bottom` above the row's
bottom; top-aligned text has it `baseline_from_top` below the row's top (a row runs from the top
of the gridline above it, which the row's own fill covers, to the top of the next); centered text
sits midway
between the two. Wrapped lines are `line_pitch` apart, the last placed as bottom-aligned text. Bold
and italic place as the regular face. The workbook's Normal font moves text in two ways.
Bottom-aligned text sits at least the Normal font's `baseline_from_bottom` up (Calibri 9 and Arial
9 to 11 have 2 points of their own and sit 3 up in a Calibri 11 workbook); pass the workbook's
Normal font as `normal_font` and `baseline_from_bottom` includes this. Wrapped text in a face with
a smaller `line_inset` than the row's other unwrapped fonts sits that difference higher, the extra
height `row_metrics` gives the row lying below the text; that one is the caller's to apply, since
it depends on every font in the row. The fields are the
distances Excel shows; the ascent and the descent are not split from the cell's insets.

`rule_metrics` gives where Excel draws a face's underlines and strikethrough, in sheet points from
the baseline, at a zoom (positions move in steps of a screen point). A double underline's first
line moves up until the double ends above the accounting underline, so in Calibri 11 it sits on
the baseline:

```python
>>> from xlformat import rule_metrics
>>> r = rule_metrics(("Calibri", 11), platform="mac")
>>> r.underline_top, r.underline_thickness, r.accounting_top, r.strikethrough_top
(1.0, 1.0, 3.0, 3.0)
>>> r.double_top, r.double_thickness, r.double_gap
(0.0, 1.0, 2.0)
>>> rule_metrics(("Calibri", 36), zoom=200, platform="mac").underline_thickness
2.0
```

## Measuring text yourself

The measure behind Excel mode is public, for layout that has to agree with it: spilling text
into the next column, wrapping, indents.

```python
>>> from xlformat import autofit_width, digit_width, glyph_widths, text_width
>>> text_width("1,000.0", platform="mac")
36.0
>>> text_width("Hello World", ("Calibri", 14), platform="mac")
66.0
>>> glyph_widths("W i", Font("Calibri", 11, italic=True), platform="mac")
(10.0, 2.0, 3.0)
>>> digit_width(("Aptos Narrow", 16), platform="mac")
8.0
>>> autofit_width("Hello World", platform="mac")
62.0
>>> autofit_width("0" * 20, platform="mac"), autofit_width("0" * 20, platform="windows")
(129.0, 112.5)
```

`text_width` is the sum of the whole advances, the number every fit decision uses;
`autofit_width` is the column width Excel's autofit gives one line of text: its advances (on
the Mac a text holding a non-ASCII glyph summed unrounded), the padding, and a long-text bonus
of a point from 15 widths of the Normal font's digit and one more every 5 (on Windows a pixel
from 12 and one more every 4). Two captured Mac ranges disagree with that rule, and
`autofit_width` still answers there by it: from 1 to 5 points Excel's autofit widths are all
multiples of 5 points (172 of 1,080 captured texts differ from the rule), and at 383, 397 and
409 points 3,170 of 76,464 differ, mostly by 1 or 2 points;
`glyph_widths` gives the whole advance of each character, for measuring runs of a text;
`digit_width` is the width of one column unit (and one indent level) for a Normal font;
`overhang` is how far an italic face reaches past its advances on each side; `measured_fonts`
lists the fonts measured, and `font_substitutes` the names measured as another face.

## Excel for Windows

Excel for Windows, the default (`platform="windows"` or `None`), is measured as Excel
16.0.20527.20196 (Microsoft 365, en-US, 96 dpi) shows it: the same rules in whole pixels, with glyph widths hinted for the screen, so a digit of
Calibri 11 is 7 pixels (5.25 points) and a column is a whole number of pixels. Results stay in points. A column without a
width is rounded up to a multiple of 8 pixels, and a fill stops at 255 characters where the Mac
shows a `#` run. Text is measured in 25 families (`measured_fonts(platform="windows")`: the Mac's
but Courier and Helvetica) in four styles at every ppem from 11 to 32, sizes 8 to 24 and the half
sizes between; another size scales the nearest measured ppem and may differ from Excel by a
character near the edge of a column. Row heights and text placement (`row_metrics`,
`vertical_metrics`) cover Calibri, Aptos Narrow, Arial, Times New Roman and Courier New at every
size from 1 to 72 points, half sizes included.

```python
>>> from xlformat import digit_width
>>> digit_width(("Calibri", 11), platform="windows")
5.25
>>> file_column_metrics(normal_font=("Calibri", 11), platform="windows").points
48.0
>>> format_value(1234567.891, "#,##0.00", column_width=8)
'########'
>>> format_value(1234567.891, "#,##0.00", column_width=8, platform="mac")
'#########'
>>> row_metrics(("Calibri", 11), platform="windows").height
15.0
```

Text placement (`vertical_metrics`) is measured on Windows too, in whole pixels, in the five
faces `row_metrics` covers; ink, underline positions and insets (`ink_metrics`, `rule_metrics`,
`text_insets`) in Calibri, Aptos Narrow and Arial:

```python
>>> from xlformat import rule_metrics, vertical_metrics
>>> vertical_metrics(("Calibri", 11), platform="windows").baseline_from_bottom
3.0
>>> rule_metrics(("Calibri", 11), platform="windows").underline_top
1.5
```

## Custom measures

`measure` replaces the character count with any callable from text to a width, and `width` is
then in the callable's units. It suits a proportional font measured some other way.

```python
>>> format_value(1234.5, "0.00*-", width=20, measure=lambda text: 2 * len(text))
'1234.50---'
```
