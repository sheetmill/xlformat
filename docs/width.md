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
glyph widths of `font` in points, as Excel for Mac measures it, so the result matches the sheet
character for character, including the number of `#` and fill characters.

```python
>>> format_value(1000, "#,##0.0", width=6)
'######'
>>> format_value(1000, "#,##0.0", column_width=6)
'1,000.0'
>>> format_value(1234.5, "0.00*-", column_width=12)
'1234.50-----------'
>>> format_value(1234.5, "0.00*-", column_width=12, font=("Calibri", 14))
'1234.50---------'
```

`font` is the cell's font: a name, a `(name, size)` pair or a `Font`; the default is
`("Calibri", 11)`, the Normal style openpyxl writes. Calibri is measured at 9, 10, 11, 12, 14 and
16 points; other sizes scale the nearest measured table and may differ from Excel by a character
at the edge of a column. A font or style without a table raises `LookupError`.

Excel sizes a column by the digit width of the workbook's Normal font and measures a cell's text
in the cell's own font. When the two differ, pass the Normal font as `normal_font`; it defaults
to `font`. A reader of a workbook passes the cell's font and the workbook's Normal font.

```python
>>> from xlformat import Font
>>> format_value(88888888, "0", column_width=8, font=("Calibri", 14))
'88888888'
>>> format_value(88888888, "0", column_width=8, font=("Calibri", 14), normal_font=Font("Calibri", 11))
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
>>> file_column_metrics()
ColumnMetrics(width=8.0, points=53.0, budget=48.0, left_inset=3.0, right_inset=2.0, digit_width=6.0)
>>> file_column_metrics(8.43).width
7.666666666666667
>>> format_value(12345678, "0", column_width=file_column_metrics(8.43).width)
'#########'
>>> format_value(12345678, "0", column_width=8.43)
'12345678'
```

A column openpyxl leaves without a width is 8 units in Excel for Mac (53 points at Calibri 11),
and one it writes as `8.43` is 7.67 units (51 points), narrower than the 8.43 of the dialog,
which is 56 points. `file_column_metrics` takes `base_col_width` and `default_col_width` from
`<sheetFormatPr>` and `font`, the workbook's Normal font. `column_metrics` gives the same
record for a dialog width. `budget` is what a number may take; `left_inset` and `right_inset`
are where left-aligned text starts and right-aligned text ends, provisional values from Excel's
print layout until a screen capture confirms them.

How Excel measures, as reproduced here:

- A column `w` units wide is `round(MDW · w + 5)` points, MDW being the width of a digit in the
  Normal font: 5 points for Calibri 9 and 10, 6 for 11 and 12, 7 for 14, 8 for 16. Five of those
  points are padding; the text gets `MDW · w`. Below one unit a column is `round(w · (MDW + 5))`
  points.
- A `<col width>` of `w` is `trunc((256 · w + trunc(128 / MDW)) / 256 · MDW)` points, the formula
  of ECMA-376 with the digit width in points and `w` first truncated to 1/256.
- A text measures the sum of its glyph advances, rounded down. Advances are hinted per size, so
  a hyphen is 3 points at 11 and 4 at 12 while a digit is 6 at both.
- A number fits when its measure is at most the budget. Otherwise it becomes as many `#` as fit,
  none in a column too narrow for one (`General` still shows one) and at most 255.
- A `*` fill repeats as many times as fit in what is left. A number or text that would exceed 255
  characters with its fill shows a `#` run instead.
- `General` is the exception: Excel reduces it by characters against the width in units, one per
  character and a leading minus by its own width (half a digit at Calibri 11), not by glyph
  widths. When no shortened form fits, the shortest one shows if its glyphs fit.

## Custom measures

`measure` replaces the character count with any callable from text to a width, and `width` is
then in the callable's units. It suits a proportional font measured some other way.

```python
>>> format_value(1234.5, "0.00*-", width=20, measure=lambda text: 2 * len(text))
'1234.50---'
```
