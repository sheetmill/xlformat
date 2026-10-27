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

`font` is a name or a `(name, size)` pair; the default is `("Calibri", 11)`, the Normal style
openpyxl writes. A reader of a workbook passes the workbook's Normal font. Calibri is measured
at 9, 10, 11, 12, 14 and 16 points; other sizes scale the nearest measured table and may differ
from Excel by a character at the edge of a column. Another font raises `LookupError`.

How Excel measures, as reproduced here:

- A column `w` units wide is `round(MDW · w + 5)` points, MDW being the width of a digit in the
  Normal font: 5 points for Calibri 9 and 10, 6 for 11 and 12, 7 for 14, 8 for 16. Five of those
  points are padding; the text gets `MDW · w`.
- A text measures the sum of its glyph advances, rounded down. Advances are hinted per size, so
  a hyphen is 3 points at 11 and 4 at 12 while a digit is 6 at both.
- A number fits when its measure is at most the budget. Otherwise it becomes as many `#` as fit.
- A `*` fill repeats as many times as fit in what is left.
- `General` is the exception: Excel reduces it by characters against the width in units, one per
  character and half for a leading minus, not by glyph widths.

## Custom measures

`measure` replaces the character count with any callable from text to a width, and `width` is
then in the callable's units. It suits a proportional font measured some other way.

```python
>>> format_value(1234.5, "0.00*-", width=20, measure=lambda text: 2 * len(text))
'1234.50---'
```
