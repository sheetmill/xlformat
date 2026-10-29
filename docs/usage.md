# Usage

## Formatting a value

{func}`xlformat.format_value` takes a value and a format string and returns the text Excel
displays.

```python
>>> from xlformat import format_value
>>> format_value(-1234.5, "#,##0.00;(#,##0.00)")
'(1,234.50)'
>>> format_value(44927.75, "yyyy-mm-dd hh:mm")
'2023-01-01 18:00'
>>> format_value(0.125, "# ?/?")
' 1/8'
>>> format_value("abc", '0.00;-0.00;0;"["@"]"')
'[abc]'
>>> format_value(True, "General")
'TRUE'
```

Accepted values:

| value | treated as |
|---|---|
| `int`, `float`, `Decimal` | a number; floats are used exactly, as the double Excel stores |
| `bool` | `TRUE` or `FALSE`, through the text section if there is one |
| `str` | text, through the text section if there is one |
| `None` | an empty cell, shown as `''` |
| `date`, `datetime` | a date serial in the chosen date system |
| `time`, `timedelta` | a fraction of a day, for time and elapsed formats |

NaN and infinities raise `ValueError`; other types raise `TypeError`. An empty format string
means `General`.

## Dates and the 1904 system

Dates are serial numbers in Excel. Pass `date1904=True` for a workbook that uses the 1904 date
system, which shows the same serial as a date four years later and renders negative values
with a minus sign instead of `#`.

```python
>>> format_value(0, "yyyy-mm-dd", date1904=True)
'1904-01-01'
>>> format_value(-1, "m/d/yy")
'###############################################################################################################################################################################################################################################################'
```

The last line is 255 `#`, Excel's overflow without a width; with a width it fills the column
(see {doc}`width`).

## What Excel shows around the text

{func}`xlformat.format_cell` returns a {class}`xlformat.FormattedCell`: the text, the section
color, the default alignment, the position of a `*` fill, whether the value overflowed, and
`full_text`, the rendering without a width.

```python
>>> from xlformat import format_cell, DEFAULT_PALETTE
>>> cell = format_cell(-1234.5, "[Red]0.00;[Blue]-0.00")
>>> cell.text, cell.color, cell.align
('-1234.50', 'Blue', 'right')
>>> DEFAULT_PALETTE[cell.color]
'0000FF'
```

Alignment follows the value: numbers and dates right, text left, booleans centered. A number
under a format with only a text section, such as `@`, is left-aligned like text, as in Excel.

## Error values

A cell holding an error shows the error's text whatever its number format, centered and in the
font's color, and a column too narrow for it shows `#` as for a number. Pass the error as an
`ExcelError`:

```python
>>> from xlformat import ExcelError
>>> format_value(ExcelError.DIV0, "0.00")
'#DIV/0!'
>>> format_value(ExcelError("#N/A"), ";;;")
'#N/A'
>>> format_value(ExcelError.DIV0, "General", column_width=5)
'######'
>>> cell = format_cell(ExcelError.DIV0, "[Red]0")
>>> cell.align, cell.color
('center', None)
```

`ExcelError` lists every error current Excel documents, from `#NULL!` to `#SPILL!` and
`#TIMEOUT!`. `ExcelError(text)` also accepts a text shaped like an error code that it does not
list, so a code from a newer Excel renders like the others; any other text raises `ValueError`.
A reader of a file passes `ExcelError(cell.value)` for a cell of type `e`. Compare errors with
`==`, not `is`: an unlisted code is a new object on each lookup, though equal, with an equal
hash, to every other lookup of the same text.

## Compiled formats

{meth}`xlformat.NumberFormat.parse` compiles a format once. It is cached, so calling it again
with the same string is cheap, and the compiled object can be reused across threads.

```python
>>> from xlformat import NumberFormat
>>> fmt = NumberFormat.parse("yyyy-mm-dd")
>>> fmt.is_date
True
>>> fmt.format(44927)
'2023-01-01'
```

## Invalid formats

{class}`xlformat.FormatError` is raised only for format strings Excel refuses to open from a
file, such as `0.00e+00` with a lower-case exponent or five sections. Strings that the Format
Cells dialog rejects but a file may contain, such as two `*` fills, render as Excel renders
them. A reader of files catches the error and renders the cell as `General`, which is what
Excel shows after it repairs the file: the repair resets each cell with a refused format to
`General` and keeps every other cell's format.

```python
>>> from xlformat import FormatError
>>> try:
...     format_value(1, "0;0;0;0;0")
... except FormatError:
...     print("refused")
refused
```
