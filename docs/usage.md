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
them. A reader of files usually catches the error and falls back to `General`, which is what
Excel's repair does to the cell.

```python
>>> from xlformat import FormatError
>>> try:
...     format_value(1, "0;0;0;0;0")
... except FormatError:
...     print("refused")
refused
```
