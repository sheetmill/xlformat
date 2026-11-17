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

## Built-in formats

A cell can name its format by a built-in id instead of a string: Excel writes id 14 for every date
typed as a short date. `builtin_format` gives the string Excel uses for an id, as Excel for Mac
reports it in en-US:

```python
>>> from xlformat import builtin_format
>>> builtin_format(14)
'm/d/yy'
>>> format_value(45292, builtin_format(14))
'1/1/24'
>>> builtin_format(47)
'mm:ss.0'
```

Ids 0 to 163 are built in (`BUILTIN_FORMATS` holds them all; many are `General`); 164 and above
are defined in the file and raise `ValueError`. What some ids show depends on the viewer's system:
Excel for Mac follows a custom short date set in macOS for ids 14 and 22 (a year-first setting
shows `2024-01-01`) and pads the day of 15 and 16 to match, while the region alone changes
nothing on the Mac. For Excel on Windows, Microsoft's [Format numbers as dates or
times](https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309) says that "Date and time formats that begin with an asterisk (\*) respond to
changes in regional date and time settings that are specified in Control Panel." The table is the
en-US one.

## Sheets that hide zeros

A sheet can hide zero values (Excel's "Show a zero in cells that have zero value" turned off,
`<sheetView showZeros="0">` in the file). Pass `show_zeros=False` for such a sheet: a number equal
to 0 then shows nothing, unless the format has a zero section of its own. A value that only
rounds to 0 is not hidden.

```python
>>> format_value(0, "0.00", show_zeros=False)
''
>>> format_value(0, '#,##0;(#,##0);"-"', show_zeros=False)
'-'
>>> format_value(0.004, "0.00", show_zeros=False)
'0.00'
```

Pass the value as stored: an exact zero, `-0.0` included, is a zero. Excel loads some file texts
as 0 that Python reads otherwise (a denormal such as `5e-324`), and calculates some formulas as 0
(`=0.1+0.2-0.3`); converting those is the reader's or the calculator's part.

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

A workbook can replace Excel's palette (`<indexedColors>`, 64 entries). `section_rgb` gives the
color Excel draws with it: `[ColorN]` takes the palette's entry N + 7, while the eight names keep
their standard colors whatever the palette holds:

```python
>>> from xlformat import section_rgb
>>> palette = ["FF000000"] * 64
>>> palette[12] = "FF123456"
>>> section_rgb("Color5", palette), section_rgb("Red", palette), section_rgb("Color5")
('123456', 'FF0000', '0000FF')
```

Entries may be `RRGGBB` or `AARRGGBB`. For an index a shorter palette lacks, the default color is
used, a choice of this package: what Excel does with a short palette is not captured.

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

## Excel's TEXT function

A formula engine evaluating `TEXT` calls `text_function`. It is the width-less text, with what
`TEXT` does differently: a blank (`None`) is 0, an error passes through, and a format Excel refuses
or a value the format cannot show gives `ExcelError.VALUE` instead of raising or showing `#`.

```python
>>> from xlformat import text_function
>>> text_function(1234.5, "*-0")
'1235'
>>> text_function(1234.5, "0_)")
'1235 '
>>> text_function(None, "0.00")
'0.00'
>>> text_function(-1, "m/d/yy") == ExcelError.VALUE
True
>>> text_function(1, "0.00e+00") == ExcelError.VALUE
True
>>> text_function(-5, "")
'-'
```

Text that reads as a number stays text: a caller evaluating `TEXT("12", "0.00")` converts it first.

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
