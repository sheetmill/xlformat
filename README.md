# xlformat

xlformat applies an Excel number format to a value and returns the text Excel displays. It is
pure Python with no runtime dependencies, and its output is tested against text captured from
Excel 16 for Mac and, with `platform="windows"`, from Excel for Windows (16.0.20527.20196,
Microsoft 365, en-US, 96 dpi).

```python
>>> from xlformat import format_value
>>> format_value(1234.5, "#,##0.00")
'1,234.50'
```

Documentation: <https://xlformat.readthedocs.io>.

## Installation

```bash
pip install xlformat
```

Python 3.11 or later.

## Usage

`format_value` returns the displayed text as a string.

```python
>>> format_value(-1234.5, "#,##0.00;(#,##0.00)")
'(1,234.50)'
>>> format_value(44927.75, "yyyy-mm-dd hh:mm")
'2023-01-01 18:00'
>>> format_value(0.125, "# ?/?")
' 1/8'
>>> import datetime
>>> format_value(datetime.date(2023, 1, 1), "dddd, mmmm d, yyyy")
'Sunday, January 1, 2023'
```

Values may be `int`, `float`, `Decimal`, `bool`, `str`, `None`, `date`, `datetime`, `time` or
`timedelta`. Pass `date1904=True` for workbooks that use the 1904 date system.

`format_cell` takes the same arguments and returns a `FormattedCell` with what Excel shows
around the text: the section color, the default alignment, the position of a `*` fill, whether
the value overflowed, and the text without any width applied.

```python
>>> from xlformat import format_cell
>>> cell = format_cell(-1234.5, "[Red]0.00;[Blue]-0.00")
>>> cell.text, cell.color, cell.align
('-1234.50', 'Blue', 'right')
```

`xlformat.DEFAULT_PALETTE` maps color names such as `Blue` or `Color12` to `RRGGBB`.

`NumberFormat.parse` compiles a format string once. Parsing is cached, and the compiled object
has `format` and `format_cell` methods and facts about the pattern such as `is_date`. An
invalid format string raises `FormatError`, but only when Excel itself would refuse to open a
file containing it. Callers reading files catch it and render the cell as `General`: Excel's
repair resets each cell with such a format to `General` and keeps every other cell's format.

### Width

Without a width, the result is what copy and paste from Excel gives: `*` fills add nothing,
`_x` pads add one space, and a value that cannot be shown is 255 `#` characters. Two parameters
add a width, and they cannot be combined.

- **`width`** is a number of characters. It suits terminals and plain text. `*` fills repeat
  to the width, a number or date that does not fit becomes `width` `#` characters, `General`
  drops digits before it overflows, and text never overflows.
- **`column_width`** is the number from Excel's Column Width dialog, such as 8.43. The text is
  measured with the glyph widths of `font`, so the result matches the sheet character for
  character, including how many `#` and fill characters fit. `font` is a name or a
  `(name, size)` pair or a `Font`, and defaults to `("Calibri", 11)`, the font openpyxl writes.
  When the cell's font differs from the workbook's Normal font, `font` is the cell's and
  `normal_font` the Normal one, which sets the column's units. A width read from a file
  (`<col width>`, or none) is not the dialog number: `file_column_metrics` converts it.

```python
>>> format_value(1234567.891, "General", width=7)
'1234568'
>>> format_value(1000, "#,##0.0", width=6)
'######'
>>> format_value(1000, "#,##0.0", column_width=6)
'1,000.0'
>>> format_value(1234.5, "0.00*-", column_width=12, font=("Calibri", 14))
'1234.50---------'
```

The two modes differ on purpose. In Excel a comma or a point is narrower than a digit, so
`1,000.0` fits a column 6 wide, while in character mode it is seven characters. A `measure`
callable can replace the character count when neither mode fits, and `width` is then in its
units.

### Quirks

Three Excel behaviors are reproduced by default and can be switched off with
`quirks=Quirks.STRICT` or by passing `Quirks` with individual fields set. [QUIRKS.md](https://github.com/sheetmill/xlformat/blob/main/QUIRKS.md) has the
evidence for each.

| quirk | Excel | strict |
|---|---|---|
| `leap_year_1900` | serial 60 under `mmmm d, yyyy` is `February 29, 1900` | `February 28, 1900` |
| `fifteen_digit_precision` | 0.1234512345123456 under sixteen decimals is `0.1234512345123460` | `0.1234512345123456` |
| `fraction_denominators` | 0.91 under `# ?/?` is `1` | ` 8/9`, the closest fraction |

## Scope and known differences

- **Excel mode measures 27 families on the Mac** (`measured_fonts()`: Calibri, Aptos Narrow, Arial,
  Times New Roman, Verdana, Tahoma, Courier New and 20 more) in four styles at any size, and the
  same families but Courier and Helvetica on Windows (`platform="windows"`, 96 dpi); any other
  font raises `LookupError`. The [Width page](https://xlformat.readthedocs.io/en/stable/width.html)
  gives the sizes captured on each platform and how text is measured.
- **Character mode is its own contract.** A `#` run is always `width` characters, and a value
  is judged by its character count, where Excel would fit or overflow some texts by glyph
  width.
- **Expected outputs come from Excel 16 for Mac and Excel for Windows** under Microsoft 365,
  with the en-US locale. Month and day names are English; locale tags such as `[$-409]` are
  parsed but not applied. The default platform is the Mac; on Windows, where text, columns and
  rows are measured in pixels, pass `platform="windows"`. Text placement and underline positions
  (`vertical_metrics`, `ink_metrics`, `rule_metrics`) are measured on both, on Windows for
  Calibri, Aptos Narrow and Arial.
- **Error values** are passed as `ExcelError` (`ExcelError("#N/A")`); a plain string such as
  `"#N/A"` is text.
- **Conditional formatting** and cell styles beyond the format string are out of scope.

Every captured row passes. Two known deviations are listed in [QUIRKS.md](https://github.com/sheetmill/xlformat/blob/main/QUIRKS.md): a Buddhist-year code
shows the Buddhist year without a width, as the sheet does, where Excel's clipboard gives the plain
year; and an elapsed code after a date code (`d [s].00`) is not modeled.

Seventeen color readings in the Excel captures are skipped because Excel's scripting interface
reports automatic black for them. [QUIRKS.md](https://github.com/sheetmill/xlformat/blob/main/QUIRKS.md)
describes them and says which have been checked on screen.

## Security

To report a security problem, see
[SECURITY.md](https://github.com/sheetmill/xlformat/blob/main/SECURITY.md).

## Related packages

xlformat is one of three packages from the [sheetmill](https://github.com/sheetmill/sheetmill)
project, with [sheetpeek](https://github.com/sheetmill/sheetpeek), which renders workbooks to HTML,
Markdown and terminal output, and [xlrecalc](https://github.com/sheetmill/xlrecalc), which evaluates
formulas. xlformat has no dependencies; the other two require it. The `sheetmill` meta-package
installs all three.

## Development

With uv, which installs the exact versions pinned in `uv.lock`:

```bash
uv sync --group docs
uv run pytest                 # unit tests and every captured Excel row
uv run pytest --tier=fast     # a per-format sample of the Excel captures
uv run ruff check . && uv run ruff format --check . && uv run mypy
uv run --group docs sphinx-build -n -W docs docs/_build/html
npx --yes pagefind@1.5.2 --site docs/_build/html   # search index for served copies
```

With pip 25.1 or later, which ignores `uv.lock` and installs the latest versions the
dependency groups allow:

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e . --group dev --group docs
pytest && ruff check . && ruff format --check . && mypy
sphinx-build -n -W docs docs/_build/html
```

[CONTRIBUTING.md](https://github.com/sheetmill/xlformat/blob/main/CONTRIBUTING.md) lists the checks a change must pass and how commits are written.

The expected outputs under `tests/captures/` were captured from Excel for Mac by the scripts in
`capture/`, and those under `tests/captures/windows/` from Excel for Windows by capture jobs whose
manifests are committed beside them. On the Mac each suite is written as a workbook with
openpyxl, opened in Excel through AppleScript, and read back as displayed text at each column
width. Regenerating a Mac suite needs Excel for Mac, and a Windows suite runs its job again; both
are described in [CAPTURING.md](https://github.com/sheetmill/xlformat/blob/main/CAPTURING.md):

```bash
uv run python -m capture numeric_core
```

Most of the code is written with Claude Code, reviewed before each commit and tested against
output captured from Excel; see [CONTRIBUTING.md](https://github.com/sheetmill/xlformat/blob/main/CONTRIBUTING.md).

## License

MIT.
