# xlformat

xlformat applies an Excel number format to a value and returns the text Excel displays. It is
pure Python with no runtime dependencies, and its output is tested against text captured from
Excel 16 for Mac.

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
`quirks=Quirks.STRICT` or by passing `Quirks` with individual fields set. `QUIRKS.md` has the
evidence for each.

| quirk | Excel | strict |
|---|---|---|
| `leap_year_1900` | serial 60 under `mmmm d, yyyy` is `February 29, 1900` | `February 28, 1900` |
| `fifteen_digit_precision` | 0.1234512345123456 under sixteen decimals is `0.1234512345123460` | `0.1234512345123456` |
| `fraction_denominators` | 0.91 under `# ?/?` is `1` | ` 8/9`, the closest fraction |

## Scope and known differences

- **Excel mode measures Calibri and Aptos Narrow**, regular, bold, italic and bold italic, at
  9, 10, 11, 12, 14 and 16 points, in Excel for Mac (which bundles Aptos Narrow, the Windows
  default). Other sizes are scaled from the nearest measured size and may differ from Excel by
  a character near the edge of a column. Any other font raises `LookupError`.
- **Character mode is its own contract.** A `#` run is always `width` characters, and a value
  is judged by its character count, where Excel would fit or overflow some texts by glyph
  width.
- **Expected outputs come from Excel 16 for Mac** with the en-US locale. Month and day names are English;
  locale tags such as `[$-409]` are parsed but not applied. Windows Excel renders text with a
  different engine and may disagree on a borderline fit in Excel mode.
- **Error values** are passed as `ExcelError` (`ExcelError("#N/A")`); a plain string such as
  `"#N/A"` is text.
- **Conditional formatting** and cell styles beyond the format string are out of scope.

Every captured row passes; no known deviations remain. Fifteen color readings in the Excel captures
are skipped because Excel's scripting interface misreports them; the sheet itself shows the
color the package returns.

## Related packages

xlformat is one of three packages from the `sheetmill` project, with `sheetpeek`, which
renders workbooks to HTML, Markdown and terminal output, and `xlformula`, which evaluates
formulas. Each installs and works on its own. A `sheetmill` meta-package that installs all
three is planned.

## Development

With uv, which installs the exact versions pinned in `uv.lock`:

```bash
uv sync --group docs
uv run pytest                 # unit tests and every captured Excel row, about eight seconds
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

`CONTRIBUTING.md` lists the checks a change must pass and how commits are written.

The documentation's search page uses Pagefind when the site is served over http or https,
as on Read the Docs or with `python -m http.server -d docs/_build/html`. A copy opened
straight from disk uses Sphinx's built-in search instead, because browsers do not let a
`file://` page fetch Pagefind's index.

The expected outputs under `tests/captures/` were captured from Excel for Mac by the scripts in
`capture/`: each suite is written as a workbook with openpyxl, opened in Excel through
AppleScript, and read back as displayed text at each column width. Regenerating a suite needs
Excel for Mac and is described in `CAPTURING.md`:

```bash
uv run python -m capture numeric_core
```

Most of the code is written with Claude Code, reviewed before each commit and tested against
output captured from Excel; see `CONTRIBUTING.md`.

## License

MIT.
