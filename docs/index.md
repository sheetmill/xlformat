# xlformat

xlformat applies an Excel number format to a value and returns the text Excel displays. It is
pure Python with no runtime dependencies, and its output is tested against text captured from
Excel 16 for Mac.

```python
>>> from xlformat import format_value
>>> format_value(1234.5, "#,##0.00")
'1,234.50'
>>> format_value(1000, "#,##0.0", column_width=6)
'1,000.0'
```

```bash
pip install xlformat
```

Python 3.11 or later. The source is at <https://github.com/sheetmill/xlformat>.

xlformat is one of three packages from the `sheetmill` project, with `sheetpeek`, which
renders workbooks, and `xlformula`, which evaluates formulas. Each installs and works on its
own. A `sheetmill` meta-package that installs all three is planned.

The search box uses Pagefind on the published site and on any copy served over http. A copy
of these pages opened straight from disk uses Sphinx's built-in search instead, because
browsers do not let a `file://` page fetch Pagefind's index.

```{toctree}
:maxdepth: 2

usage
width
quirks
testing
api
changelog
```
