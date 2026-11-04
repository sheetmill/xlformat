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

xlformat is one of three packages from the [sheetmill](https://github.com/sheetmill/sheetmill)
project, with [sheetpeek](https://github.com/sheetmill/sheetpeek), which renders workbooks, and
[xlrecalc](https://github.com/sheetmill/xlrecalc), which evaluates formulas. xlformat has no
dependencies; the other two require it. The `sheetmill` meta-package installs all three.

```{toctree}
:maxdepth: 2

usage
width
quirks
testing
api
changelog
```
