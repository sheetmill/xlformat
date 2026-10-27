# Quirks

Three Excel behaviors have a switch in {class}`xlformat.Quirks`. They are on by default
(`Quirks.EXCEL`); `Quirks.STRICT` turns all three off, and individual fields can be set.

```python
>>> from xlformat import Quirks, format_value
>>> format_value(60, "mmmm d, yyyy")
'February 29, 1900'
>>> format_value(60, "mmmm d, yyyy", quirks=Quirks.STRICT)
'February 28, 1900'
>>> format_value(60, "mmmm d, yyyy", quirks=Quirks(leap_year_1900=False))
'February 28, 1900'
```

The rest of this page is the project's `QUIRKS.md`: what Excel does, what strict mode does
instead, and the captured rows behind each.

```{include} ../QUIRKS.md
:start-line: 2
```
