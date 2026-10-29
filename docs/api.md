# API reference

Everything below is importable from `xlformat`.

## Formatting

```{eval-rst}
.. autofunction:: xlformat.format_value
.. autofunction:: xlformat.format_cell
.. autoclass:: xlformat.FormattedCell
   :members:
```

## Compiled formats

```{eval-rst}
.. autoclass:: xlformat.NumberFormat
   :members: parse, format, format_cell, is_date, is_text, is_general
.. autoclass:: xlformat.Section
.. autoclass:: xlformat.SectionKind
   :members:
   :undoc-members:
```

## Error values

```{eval-rst}
.. autoclass:: xlformat.ExcelError
   :members:
   :undoc-members:
```

## Measuring text

```{eval-rst}
.. autofunction:: xlformat.text_width
.. autofunction:: xlformat.glyph_widths
.. autofunction:: xlformat.digit_width
.. autofunction:: xlformat.overhang
.. autofunction:: xlformat.measured_fonts
```

## Columns and fonts

```{eval-rst}
.. autoclass:: xlformat.Font
   :members: style
.. autofunction:: xlformat.file_column_metrics
.. autofunction:: xlformat.column_metrics
.. autoclass:: xlformat.ColumnMetrics
   :members:
```

## Settings, errors and colors

```{eval-rst}
.. autoclass:: xlformat.Quirks
   :members: EXCEL, STRICT
.. autoexception:: xlformat.FormatError
.. py:data:: xlformat.DEFAULT_PALETTE
   :type: ~collections.abc.Mapping[str, str]

   Section color names to ``RRGGBB`` in Excel's default palette: the eight names
   (``Black``, ``White``, ``Red``, ``Green``, ``Blue``, ``Yellow``, ``Magenta``, ``Cyan``) and
   ``Color1`` to ``Color56``, the names ``FormattedCell.color`` uses. Several indices share
   one color (``Color25`` and ``Color11`` are both navy), so compare colors, not names. The
   mapping is read-only.
```
