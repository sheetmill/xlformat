# API reference

Everything below is importable from `xlformat`.

## Formatting

```{eval-rst}
.. autofunction:: xlformat.format_value
.. autofunction:: xlformat.format_cell
.. autofunction:: xlformat.text_function
.. autofunction:: xlformat.builtin_format
.. autodata:: xlformat.BUILTIN_FORMATS
   :no-value:
.. autoclass:: xlformat.FormattedCell
   :members:
```

## Compiled formats

```{eval-rst}
.. autoclass:: xlformat.NumberFormat
   :members: parse, format, format_cell, text_function, is_date, is_text, is_general
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
.. autofunction:: xlformat.autofit_width
.. autofunction:: xlformat.glyph_widths
.. autofunction:: xlformat.digit_width
.. autofunction:: xlformat.overhang
.. autofunction:: xlformat.measured_fonts
.. autofunction:: xlformat.font_substitutes
```

## Columns and fonts

```{eval-rst}
.. autoclass:: xlformat.Font
   :members: style
.. autofunction:: xlformat.file_column_metrics
.. autofunction:: xlformat.column_metrics
.. autofunction:: xlformat.column_metrics_from_points
.. autoclass:: xlformat.ColumnMetrics
   :members:
.. autofunction:: xlformat.row_metrics
.. autoclass:: xlformat.RowMetrics
   :members:
.. autofunction:: xlformat.vertical_metrics
.. autoclass:: xlformat.VerticalMetrics
   :members:
.. autofunction:: xlformat.ink_metrics
.. autoclass:: xlformat.InkMetrics
.. autofunction:: xlformat.text_insets
.. autoclass:: xlformat.TextInsets
   :members:
.. autofunction:: xlformat.rule_metrics
.. autoclass:: xlformat.RuleMetrics
   :members:
```

## Settings, errors and colors

```{eval-rst}
.. autoclass:: xlformat.Quirks
   :members: EXCEL, STRICT
.. autoexception:: xlformat.FormatError
.. autofunction:: xlformat.section_rgb
.. py:data:: xlformat.DEFAULT_PALETTE
   :type: ~collections.abc.Mapping[str, str]

   Section color names to ``RRGGBB`` in Excel's default palette: the eight names
   (``Black``, ``White``, ``Red``, ``Green``, ``Blue``, ``Yellow``, ``Magenta``, ``Cyan``) and
   ``Color1`` to ``Color56``, the names ``FormattedCell.color`` uses. Several indices share
   one color (``Color25`` and ``Color11`` are both navy), so compare colors, not names. The
   mapping is read-only.
.. autoclass:: xlformat.Platform
   :members: MAC, WINDOWS
.. autoclass:: xlformat.RegionalSettings
   :members: EN_US_MAC, EN_US_WINDOWS
.. autodata:: xlformat.DEFAULT_PLATFORM
```
