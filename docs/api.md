# API reference

Everything below is importable from `sheetpeek`.

## Reading

```{eval-rst}
.. autofunction:: sheetpeek.read
.. autofunction:: sheetpeek.open
.. autoclass:: sheetpeek.Book
   :members: sheet, sheet_names, date1904, close
```

## The sheet model

```{eval-rst}
.. autoclass:: sheetpeek.Sheet
   :members:
.. autoclass:: sheetpeek.Cell
   :members:
.. autoclass:: sheetpeek.CellStyle
   :members:
.. autoclass:: sheetpeek.Font
   :members:
.. autoclass:: sheetpeek.Column
   :members:
.. autoclass:: sheetpeek.Row
   :members:
.. autoclass:: sheetpeek.CellRange
   :members: parse, rows, columns, intersects
```

## Layout

```{eval-rst}
.. automethod:: sheetpeek.Sheet.layout
   :no-index:
.. autoclass:: sheetpeek.RenderedSheet
   :members:
.. autoclass:: sheetpeek.RenderedRow
   :members:
.. autoclass:: sheetpeek.RenderedCell
   :members:
.. autoclass:: sheetpeek.RenderedColumn
   :members:
.. autoclass:: sheetpeek.TextRun
   :members:
.. autoclass:: sheetpeek.Gap
   :members:
```

## Rendering

```{eval-rst}
.. autofunction:: sheetpeek.render
.. autofunction:: sheetpeek.register_backend
.. autoclass:: sheetpeek.Backend
   :members:
.. autoclass:: sheetpeek.backends.TextBackend
   :members: render
.. autoclass:: sheetpeek.backends.RichBackend
   :members: render
.. autoclass:: sheetpeek.backends.MarkdownBackend
   :members: render
.. autoclass:: sheetpeek.backends.DelimitedBackend
   :members: render
.. autoclass:: sheetpeek.backends.HTMLBackend
   :members: render
```
