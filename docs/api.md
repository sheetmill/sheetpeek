# API reference

Everything below is importable from `sheetpeek`.

## Reading

```{eval-rst}
.. autofunction:: sheetpeek.read
.. autofunction:: sheetpeek.open
.. autoclass:: sheetpeek.Book
   :members: sheet, sheet_names, date1904, close
```

## Limits and refusals

```{eval-rst}
.. autoclass:: sheetpeek.ReadLimits
.. autoexception:: sheetpeek.PackageError
.. autoexception:: sheetpeek.LimitError
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
.. autoclass:: sheetpeek.Hyperlink
   :members:
.. autoclass:: sheetpeek.Pattern
   :members:
.. autoclass:: sheetpeek.Gradient
   :members:
.. autoclass:: sheetpeek.Borders
   :members:
.. autoclass:: sheetpeek.Side
   :members:
.. py:data:: sheetpeek.Fill

   A cell's fill: a :class:`Pattern` or a :class:`Gradient`.

.. py:data:: sheetpeek.PatternKind

   A pattern fill's ``patternType``: ``"solid"`` or one of the 17 patterns.

.. py:data:: sheetpeek.BorderStyle

   One of the 13 border styles, as the file names them.
.. autoclass:: sheetpeek.Column
   :members:
.. autoclass:: sheetpeek.Row
   :members:
.. autoclass:: sheetpeek.Outline
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
.. autoclass:: sheetpeek.TextRule
   :members:
.. autoclass:: sheetpeek.Gap
   :members:
.. autoclass:: sheetpeek.RenderedBorders
   :members:
.. autoclass:: sheetpeek.Diagonal
   :members:
.. autoclass:: sheetpeek.RenderedOutline
   :members:
.. autoclass:: sheetpeek.OutlineMark
   :members:
.. py:data:: sheetpeek.OutlineMarkKind

   What an :class:`OutlineMark` draws: ``"start"``, ``"bar"``, ``"open"``, ``"closed"`` or
   ``"dot"``.
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

## What is not shown

```{eval-rst}
.. autofunction:: sheetpeek.notices
.. autoclass:: sheetpeek.Notice
   :members:
.. autoclass:: sheetpeek.Support
   :members:
.. py:data:: sheetpeek.SHOWN

   The :class:`Support` of a feature an output shows as Excel does.

.. py:data:: sheetpeek.FEATURES

   Every notice code, mapped to its :class:`Feature`.

.. autoclass:: sheetpeek.Feature
   :members:
.. autoclass:: sheetpeek.Presence
   :members:
.. autoclass:: sheetpeek.NotShownWarning
.. py:data:: sheetpeek.Outcome

   ``"omitted"`` (nothing of the feature is drawn) or ``"approximated"`` (something else is).

.. py:data:: sheetpeek.Reason

   ``"not_implemented"``, ``"not_representable"``, ``"needs_dependency"``, ``"option"``,
   ``"unsafe"`` or ``"file"``.
```
