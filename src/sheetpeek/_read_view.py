"""The sheet view as the file states it: zoom, zero display and frozen panes (§4.18).

Read from the first ``<sheetView>`` (the main window's), as xlio reads it.
"""

from __future__ import annotations

import xlio

from sheetpeek._view_model import FILE_MAX_ZOOM, FILE_MIN_ZOOM, SheetView


def _whole(value: float | None) -> int | None:
    """A whole number of rows or columns, else ``None``."""
    if value is None or not float(value).is_integer() or not 0 <= value <= 2**31:
        return None
    return int(value)


def sheet_view(views: tuple[xlio.SheetView, ...]) -> SheetView:
    """The first view's zoom, zero display and frozen pane; the defaults without a view.

    A zoom that is not whole or not a number is 100, as Excel opens it; a whole one is kept,
    past Excel's 10 to 400 too, as Excel opens and draws 5 and 401 (``freeze_view``), within
    sheetpeek's bound of 1 to 1000 (100 past it); a frozen pane's ``ySplit`` and
    ``xSplit`` count rows and columns from the view's top-left cell, and one that is not a whole
    number freezes nothing that way.
    """
    if not views:
        return SheetView()
    view = views[0]
    zoom = _whole(view.zoom_scale)
    zoom = zoom if zoom is not None and FILE_MIN_ZOOM <= zoom <= FILE_MAX_ZOOM else 100
    pane = view.pane
    rows = columns = 0
    state = None
    if pane is not None:
        state = pane.state
        if state in ("frozen", "frozenSplit"):
            rows = _whole(pane.y_split) or 0
            columns = _whole(pane.x_split) or 0
    top_left = str(view.top_left_cell) if view.top_left_cell else "A1"
    return SheetView(zoom, bool(view.show_zeros), rows, columns, top_left, state)
