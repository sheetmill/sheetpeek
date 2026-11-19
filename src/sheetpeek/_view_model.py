"""The sheet view's model: zoom, zero display and frozen panes (§4.18)."""

from __future__ import annotations

from dataclasses import dataclass

MIN_ZOOM, MAX_ZOOM = 10, 400
"""The zooms Excel's zoom control allows a person, in percent: the range of the ``zoom=``
override."""
FILE_MIN_ZOOM, FILE_MAX_ZOOM = 1, 1000
"""The file zooms read as they stand: Excel opens and draws a sheet at 5 and at 401
(``freeze_view``); the rest of this range is unverified, and its ends are sheetpeek's bound."""


@dataclass(frozen=True, slots=True)
class SheetView:
    """What the first ``<sheetView>`` says about how the sheet is shown.

    ``zoom`` is the saved zoom in percent (``zoomScale``, 100 when absent; a value outside 10 to
    400 or not a whole number is 100, unverified). ``show_zeros`` is ``showZeros``.
    ``frozen_rows`` and ``frozen_columns`` count the rows and columns a frozen pane keeps in
    place (``ySplit`` and ``xSplit`` of a ``<pane>`` whose state is ``frozen`` or
    ``frozenSplit``), from ``top_left``, the view's top-left cell (``topLeftCell``, ``"A1"``
    when absent). ``pane_state`` is the pane's state as written (``frozen``, ``frozenSplit``,
    ``split``) or ``None`` without a pane.
    """

    zoom: int = 100
    show_zeros: bool = True
    frozen_rows: int = 0
    frozen_columns: int = 0
    top_left: str = "A1"
    pane_state: str | None = None
