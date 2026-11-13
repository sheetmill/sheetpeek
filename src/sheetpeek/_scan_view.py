"""The sheet view as the file states it: zoom, zero display and frozen panes (§4.18).

Read from the first ``<sheetView>`` (the main window's), in the structure scan.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from sheetpeek._structure import _NS, Scan, Structure, handler
from sheetpeek._view_model import MAX_ZOOM, MIN_ZOOM, SheetView

KEY = "view"


def _whole(value: str | None) -> int | None:
    """A whole number as Excel writes it (``"3"``, ``"3.0"``), else ``None``."""
    if value is None or not value.isascii():
        return None
    try:
        number = float(value)
    except ValueError:
        return None
    return int(number) if number.is_integer() and 0 <= number <= 2**31 else None


@handler(f"{_NS}sheetView")
def _view(state: Scan, element: ET.Element) -> None:
    if KEY in state.extra:
        return
    zoom = _whole(element.get("zoomScale"))
    zoom = zoom if zoom is not None and MIN_ZOOM <= zoom <= MAX_ZOOM else 100
    show_zeros = element.get("showZeros") not in ("0", "false")
    top_left = element.get("topLeftCell") or "A1"
    pane = element.find(f"{_NS}pane")
    rows = columns = 0
    pane_state = None
    if pane is not None:
        pane_state = pane.get("state", "split")
        if pane_state in ("frozen", "frozenSplit"):
            rows = _whole(pane.get("ySplit")) or 0
            columns = _whole(pane.get("xSplit")) or 0
    state.extra[KEY] = SheetView(zoom, show_zeros, rows, columns, top_left, pane_state)


def sheet_view(structure: Structure) -> SheetView:
    """The view the scan found, or the defaults without a ``<sheetView>``."""
    found = structure.extra.get(KEY)
    return found if isinstance(found, SheetView) else SheetView()
