"""Capture Excel's window and map cells to pixels.

The window is found through CoreGraphics (by title, which is the workbook's name) and captured
with ``screencapture -l``, which stores it at the display's scale (2 px a point on a 2x display).
The capture is repeated a second apart until two shots are identical: Excel may still be
redrawing after a zoom change. Cells are mapped to pixels from Excel's own positions in points
(``left position``, ``top``, ``width``, ``height`` of a range, from the sheet's origin) and two
marker cells filled with a marker color: their pixel boxes give the origin and the scale per
axis, and the scale is checked against the window's own size.
"""

from __future__ import annotations

import subprocess
import time
from dataclasses import dataclass
from pathlib import Path

from capture.screen import png
from capture.screen.session import excel_windows

MARKER = "FF00FF"
"""Magenta: absent from Excel's chrome and from the test content."""


def capture_window(title: str, out: Path) -> tuple[dict[str, object], png.Image]:
    """Capture the on-screen Excel window whose title contains ``title``."""
    for _ in range(20):
        windows = [w for w in excel_windows() if title in w["title"]]
        if windows:
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(f"no Excel window titled {title!r} on screen")
    window = windows[0]
    # Excel may still be redrawing after a zoom change (seen at zoom 200): capture until two
    # shots a second apart are identical, at most 8 times.
    previous = None
    for _ in range(8):
        time.sleep(1.0)
        result = subprocess.run(
            ["screencapture", "-x", "-o", "-l", str(window["id"]), str(out)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if result.returncode or not out.exists():
            raise RuntimeError(f"screencapture failed: {result.stderr.strip()}")
        data = out.read_bytes()
        if data == previous:
            break
        previous = data
    else:
        raise RuntimeError("Excel's window did not settle within 8 captures")
    return window, png.read(out)


def is_marker(rgb: tuple[int, int, int]) -> bool:
    r, g, b = rgb
    return r > 180 and b > 180 and g < 90  # tolerant: the capture is in the display's profile


def marker_boxes(image: png.Image) -> list[tuple[int, int, int, int]]:
    """Bounding boxes (x0, y0, x1, y1 inclusive) of marker-colored regions, left to right.

    Regions are separated by columns or rows free of marker pixels (enough for markers placed
    apart in both directions).
    """
    xs: dict[int, list[int]] = {}
    for y in range(image.height):
        row = image.rows[y]
        step = image.channels
        for x in range(image.width):
            i = x * step
            if row[i] > 180 and row[i + 2] > 180 and row[i + 1] < 90:
                xs.setdefault(x, []).append(y)
    if not xs:
        return []
    columns = sorted(xs)
    groups, current = [], [columns[0]]
    for x in columns[1:]:
        if x == current[-1] + 1:
            current.append(x)
        else:
            groups.append(current)
            current = [x]
    groups.append(current)
    boxes = []
    for group in groups:
        ys = [y for x in group for y in xs[x]]
        box = (group[0], min(ys), group[-1], max(ys))
        # A marker is a cell's full height (at least 20 px at zoom 100 on a 2x display), though
        # an object may cover most of its width (a shown note leaves a 12 px sliver). Microsoft
        # 365's floating Copilot button, when it is on, is a 15 x 10 px magenta-ish region at the
        # bottom right (Excel 16.115): too short to pass.
        if box[3] - box[1] >= 20:
            boxes.append(box)
    return boxes


@dataclass
class Grid:
    """Points (Excel's cell positions) to image pixels."""

    origin_x: float
    origin_y: float
    scale: float
    """The mean of the two axes' scales, for converting lengths to points."""
    scale_x: float = 0.0
    scale_y: float = 0.0

    def x(self, points: float) -> float:
        return self.origin_x + points * (self.scale_x or self.scale)

    def y(self, points: float) -> float:
        return self.origin_y + points * (self.scale_y or self.scale)


def grid_from_markers(image: png.Image, first: dict, second: dict) -> tuple[Grid, dict]:
    """The grid from two marker cells' geometry (points) and their pixel boxes.

    Each box's left and top edge are taken as the cell's left and top position; how the fill sits
    relative to the gridlines is reported, not assumed: ``check`` compares box sizes with the cells'
    sizes in points times the scale.
    """
    boxes = marker_boxes(image)
    if len(boxes) != 2:
        raise RuntimeError(f"expected 2 marker regions, found {len(boxes)}: {boxes}")
    (ax0, ay0, ax1, ay1), (bx0, by0, bx1, by1) = boxes
    scale_x = (bx0 - ax0) / (second["left"] - first["left"])
    scale_y = (by0 - ay0) / (second["top"] - first["top"])
    scale = (scale_x + scale_y) / 2
    # Per axis: at zooms other than 100 and 200 Excel rounds rows and columns to whole pixels, so
    # the two axes' scales differ (zoom 150: 3.008 across, 3.060 down, measured).
    grid = Grid(
        ax0 - first["left"] * scale_x, ay0 - first["top"] * scale_y, scale, scale_x, scale_y
    )
    check = {
        "boxes": boxes,
        "scale_x": round(scale_x, 4),
        "scale_y": round(scale_y, 4),
        "first_box_wh": (ax1 - ax0 + 1, ay1 - ay0 + 1),
        "first_cell_wh_px": (round(first["width"] * scale, 2), round(first["height"] * scale, 2)),
        "second_box_wh": (bx1 - bx0 + 1, by1 - by0 + 1),
        "second_cell_wh_px": (
            round(second["width"] * scale, 2),
            round(second["height"] * scale, 2),
        ),
        "first_marker_rgb": image.pixel((ax0 + ax1) // 2, (ay0 + ay1) // 2),
    }
    return grid, check
