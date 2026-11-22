"""Capture Excel's window and map cells to pixels.

The window is found through CoreGraphics (by title, which is the workbook's name) and captured
with ``screencapture -l``, which stores it at the display's scale (2 px a point on a 2x display).
The capture is repeated a second apart until two shots are identical: Excel may still be
redrawing after a zoom change. Cells are mapped to pixels from Excel's own positions in points
(``left position``, ``top``, ``width``, ``height`` of a range, from the sheet's origin) and
marker cells filled with a marker color (:func:`locate`). With two markers on screen, their
pixel boxes give the origin and the scale per axis, and the zoom's nominal scale is checked
against them. A screen too small to show both uses the first marker for the origin and the
nominal scale: screen pixels
per point at zoom 100 (:data:`POINT_PIXELS`) times the display's scale and the zoom. Excel rounds
rows and columns to whole pixels at some zooms (zoom 150 on the Mac: 3.008 across, 3.060 down,
against 3.0), so a one-marker grid holds only where a two-marker run's check shows the nominal
scale holds.
"""

from __future__ import annotations

import subprocess
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from capture.screen import png
from capture.screen.session import excel_windows

MARKER = "FF00FF"
MARKER_HEIGHT = 12
"""Image pixels a marker region spans at least, top to bottom (inclusive bounds)."""
"""Magenta: absent from Excel's chrome and from the test content."""

POINT_PIXELS = {"mac": 1.0, "windows": 96 / 72}
"""Screen pixels per sheet point at zoom 100 on a display of scale 1, by platform."""

GRIDLINE: dict[str, tuple[int, int, int] | None] = {"mac": (212, 212, 212), "windows": None}
"""The default gridline's gray in sRGB, by platform: ``#D4D4D4`` on the Mac (captured). A
platform's gray is stated once a capture in this repository records it; until then it is
``None``, and what needs it is skipped with :data:`NO_GRIDLINE`."""

NO_GRIDLINE = "no capture in this repository records this platform's gridline gray yet"
"""Why a gridline control is skipped."""


def platform_of(header: dict[str, str]) -> str:
    """A run's platform from its header: ``platform``, else ``mac`` (runs before the field)."""
    found = header.get("platform", "mac")
    if found not in POINT_PIXELS:
        raise ValueError(f"unknown platform {found!r} in the run's header")
    return found


def nominal_scale(zoom: int, display_scale: float, platform: str) -> float:
    """Image pixels per sheet point at ``zoom``, as the platform's units give it."""
    return POINT_PIXELS[platform] * display_scale * zoom / 100


STILL = 0.4
"""Seconds without a change after which a window counts as settled."""
GIVE_UP = 8.0
"""Seconds after which a window that has not settled ends the run."""
NEAR = 0.1
"""A stillness that a later change interrupted within this of :data:`STILL` is flagged."""
CONFIRM = 1.0
"""Seconds after settling at which a run's first capture is shot again and must be unchanged."""
AUDIT = 2.0
"""Seconds an audited capture is watched after settling; any change ends the run."""


def _shoot(window_id: int, out: Path) -> bytes:
    """One ``screencapture`` of the window, as the PNG file's bytes."""
    result = subprocess.run(
        ["screencapture", "-x", "-o", "-l", str(window_id), str(out)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    if result.returncode or not out.exists():
        raise RuntimeError(f"screencapture failed: {result.stderr.strip()}")
    return out.read_bytes()


def settle(
    shoot: Callable[[], bytes],
    *,
    confirm: bool = False,
    audit: bool = False,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
) -> tuple[bytes, dict[str, object]]:
    """Shots back to back until the window has not changed for :data:`STILL` seconds.

    Shots are compared as bytes (a still window gives byte-identical files); each is timed when
    it starts. Returns the settled shot and how it settled: the shots taken, when the last
    change was first seen (seconds from the first shot), how long the picture then held (from
    that shot to the last), the longest stillness a later change interrupted, and whether that
    came within :data:`NEAR` of :data:`STILL` (``near``). ``confirm`` shoots once more
    :data:`CONFIRM` seconds later; ``audit`` keeps shooting for :data:`AUDIT` seconds; either
    refuses a change.
    """
    start = clock()
    shot = shoot()
    shots, changed, latest, interrupted = 1, start, start, 0.0
    # each shot is timed when it starts: the picture shows the window then
    while latest - changed < STILL:
        if clock() - start > GIVE_UP:
            raise RuntimeError(f"Excel's window did not settle within {GIVE_UP:g} s: {shots} shots")
        held, latest = latest, clock()
        again = shoot()
        shots += 1
        if again != shot:
            # the picture held from the shot that first showed it to the last one before this
            interrupted = max(interrupted, held - changed)
            shot, changed = again, latest
    found: dict[str, object] = {
        "shots": shots,
        "last_change": round(changed - start, 2),
        "still": round(latest - changed, 2),
        "interrupted": round(interrupted, 2),
        "near": interrupted > 0 and STILL - interrupted <= NEAR,
    }
    if confirm:
        sleep(CONFIRM)
        if shoot() != shot:
            raise RuntimeError("Excel's window changed after settling (the run's first capture)")
        found["confirmed"] = True
    if audit:
        until = clock() + AUDIT
        while clock() < until:
            if shoot() != shot:
                raise RuntimeError("Excel's window changed after settling (audit)")
        found["audited"] = True
    return shot, found


def settled_text(found: dict[str, object]) -> str:
    """How a capture settled, as ``key=value`` fields joined by ``;`` (a read of the run)."""
    return ";".join(f"{key}={str(value).lower()}" for key, value in found.items())


def capture_window(
    title: str, out: Path, *, confirm: bool = False, audit: bool = False
) -> tuple[dict[str, object], png.Image, dict[str, object]]:
    """Capture the on-screen Excel window whose title contains ``title``, once it has settled.

    Returns the window, the picture and how it settled (:func:`settle`).
    """
    for _ in range(20):
        windows = [w for w in excel_windows() if title in w["title"]]
        if windows:
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(f"no Excel window titled {title!r} on screen")
    window = windows[0]
    shot, found = settle(lambda: _shoot(int(window["id"]), out), confirm=confirm, audit=audit)
    out.write_bytes(shot)
    return window, png.read(out), found


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
        # A marker is a cell's full height (at least 20 px at zoom 100 on a 2x display, 15 at
        # zoom 50), though an object may cover most of its width (a shown note leaves a 12 px
        # sliver). Microsoft 365's floating Copilot button, when it is on, is a 15 x 10 px
        # magenta-ish region at the bottom right (Excel 16.115): too short to pass.
        if box[3] - box[1] >= MARKER_HEIGHT:
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


def locate(
    image: png.Image,
    first: dict[str, float],
    second: dict[str, float],
    nominal: float | None = None,
    *,
    one_marker: bool = False,
) -> tuple[Grid, dict[str, object]]:
    """The grid from the markers on screen: both when both show, else the first and ``nominal``.

    With two markers the grid is measured (:func:`grid_from_markers`), and ``nominal`` (if
    given) is compared with the measured scales in ``check``; with one, :func:`grid_from_marker`.
    ``one_marker`` measures from the first marker and ``nominal`` even when both show, to
    compare the two ways on the same image.
    """
    boxes = marker_boxes(image)
    if one_marker and boxes and nominal is not None:
        return grid_from_marker(image, first, nominal, boxes[0])
    if len(boxes) == 2:
        grid, check = grid_from_markers(image, first, second)
        if nominal is not None:
            check["nominal_scale"] = round(nominal, 4)
            check["nominal_off"] = round(
                max(abs(grid.scale_x - nominal), abs(grid.scale_y - nominal)) / nominal, 5
            )
        return grid, check
    if len(boxes) == 1 and nominal is not None:
        return grid_from_marker(image, first, nominal)
    raise RuntimeError(
        f"expected 2 marker regions, or 1 and a nominal scale; found {len(boxes)}: {boxes}"
    )


def grid_from_marker(
    image: png.Image,
    first: dict[str, float],
    scale: float,
    box: tuple[int, int, int, int] | None = None,
) -> tuple[Grid, dict[str, object]]:
    """The grid from the first marker's box and a known scale, the same on both axes.

    ``box`` is the first marker's pixel box, by default the one marker region on screen. Its size
    is checked against the cell's at ``scale``, within 2 pixels each way.
    """
    if box is None:
        boxes = marker_boxes(image)
        if len(boxes) != 1:
            raise RuntimeError(f"expected 1 marker region, found {len(boxes)}: {boxes}")
        box = boxes[0]
    x0, y0, x1, y1 = box
    boxes = [box]
    size = (x1 - x0 + 1, y1 - y0 + 1)
    expected = (first["width"] * scale, first["height"] * scale)
    if any(abs(got - want) > 2 for got, want in zip(size, expected, strict=True)):
        raise RuntimeError(f"the marker box {size} is not the first marker's {expected}")
    grid = Grid(x0 - first["left"] * scale, y0 - first["top"] * scale, scale, scale, scale)
    check = {
        "boxes": boxes,
        "scale": "nominal",
        "nominal_scale": round(scale, 4),
        "first_box_wh": size,
        "first_cell_wh_px": (round(expected[0], 2), round(expected[1], 2)),
        "first_marker_rgb": image.pixel((x0 + x1) // 2, (y0 + y1) // 2),
    }
    return grid, check


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
