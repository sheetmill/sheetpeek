"""The screen suites' measuring boxes (``capture.screen.ink``), on hand-built captures."""

import pytest
from capture.screen import png
from capture.screen.ink import cells_box, content_box, gridline_pixels, ink
from capture.screen.run import Capture
from capture.screen.window import POINT_PIXELS, Grid

GRAY = (212, 212, 212)
BLACK = (0, 0, 0)
CELL = {"C4": {"left": 96.0, "top": 45.0, "width": 48.0, "height": 15.0}}


def _capture(platform, display, zoom):
    scale = POINT_PIXELS[platform] * display * zoom / 100
    width, height = round(160 * scale), round(80 * scale)
    rows = [bytearray(b"\xff" * (width * 3)) for _ in range(height)]
    image = png.Image(width, height, 3, rows, "sRGB")
    return Capture(zoom, image, Grid(0.0, 0.0, scale), {}, CELL, [], platform)


def _paint(cap, x0, y0, x1, y1, rgb):
    for y in range(y0, y1):
        for x in range(x0, x1):
            cap.image.rows[y][3 * x : 3 * x + 3] = bytes(rgb)


@pytest.mark.parametrize(
    ("platform", "display", "zoom", "line"),
    [("mac", 2, 100, 2), ("mac", 2, 200, 2), ("windows", 1, 100, 1), ("windows", 1, 200, 1)],
)
def test_a_gridline_is_one_screen_point_at_every_zoom(platform, display, zoom, line):
    assert gridline_pixels(_capture(platform, display, zoom)) == line


@pytest.mark.parametrize(("platform", "display"), [("mac", 2), ("windows", 1)])
def test_the_content_box_keeps_the_last_row_and_leaves_out_the_leading_gridline(platform, display):
    cap = _capture(platform, display, 100)
    x0, y0, x1, y1 = cells_box(cap, "C4")
    line = gridline_pixels(cap)
    # the cell's own gridlines on its first rows and columns, a rule on its last pixel row
    _paint(cap, x0, y0, x1, y0 + line, GRAY)
    _paint(cap, x0, y0, x0 + line, y1, GRAY)
    _paint(cap, x0 + 10, y1 - 1, x0 + 30, y1, BLACK)
    box = content_box(cap, "C4")
    assert box == (x0 + line, y0 + line, x1, y1)
    assert ink(cap, box) == (x0 + 10, y1 - 1, x0 + 30, y1)
    assert ink(cap, box, lambda rgb: rgb == GRAY) is None
