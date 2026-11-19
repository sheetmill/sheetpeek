"""A synthetic screen capture for testing a suite's measuring without Excel.

A white image with cells laid out from the worksheet's own column widths and row heights at 2
pixels a point (zoom 100 on a 2x display), in which a test paints known marks; ``capture()``
gives the ``Capture`` a suite's ``measure`` takes. Cells left of column A or above row 1 (the
headings) are not modelled: a test that needs them paints at negative points through the grid
it builds itself.
"""

from capture.screen import png
from capture.screen.run import Capture
from capture.screen.window import Grid
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

SCALE = 2


def _geometry(ws, cells):
    """Each cell's left, top, width and height in points, from the worksheet's sizes."""

    def width(column):
        found = ws.column_dimensions[get_column_letter(column)].width
        return round(found * 7 + 5) * 0.75 if found else 48.0

    def height(row):
        return ws.row_dimensions[row].height or 15.0

    geometry = {}
    for cell in cells:
        letters, row = coordinate_from_string(cell)
        column = column_index_from_string(letters)
        geometry[cell] = {
            "left": sum(width(c) for c in range(1, column)),
            "top": sum(height(r) for r in range(1, row)),
            "width": width(column),
            "height": height(row),
        }
    return geometry


class Canvas:
    """A white image with cells at known places, and a ``Capture`` of it."""

    def __init__(self, ws, cells, rows=()):
        self.geometry = _geometry(ws, cells)
        right = max(g["left"] + g["width"] for g in self.geometry.values())
        bottom = max(g["top"] + g["height"] for g in self.geometry.values())
        self.width, self.height = round(right * SCALE) + 40, round(bottom * SCALE) + 40
        self.rows = [bytearray(b"\xff" * (self.width * 3)) for _ in range(self.height)]
        self.reads = list(rows)

    def paint(self, x0, y0, x1, y1, rgb):
        for y in range(y0, y1):
            for x in range(x0, x1):
                self.rows[y][3 * x : 3 * x + 3] = bytes(rgb)

    def capture(self):
        image = png.Image(self.width, self.height, 3, self.rows, "sRGB")
        grid = Grid(0.0, 0.0, float(SCALE))
        return Capture(100, image, grid, {}, self.geometry, self.reads)

    def box(self, cell):
        g = self.geometry[cell]
        return (
            round(g["left"] * SCALE),
            round(g["top"] * SCALE),
            round((g["left"] + g["width"]) * SCALE),
            round((g["top"] + g["height"]) * SCALE),
        )
