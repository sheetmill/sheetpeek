"""Pattern and gradient fills: a layer of boxes under the text, in sheet coordinates.

A solid fill is the cell's own background (``_table.py``). A pattern or a gradient is a box in
this layer, written before the table's text, so text lies over it: the cell and its far
gridlines, ``[x, x + width + 1)``, unless the neighbor to the right or below has a fill of its
own, which Excel paints over that gridline (``ARCHITECTURE.md`` §4.12, captures on Excel 16.114).

- **Patterns** are the captured tiles in device pixels, two to a CSS pixel, as SVG images drawn
  with ``image-rendering: pixelated`` and anchored at the sheet's origin. Excel keeps them in
  screen pixels at any zoom, so they stay in CSS px at any page scale.
- **Linear gradients** are CSS ``linear-gradient`` at ``90° + atan2(w sin d, h cos d)``, Excel's
  angle being taken in the unit cell; the corners take the end stops in both.
- **Path gradients** are an SVG of four bands around the first stop's rectangle.

Each distinct fill is one generated class (``sp-f-`` and letters) in the page's stylesheet,
which the page writes before the sheet; the classes are not a public interface.
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, TextIO

from sheetpeek._fill_border_census import fill_style
from sheetpeek._fill_border_model import TILES, Gradient, Pattern

if TYPE_CHECKING:
    from sheetpeek._rendered import RenderedCell
    from sheetpeek.backends.html._table import Page

_LAYER_CSS = """\
.sheetpeek .sp-fills i, .sheetpeek .sp-borders i { position: absolute; display: block; }
.sheetpeek .sp-fills i { image-rendering: pixelated; background-repeat: repeat; }
"""


class Classes:
    """Generated class names, one per distinct declaration block."""

    def __init__(self, prefix: str) -> None:
        self.prefix = prefix
        self.rules: dict[str, str] = {}

    def name(self, declarations: str) -> str:
        """The class for ``declarations``, made the first time it is asked for."""
        found = self.rules.get(declarations)
        if found is None:
            found = self.rules[declarations] = self.prefix + letters(len(self.rules))
        return found

    def css(self) -> str:
        return "".join(f".sheetpeek .{name} {{ {body} }}\n" for body, name in self.rules.items())


def letters(index: int) -> str:
    """``a`` to ``z``, then ``ba``: a number in letters, since class names hold no digits."""
    text = ""
    while True:
        index, remainder = divmod(index, 26)
        text = chr(97 + remainder) + text
        if not index:
            return text


def svg_uri(svg: str) -> str:
    """An SVG document as a CSS ``url()``, its markup escaped for a data URI."""
    for char, code in (("%", "%25"), ("#", "%23"), ("<", "%3C"), (">", "%3E"), ('"', "%22")):
        svg = svg.replace(char, code)
    return f'url("data:image/svg+xml,{svg}")'


def tile_svg(rows: tuple[str, ...] | list[str], ink: str, paper: str) -> str:
    """A bitmap of ``#`` (``ink``) and ``.`` (``paper``) as an SVG of whole pixels."""
    width, height = len(rows[0]), len(rows)
    rects = "".join(
        f"M{x} {y}h1v1h-1z" for y, row in enumerate(rows) for x, c in enumerate(row) if c == "#"
    )
    return (
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}' "
        f"shape-rendering='crispEdges'><rect width='{width}' height='{height}' fill='#{paper}'/>"
        f"<path fill='#{ink}' d='{rects}'/></svg>"
    )


def solid(page: Page, cell: RenderedCell) -> str | None:
    """The color a cell's own background takes: its fill, unless this layer draws it."""
    if cell.fill is None:
        return None
    style = fill_style(page.sheet.sheet, cell)
    spec = style.fill_spec if style is not None else None
    if spec is None or (isinstance(spec, Pattern) and spec.kind == "solid"):
        return cell.fill
    return None


class Fills:
    """The fill layer of a page: one box per laid-out cell or merge with a pattern or gradient."""

    def __init__(self, page: Page) -> None:
        self.page = page
        self.classes = Classes(f"sp-f-{page.tag}")
        self.boxes: list[str] = []
        sheet = page.sheet
        rows, columns = sheet.rows, sheet.columns
        if not rows or not columns:
            return
        xs = [page.x(c.x_pt or 0) for c in columns]
        xs.append(page.x((columns[-1].x_pt or 0) + (columns[-1].width_pt or 0)))
        ys = [page.y(r.y_pt or 0) for r in rows]
        ys.append(page.y((rows[-1].y_pt or 0) + (rows[-1].height_pt or 0)))
        done: set[object] = set()
        for i, row in enumerate(rows):
            for p, cell in enumerate(row.cells):
                if cell.fill is None or cell.merge in done:
                    continue
                style = fill_style(sheet.sheet, cell)
                spec = style.fill_spec if style is not None else None
                if spec is None or (isinstance(spec, Pattern) and spec.kind == "solid"):
                    continue
                bottom, right = i, p
                if cell.merge is not None:
                    done.add(cell.merge)
                    while bottom + 1 < len(rows) and rows[bottom + 1].cells[p].merge == cell.merge:
                        bottom += 1
                    while right + 1 < len(columns) and row.cells[right + 1].merge == cell.merge:
                        right += 1
                # The far gridlines are this fill's unless the neighbor paints its own fill.
                grow_x = right + 1 == len(columns) or row.cells[right + 1].fill is None
                grow_y = bottom + 1 == len(rows) or rows[bottom + 1].cells[p].fill is None
                left, top = xs[p], ys[i]
                width = xs[right + 1] - left + (1 if grow_x else 0)
                height = ys[bottom + 1] - top + (1 if grow_y else 0)
                self.box(spec, left, top, width, height)

    def box(self, spec: Pattern | Gradient, left: int, top: int, width: int, height: int) -> None:
        if width <= 0 or height <= 0:
            return
        style = f"left:{left}px;top:{top}px;width:{width}px;height:{height}px"
        if isinstance(spec, Pattern):
            tile = TILES[spec.kind]
            # Tiles are in device pixels, two to a CSS pixel, anchored at the sheet's origin.
            tile_w, tile_h = len(tile[0]) / 2, len(tile) / 2
            name = self.classes.name(
                f"background-image:{svg_uri(tile_svg(tile, spec.foreground, spec.background))};"
                f"background-size:{tile_w:g}px {tile_h:g}px"
            )
            dx = (self.page.x(0) - left) % tile_w
            dy = (self.page.y(0) - top) % tile_h
            if dx or dy:
                style += f";background-position:{dx:g}px {dy:g}px"
        else:
            name = self.classes.name(gradient(spec, width, height))
        self.boxes.append(f'<i class="{name}" style="{style}"></i>')

    def css(self) -> str:
        """The layer's stylesheet, empty on a page without patterns or gradients."""
        return self.classes.css()

    def write(self, out: TextIO) -> None:
        if self.boxes:
            out.write(f'<div class="sp-fills">{"".join(self.boxes)}</div>')


def gradient(spec: Gradient, width: int, height: int) -> str:
    """The declarations that draw a gradient on a box of ``width`` by ``height`` px."""
    stops = list(spec.stops)
    if stops[0][0] > 0:
        # Before a first stop past 0 the color runs from the last stop's (one capture).
        stops.insert(0, (0.0, stops[-1][1]))
    if spec.kind == "linear":
        radians = math.radians(spec.degree)
        angle = 90 + math.degrees(math.atan2(width * math.sin(radians), height * math.cos(radians)))
        listed = ",".join(f"#{c} {p * 100:.4g}%" for p, c in stops)
        return f"background-image:linear-gradient({angle % 360:.3f}deg,{listed})"
    return f"background-image:{svg_uri(path_svg(spec, stops, width, height))}"


def path_svg(spec: Gradient, stops: list[tuple[float, str]], w: int, h: int) -> str:
    """Four bands from the first stop's rectangle out to the box's edges, each a gradient."""
    left, right = sorted((spec.left * w, spec.right * w))
    top, bottom = sorted((spec.top * h, spec.bottom * h))
    bands = (
        (f"{left},{top} {right},{top} {w},0 0,0", (0, top, 0, 0)),
        (f"{left},{bottom} {right},{bottom} {w},{h} 0,{h}", (0, bottom, 0, h)),
        (f"{left},{top} {left},{bottom} 0,{h} 0,0", (left, 0, 0, 0)),
        (f"{right},{top} {right},{bottom} {w},{h} {w},0", (right, 0, w, 0)),
    )
    listed = "".join(f"<stop offset='{p:.4g}' stop-color='#{c}'/>" for p, c in stops)
    defs = "".join(
        f"<linearGradient id='b{k}' gradientUnits='userSpaceOnUse' x1='{a:g}' y1='{b:g}' "
        f"x2='{c:g}' y2='{d:g}'>{listed}</linearGradient>"
        for k, (_, (a, b, c, d)) in enumerate(bands)
    )
    # Each band is stroked with its own gradient to close the seams between them.
    polygons = "".join(
        f"<polygon points='{points}' fill='url(#b{k})' stroke='url(#b{k})' stroke-width='0.5'/>"
        for k, (points, _) in enumerate(bands)
    )
    inner = (
        f"<rect x='{left:g}' y='{top:g}' width='{right - left:g}' height='{bottom - top:g}' "
        f"fill='#{spec.stops[0][1]}'/>"
        if right > left and bottom > top
        else ""
    )
    return (
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' "
        f"preserveAspectRatio='none'><defs>{defs}</defs>{polygons}{inner}</svg>"
    )


def layer_css(fills: Fills, borders_css: str) -> str:
    """The stylesheet of the fill and border layers: nothing on a page without either."""
    own = fills.css()
    if not own and not borders_css:
        return ""
    return _LAYER_CSS + own + borders_css
