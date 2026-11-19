"""Fills and borders: what a cell's ``fillId`` and ``borderId`` resolve to, and Excel's rules.

The rules are ``ARCHITECTURE.md`` §4.12: the pattern tiles are captured by the screen suite
``fills``; the colors a pattern's automatic colors take, the conflict order of border styles and
the tie between two colors are unverified.
"""

from __future__ import annotations

import itertools
import math
from dataclasses import dataclass
from typing import Literal, TypeAlias

PatternKind: TypeAlias = Literal[
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
]
"""A pattern fill's ``patternType``; ``none`` is no fill and has no :class:`Pattern`."""

BorderStyle: TypeAlias = Literal[
    "double",
    "thick",
    "medium",
    "mediumDashed",
    "mediumDashDot",
    "slantDashDot",
    "mediumDashDotDot",
    "thin",
    "dashed",
    "dotted",
    "dashDot",
    "dashDotDot",
    "hair",
]
"""A border side's style, as the file names it."""

# Captured tiles, one character a screen point, rows top to bottom, anchored at the sheet's
# origin; '#' is the pattern (foreground) color, '.' the background (screen suite fills,
# Excel 16.115, which draws a tile pixel as a whole screen point at any zoom).
TILES: dict[str, tuple[str, ...]] = {
    "mediumGray": ("#.", ".#"),
    "darkGray": ("#.##", "###."),
    "lightGray": ("..#.", "#..."),
    "darkHorizontal": (".", ".", "#", "#"),
    "lightHorizontal": (".", ".", "#", "."),
    "darkVertical": ("..##",),
    "lightVertical": ("..#.",),
    "darkDown": ("##..", ".##.", "..##", "#..#"),
    "darkUp": ("..##", ".##.", "##..", "#..#"),
    "lightDown": ("#...", ".#..", "..#.", "...#"),
    "lightUp": ("...#", "..#.", ".#..", "#..."),
    "darkGrid": ("#..#", ".##.", ".##.", "#..#"),
    "darkTrellis": ("####", ".##.", "####", "#..#"),
    "lightGrid": ("..#.", "..#.", "####", "..#."),
    "lightTrellis": ("#...", ".#.#", "..#.", ".#.#"),
    "gray125": ("#...", "....", "..#.", "...."),
    "gray0625": ("..#.....", "........", "......#.", "........"),
}

STYLES: tuple[BorderStyle, ...] = (
    "double",
    "thick",
    "medium",
    "mediumDashed",
    "mediumDashDot",
    "slantDashDot",
    "mediumDashDotDot",
    "thin",
    "dashed",
    "dotted",
    "dashDot",
    "dashDotDot",
    "hair",
)
"""Every border style, strongest first: the order in which one side wins over another
(weight first, then this fixed order within a weight; captured, ``border_conflicts``)."""

_RANK = {style: rank for rank, style in enumerate(STYLES)}
BLACK = "000000"
WHITE = "FFFFFF"


def share(kind: str) -> float:
    """The part of a pattern's tile in its foreground color: 1 for ``solid``."""
    tile = TILES.get(kind)
    if tile is None:
        return 1.0
    return sum(row.count("#") for row in tile) / sum(len(row) for row in tile)


def blend(foreground: str, background: str, part: float) -> str:
    """Two ``RRGGBB`` colors mixed channel by channel, ``part`` of the first."""
    channels = []
    for k in (0, 2, 4):
        a, b = int(foreground[k : k + 2], 16), int(background[k : k + 2], 16)
        channels.append(round(a * part + b * (1 - part)))
    return "".join(f"{c:02X}" for c in channels)


@dataclass(frozen=True, slots=True)
class Pattern:
    """A pattern fill: ``solid`` or one of the 17 patterns, in two colors.

    ``fg`` draws the pattern (the whole of a ``solid`` fill) and ``bg`` the rest, as ``RRGGBB``;
    ``None`` is automatic, which Excel draws black for ``fg`` and white for ``bg``.
    """

    kind: PatternKind
    fg: str | None = None
    bg: str | None = None

    @property
    def foreground(self) -> str:
        """The pattern's color as drawn."""
        return self.fg or BLACK

    @property
    def background(self) -> str:
        """The background's color as drawn."""
        return self.bg or WHITE

    @property
    def color(self) -> str:
        """The one color shown by outputs that show one, the two mixed by the pattern's share."""
        if self.kind == "solid":
            return self.foreground
        return blend(self.foreground, self.background, share(self.kind))


@dataclass(frozen=True, slots=True)
class Gradient:
    """A gradient fill.

    ``linear`` runs along ``degree``, clockwise from left to right and taken in the unit cell;
    ``path`` runs from the rectangle ``left``, ``right``, ``top``, ``bottom`` (fractions of the
    cell, in the first color drawn) out to the cell's edges in rectangular bands. ``stops`` are
    ``(position, RRGGBB)`` in order of position (stops at one position in the file's order),
    positions from 0 to 1; :attr:`drawn_stops` gives them as Excel draws them.
    """

    kind: Literal["linear", "path"]
    stops: tuple[tuple[float, str], ...]
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0

    @property
    def drawn_stops(self) -> tuple[tuple[float, str], ...]:
        """The stops as Excel draws them (screen suite ``gradients``), for any output to draw.

        Normalized: black is implied at 0 before a first stop past 0, and at 1 after a last stop
        before 1; of stops at one position, the last in the file ends the stretch before it and
        the first starts the stretch after, so each such group is reversed. A gradient without
        stops is never a ``Gradient``: the reader makes it a solid black ``Pattern``.
        """
        found: list[tuple[float, str]] = []
        for _, group in itertools.groupby(self.stops, key=lambda stop: stop[0]):
            found.extend(reversed(list(group)))
        if found[0][0] > 0:
            found.insert(0, (0.0, BLACK))
        if found[-1][0] < 1:
            found.append((1.0, BLACK))
        return tuple(found)

    def color_at(self, t: float) -> str:
        """The color at ``t`` from 0 to 1 along the gradient (:attr:`drawn_stops`)."""
        stops = self.drawn_stops
        for (p0, c0), (p1, c1) in itertools.pairwise(stops):
            if p0 <= t <= p1 and p1 > p0:
                return blend(c1, c0, (t - p0) / (p1 - p0))
        return stops[-1][1]

    @property
    def color(self) -> str:
        """The one color shown by outputs that show one, the gradient's at its midpoint."""
        return self.color_at(0.5)


Fill: TypeAlias = Pattern | Gradient


@dataclass(frozen=True, slots=True)
class Side:
    """One border side: its style and its color, ``RRGGBB`` or ``None`` for automatic (black)."""

    style: BorderStyle
    color: str | None = None

    @property
    def drawn_color(self) -> str:
        """The color as drawn."""
        return self.color or BLACK


@dataclass(frozen=True, slots=True)
class Borders:
    """A cell's border sides, each a :class:`Side` or ``None``, and its diagonals.

    ``diagonal`` is set only when a diagonal is drawn: the file gives it a style and sets
    ``diagonal_up`` (bottom left to top right) or ``diagonal_down`` (top left to bottom right).
    """

    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    diagonal: Side | None = None
    diagonal_up: bool = False
    diagonal_down: bool = False

    @property
    def sides(self) -> tuple[Side | None, Side | None, Side | None, Side | None]:
        """Left, right, top and bottom."""
        return self.left, self.right, self.top, self.bottom


def luma(color: str) -> float:
    """Rec. 601 luma of an ``RRGGBB`` color."""
    red, green, blue = (int(color[k : k + 2], 16) for k in (0, 2, 4))
    return 0.299 * red + 0.587 * green + 0.114 * blue


def winner(a: Side | None, b: Side | None) -> Side | None:
    """The side Excel draws where two cells' sides meet: never both, never by position.

    The stronger style wins (:data:`STYLES`, captured); of one style, the darker color by Rec.
    601 luma and of equal luma the larger ``RRGGBB``: a known deviation, as Excel's own color
    rule is not found (``border_conflicts``: luma fits 28 of 68 cases, no single measure all).
    """
    if a is None or b is None:
        return a or b
    if a.style != b.style:
        return a if _RANK[a.style] < _RANK[b.style] else b
    ca, cb = a.drawn_color, b.drawn_color
    la, lb = luma(ca), luma(cb)
    if la != lb:
        return a if la < lb else b
    return a if int(ca, 16) >= int(cb, 16) else b


@dataclass(frozen=True, slots=True)
class Diagonal:
    """A cell's or a merge's diagonals, over a rectangle of the laid-out grid.

    ``top`` to ``bottom`` index ``RenderedSheet.rows``, and
    ``left`` to ``right`` their cells, inclusive.
    """

    top: int
    left: int
    bottom: int
    right: int
    side: Side
    up: bool
    down: bool


@dataclass(frozen=True, slots=True)
class SlantedBorders:
    """An angled cell's own side borders, which Excel slants with its text (§4.17).

    ``row`` indexes ``RenderedSheet.rows`` and ``column`` its cells. The sides form a
    parallelogram whose bottom edge is the cell's bottom edge: the left and right sides lean at
    the text's ``angle``, and the top edge lies along the cell's top edge, ``shift_pt`` to the
    right of it (to the left when negative), which is the cell's height times the cotangent of
    the angle. These sides are not in the ``horizontal`` or ``vertical`` lines of
    :class:`RenderedBorders`, where a neighbor's own side on a shared edge is drawn upright.
    """

    row: int
    column: int
    angle: int
    shift_pt: float
    left: Side | None
    right: Side | None
    top: Side | None
    bottom: Side | None


@dataclass(frozen=True, slots=True)
class RenderedBorders:
    """The border lines a laid-out sheet draws, one resolved side per edge (``excel`` mode).

    ``horizontal[k][p]`` is the line on the top edge of grid row ``k`` (``k == len(rows)`` is
    the bottom edge of the last) in grid column ``p``; ``vertical[i][q]`` the line on the left
    edge of grid column ``q`` (``q == len(columns)`` the right edge of the last) in grid row
    ``i``. Conflicts are resolved, sides inside a merge dropped and vertical sides that spilled
    text crosses removed, so an output draws each edge as it stands.
    """

    horizontal: tuple[tuple[Side | None, ...], ...]
    vertical: tuple[tuple[Side | None, ...], ...]
    diagonals: tuple[Diagonal, ...]
    slanted: tuple[SlantedBorders, ...] = ()
    """The borders of angled cells, slanted with their text, in grid order."""


def finite(value: object, default: float = 0.0) -> float:
    """A number from the file as a finite float, else ``default``."""
    try:
        number = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default
    return number if math.isfinite(number) else default


def unit(value: object) -> float:
    """A fraction from the file, between 0 and 1."""
    return min(1.0, max(0.0, finite(value)))
