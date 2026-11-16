"""Measuring ink in a window capture: boxes, bands of rows, horizontal rules, on/off patterns.

Boxes are image pixels ``(x0, y0, x1, y1)``, end exclusive. Text is black (or a marked color) on
white; gridlines are luma 212 on the Mac (``window.GRIDLINE``), so the default ink threshold
(150) leaves them out. Standard library only.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from typing import Any

from capture.screen.run import luma, runs

DARK = 150
"""Luma below which a pixel counts as text ink (the Mac's gridlines are 212)."""

Box = tuple[int, int, int, int]
Test = Callable[[tuple[int, int, int]], bool]


def is_dark(rgb: tuple[int, int, int]) -> bool:
    return luma(rgb) < DARK


def nearest(
    rgb: tuple[int, int, int], palette: dict[str, tuple[int, int, int]], tolerance: float = 70.0
) -> str:
    """The palette name closest to ``rgb``, or ``"?"`` when none is within ``tolerance``."""
    best, distance = "?", tolerance
    for name, reference in palette.items():
        d = sum((a - b) ** 2 for a, b in zip(rgb, reference, strict=True)) ** 0.5
        if d < distance:
            best, distance = name, d
    return best


def cells_box(cap: Any, first: str, last: str | None = None, inset: int = 0) -> Box:
    """The box from ``first``'s top left to ``last``'s bottom right, shrunk by ``inset`` px."""
    last = last or first
    return (
        round(cap.edge(first, "left")) + inset,
        round(cap.edge(first, "top")) + inset,
        round(cap.edge(last, "right")) - inset,
        round(cap.edge(last, "bottom")) - inset,
    )


def visible(cap: Any, cell: str) -> bool:
    """Whether the cell lies inside the captured image (a large zoom may push it out)."""
    return (
        cap.edge(cell, "left") >= 0
        and cap.edge(cell, "right") < cap.image.width
        and cap.edge(cell, "top") >= 0
        and cap.edge(cell, "bottom") < cap.image.height
    )


def ink(cap: Any, box: Box, test: Test = is_dark) -> Box | None:
    """The bounding box (end exclusive) of the pixels passing ``test``, or ``None``."""
    x0, y0, x1, y1 = box
    xs, ys = [], []
    for y in range(y0, y1):
        for x in range(x0, x1):
            if test(cap.image.pixel(x, y)):
                xs.append(x)
                ys.append(y)
    if not xs:
        return None
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1


def bands(cap: Any, box: Box, test: Test = is_dark, axis: str = "y") -> list[tuple[int, int]]:
    """Runs of rows (axis y) or columns (axis x) holding any ink: ``[(start, length)]`` in px."""
    x0, y0, x1, y1 = box
    if axis == "y":
        flags = [any(test(cap.image.pixel(x, y)) for x in range(x0, x1)) for y in range(y0, y1)]
        base = y0
    else:
        flags = [any(test(cap.image.pixel(x, y)) for y in range(y0, y1)) for x in range(x0, x1)]
        base = x0
    return [(base + start, length) for on, start, length in runs(flags) if on]


def rule_rows(
    cap: Any, box: Box, min_length: int, test: Test = is_dark
) -> list[tuple[int, int, int, int]]:
    """Horizontal lines: bands of rows whose longest ink run is at least ``min_length`` px.

    Returns ``[(y, height, x_start, x_end)]``, the x extent being the longest runs' in the band.
    """
    x0, y0, x1, y1 = box
    hits: list[tuple[int, int] | None] = []
    for y in range(y0, y1):
        best = (0, 0)
        for on, start, length in runs([test(cap.image.pixel(x, y)) for x in range(x0, x1)]):
            if on and length > best[1]:
                best = (x0 + start, length)
        hits.append(best if best[1] >= min_length else None)
    found = []
    for on, start, length in runs([hit is not None for hit in hits]):
        if on:
            band = [hit for hit in hits[start : start + length] if hit is not None]
            found.append(
                (y0 + start, length, min(h[0] for h in band), max(h[0] + h[1] for h in band))
            )
    return found


def pattern(flags: list[bool], scale: float) -> str:
    """The repeating on/off pattern of a scanline in points, partial runs at the ends dropped.

    ``"solid"``, ``"empty"``, ``"on 3 off 1"`` (one period, starting with ink), or
    ``"irregular: ..."`` with the first runs when nothing repeats.
    """
    found = runs(flags)
    if len(found) == 1:
        return "solid" if found[0][0] else "empty"
    inner = [(on, length) for on, _, length in found[1:-1]]
    if not inner:
        return "solid" if sum(flags) > len(flags) / 2 else "empty"
    for k in range(1, len(inner) // 2 + 1):
        if all(inner[i] == inner[i + k] for i in range(len(inner) - k)):
            period = inner[:k]
            while not period[0][0]:
                period = period[1:] + period[:1]
            return _runs_text(period, scale)
    return "irregular: " + _runs_text(inner[:12], scale)


def _runs_text(found: Iterable[tuple[bool, int]], scale: float) -> str:
    return " ".join(f"{'on' if on else 'off'} {length / scale:g}" for on, length in found)


def pt(cap: Any, pixels: float) -> str:
    """A length in image pixels as points at the capture's scale, for a TSV field."""
    return f"{round(pixels / cap.scale, 3):g}"
