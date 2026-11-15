"""The sheet list: each sheet's name, kind, state and tab color, in workbook order (§4.15).

xlio lists the sheets with the kind their relationship type gives and the state and tab color
the file states (``Book._listing``); this module holds the types and the default-sheet rule.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal, TypeAlias

from sheetpeek._model import SheetState

SheetKind: TypeAlias = Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet"]

_SHOWN_KINDS = ("worksheet", "macrosheet")


class NotAWorksheetError(ValueError):
    """The sheet asked for is a chartsheet or a dialog sheet, which sheetpeek does not draw.

    A :class:`ValueError`, as before it had a class of its own. :attr:`Book.sheets` lists such
    sheets with their kind.
    """


@dataclass(frozen=True, slots=True)
class SheetInfo:
    """One sheet of a workbook, as :attr:`Book.sheets` lists it."""

    name: str
    kind: SheetKind
    """``"worksheet"``, ``"chartsheet"``, ``"dialogsheet"`` (a legacy dialog) or
    ``"macrosheet"`` (Excel 4.0 macros), from the workbook's relationship type. Worksheets and
    macro sheets can be read; the others raise :class:`NotAWorksheetError`."""
    state: SheetState
    """``"visible"``, ``"hidden"`` or ``"very_hidden"`` (hidden from Excel's own menus)."""
    tab_color: str | None = None
    """The tab's color as ``RRGGBB``, through the palette and the theme; ``None`` when the file
    gives none, or ``auto``, which Excel draws as an uncolored tab."""

    @property
    def readable(self) -> bool:
        """Whether :meth:`Book.sheet` reads it: a worksheet or a macro sheet."""
        return self.kind in _SHOWN_KINDS


@dataclass(frozen=True, slots=True)
class StatedColor:
    """A ``CT_Color`` as the file states it, in the shape the reader's color resolver takes."""

    type: str
    rgb: str | None = None
    indexed: int | None = None
    theme: int | None = None
    tint: float = 0.0


@dataclass(frozen=True, slots=True)
class ListedSheet:
    """A ``<sheet>`` of the workbook part: what :class:`SheetInfo` needs, color unresolved."""

    name: str
    kind: SheetKind
    state: SheetState
    part: str | None


def from_xlio(color: Any) -> StatedColor | None:
    """An ``xlio.Color`` as a :class:`StatedColor`.

    Its attributes are taken in openpyxl's order, indexed, theme, auto, rgb, as
    :func:`stated_color` takes them.
    """
    if color is None:
        return None
    if color.indexed is not None:
        return StatedColor("indexed", indexed=color.indexed, tint=color.tint)
    if color.theme is not None:
        return StatedColor("theme", theme=color.theme, tint=color.tint)
    if color.auto:
        return StatedColor("auto", tint=color.tint)
    if color.rgb is not None:
        return StatedColor("rgb", rgb=color.rgb, tint=color.tint)
    return None


def default_sheet(sheets: list[ListedSheet], active: ListedSheet | None) -> ListedSheet | None:
    """The sheet Excel opens on that sheetpeek reads.

    The active one when it is a visible worksheet or macro sheet; else the first visible one;
    else the first readable one, hidden or not; ``None`` when no sheet is readable.
    """
    if active is not None and active.state == "visible" and active.kind in _SHOWN_KINDS:
        return active
    readable = [s for s in sheets if s.kind in _SHOWN_KINDS]
    visible = [s for s in readable if s.state == "visible"]
    found = visible or readable
    return found[0] if found else None
