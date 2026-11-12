"""The sheet list: each sheet's name, kind, state and tab color, in workbook order (§4.15).

A sheet's kind is told only by the relationship type from the workbook part: openpyxl loads a
dialog sheet as an empty worksheet. A tab color is read from the head of the sheet's part
(``sheetPr/tabColor``), which openpyxl's read-only worksheets do not parse; the parse stops at the
first child after ``sheetPr``, so a sheet's cells are never read for it.
"""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from typing import Literal, TypeAlias

from sheetpeek._model import SheetState
from sheetpeek._package import related, workbook_part

SheetKind: TypeAlias = Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet"]

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_KINDS: tuple[tuple[str, SheetKind], ...] = (
    ("/worksheet", "worksheet"),
    ("/chartsheet", "chartsheet"),
    ("/dialogsheet", "dialogsheet"),
    ("macrosheet", "macrosheet"),  # Microsoft's xlMacrosheet and xlIntlMacrosheet
)
_STATES: dict[str, SheetState] = {
    "visible": "visible",
    "hidden": "hidden",
    "veryHidden": "very_hidden",
}
_INDEX = re.compile(r"[0-9]{1,9}")  # an unsignedInt; longer is no index any list has
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


def listed_sheets(archive: zipfile.ZipFile) -> tuple[list[ListedSheet], int | None]:
    """The workbook part's sheets in order, and its first ``workbookView``'s ``activeTab``.

    ``activeTab`` is an index into this list; ``None`` when absent or not an index.
    """
    workbook = workbook_part(archive)
    rels = {rid: (kind, target) for rid, kind, target, external in related(archive, workbook)}
    root = ET.fromstring(archive.read(workbook))
    found = []
    for sheet in root.iter(f"{_MAIN}sheet"):
        kind, target = rels.get(sheet.get(f"{_REL}id", ""), ("", None))
        found.append(
            ListedSheet(
                sheet.get("name", ""),
                kind_of(kind),
                _STATES.get(sheet.get("state", "visible"), "visible"),
                target,
            )
        )
    view = root.find(f"{_MAIN}bookViews/{_MAIN}workbookView")
    active = _index(view.get("activeTab")) if view is not None else None
    return found, active


def kind_of(relationship_type: str) -> SheetKind:
    """A sheet's kind from its relationship type; a worksheet when the type names none."""
    lowered = relationship_type.lower()
    return next((kind for suffix, kind in _KINDS if suffix in lowered), "worksheet")


def tab_color(archive: zipfile.ZipFile, part: str | None) -> StatedColor | None:
    """The ``sheetPr/tabColor`` at the head of a sheet part, as stated; ``None`` without one.

    The parse ends at the end of ``sheetPr``, or at the first other child of the root.
    """
    if part is None or part not in archive.namelist():
        return None
    depth = 0
    with archive.open(part) as handle:
        try:
            for event, element in ET.iterparse(handle, ("start", "end")):
                if event == "start":
                    depth += 1
                    if depth == 2 and element.tag != f"{_MAIN}sheetPr":
                        return None
                    continue
                depth -= 1
                if depth == 1:  # the end of sheetPr
                    color = element.find(f"{_MAIN}tabColor")
                    return stated_color(color) if color is not None else None
        except ET.ParseError:
            return None
    return None


def stated_color(element: ET.Element) -> StatedColor:
    """A ``CT_Color``'s attributes, taken in openpyxl's order: indexed, theme, auto, rgb.

    The schema allows one of them; which one Excel takes when several are given is unverified,
    so the order is that of every other color sheetpeek reads through openpyxl.
    """
    tint = _tint(element.get("tint"))
    indexed = element.get("indexed")
    if indexed is not None:
        return StatedColor("indexed", indexed=_index(indexed), tint=tint)
    theme = element.get("theme")
    if theme is not None:
        return StatedColor("theme", theme=_index(theme), tint=tint)
    if element.get("auto") is not None:
        return StatedColor("auto")
    rgb = element.get("rgb")
    if rgb is not None:
        return StatedColor("rgb", rgb=rgb.strip(), tint=tint)
    return StatedColor("auto")


def _index(text: str | None) -> int | None:
    """An ``xsd:unsignedInt`` small enough to index a list, or ``None``."""
    if text is None or not _INDEX.fullmatch(text.strip()):
        return None
    return int(text.strip())


def _tint(text: str | None) -> float:
    """A tint from -1 to 1; text that is not a finite number is 0.

    One past either end is clamped: a huge one would make the lightness not a number.
    """
    if text is None:
        return 0.0
    try:
        value = float(text)
    except ValueError:
        return 0.0
    if not math.isfinite(value):
        return 0.0
    return max(-1.0, min(1.0, value))


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
