"""Hyperlinks: the model's record, and the checks every output that links shares.

A link adds no look of its own: the cell's format draws it, as in Excel (``ARCHITECTURE.md``
§4.10). What a link may point to is decided here once, so that every output refuses the same
targets: an allow-list of URL schemes, and same-document anchors for places on the rendered sheet.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal, TypeAlias
from urllib.parse import quote, urlsplit

LinkSource: TypeAlias = Literal["element", "formula"]

ALLOWED_SCHEMES = frozenset({"http", "https", "mailto"})
"""The URL schemes an output links to; any other target is drawn as plain text and reported."""

LOCAL_TARGETS = frozenset({"file", "relative", "path"})
"""Targets Excel follows that point into the viewer's own files: ``file:`` URLs, relative paths
(UNC names included) and drive-letter paths. Linked only when a backend's ``local_links`` is set;
otherwise reported by each output (``link.local``), not refused."""

# What a browser drops from a URL before it parses it (the WHATWG URL standard): C0 controls and
# spaces at either end, tabs and line breaks anywhere. "java\tscript:" is "javascript:" to it.
_EDGES = "".join(chr(code) for code in range(0x21))
_TAB_OR_BREAK = re.compile("[\t\n\r]")
_CONTROL = re.compile("[\x00-\x1f\x7f-\x9f]")


@dataclass(frozen=True, slots=True)
class Hyperlink:
    """A cell's hyperlink, as the file states it.

    The text a linked cell shows is its cached value: ``display`` is kept, but Excel does not
    draw it, and neither does sheetpeek.
    """

    target: str | None = None
    """The external target, from the ``<hyperlink>``'s relationship or the literal first
    argument of ``HYPERLINK()``: a URL, a file path, or anything else the file says."""
    location: str | None = None
    """A place in the workbook (``Sheet2!A1``, ``'My sheet'!B3``, a defined name), or, with a
    ``target``, the part after ``#``."""
    tooltip: str | None = None
    """The text Excel shows when the pointer rests on the cell."""
    display: str | None = None
    """The ``display`` attribute; never drawn."""
    source: LinkSource = "element"
    """``"element"`` for a ``<hyperlink>`` element, ``"formula"`` for a ``HYPERLINK()`` call."""


def check_target(target: str, local: bool = False) -> tuple[str | None, str]:
    """A target as an output may link to it, or ``None`` when it is refused; and its scheme.

    The scheme is compared after what a browser drops is dropped, lower-cased. The URL returned
    has the controls a terminal or a browser would act on percent-encoded. A refused target's
    scheme is ``"relative"`` (a path, a UNC name, an anchor), ``"path"`` (a drive letter) or
    ``"unparseable"`` when it has none. With ``local``, the :data:`LOCAL_TARGETS` are allowed
    too.
    """
    cleaned = _TAB_OR_BREAK.sub("", target.strip(_EDGES))
    try:
        scheme = urlsplit(cleaned).scheme.lower()
    except ValueError:
        return None, "unparseable"
    label = scheme if scheme else "relative"
    if len(label) == 1:
        label = "path"
    if scheme in ALLOWED_SCHEMES or (local and label in LOCAL_TARGETS):
        return _CONTROL.sub(lambda m: quote(m[0]), cleaned), label
    if not scheme:
        return None, "relative"
    return None, "path" if len(scheme) == 1 else scheme


def external_url(link: Hyperlink, local: bool = False) -> str | None:
    """The URL an output links to for an external target, ``#location`` added; else ``None``."""
    if link.target is None:
        return None
    url, _ = check_target(link.target, local)
    if url is None:
        return None
    if link.location:
        url = f"{url}#{_CONTROL.sub(lambda m: quote(m[0]), link.location)}"
    return url


def location_cell(location: str, sheet: str) -> tuple[int, int] | None:
    """The cell a ``location`` points to on ``sheet``, or ``None`` when it points elsewhere.

    ``A1`` and ``Sheet!A1`` name a cell (``$`` allowed), ``'It''s'!B2:C3`` a range whose first
    cell is taken; sheet names compare ignoring case, as Excel's do. A defined name, another
    sheet's cell or text that is neither gives ``None``. A reference without a sheet is taken as
    the link's own sheet (unverified).
    """
    place = location_place(location, sheet)
    if place is None or place[0].casefold() != sheet.casefold():
        return None
    return place[1], place[2]


def location_place(location: str, sheet: str) -> tuple[str, int, int] | None:
    """The sheet name and cell a ``location`` names, the name as written; ``None`` for no cell.

    As :func:`location_cell`, for any sheet: a reference without a sheet names ``sheet``.
    """
    text = location.strip()
    if text.startswith("'"):
        name, rest = _quoted(text)
        if name is None or not rest.startswith("!"):
            return None
        reference = rest[1:]
    elif "!" in text:
        name, _, reference = text.partition("!")
    else:
        name, reference = sheet, text
    from sheetpeek._model import CellRange  # the model imports this module

    try:
        cells = CellRange.parse(reference)
    except ValueError:
        return None
    return name, cells.min_row, cells.min_col


def _quoted(text: str) -> tuple[str | None, str]:
    """A quoted sheet name at the start of ``text`` (``''`` is one quote), and what follows."""
    name = []
    i = 1
    while i < len(text):
        if text[i] == "'":
            if text[i + 1 : i + 2] == "'":
                name.append("'")
                i += 2
                continue
            return "".join(name), text[i + 1 :]
        name.append(text[i])
        i += 1
    return None, ""


LinkKind: TypeAlias = Literal["url", "local", "internal", "elsewhere", "refused"]


def classify(
    link: Hyperlink, sheet: str, shown: Callable[[int, int], bool], local: bool = False
) -> tuple[LinkKind, str | None, tuple[int, int] | None]:
    """What an output can make of a link on ``sheet``, whose shown cells ``shown`` tells.

    ``local`` is a target in the viewer's own files, not linked unless ``local`` is set.
    Returns the kind, the URL of a ``url`` link, and the cell an ``internal`` link points to.
    ``elsewhere`` is a place in the workbook that is not on the rendered sheet: another sheet, a
    defined name, or a cell outside what is shown.
    """
    if link.target is not None:
        url = external_url(link, local)
        if url is not None:
            return "url", url, None
        label = check_target(link.target)[1]
        return ("local" if label in LOCAL_TARGETS else "refused"), None, None
    cell = location_cell(link.location or "", sheet)
    if cell is not None and shown(*cell):
        return "internal", None, cell
    return "elsewhere", None, None
