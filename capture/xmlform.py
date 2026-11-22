"""openpyxl's XML in one form, whichever writer it used.

openpyxl writes through lxml when lxml is installed, which closes an empty element as ``<v/>`` or
``<v></v>``; the standard library writes ``<v />``. The generators edit the XML by its text, and
the edits name the standard library's form.
"""

from __future__ import annotations

import re


def standard_form(text: str) -> str:
    """``text`` with every empty element closed as the standard library closes it."""
    return re.sub(r"(?<! )/>", " />", text).replace("<v></v>", "<v />")


def serialized(part: str) -> bool:
    """Whether openpyxl serializes ``part`` (its writer then decides the form); the theme it
    writes as a fixed text, the same with either writer, which stays as it is."""
    return part.endswith((".xml", ".rels")) and not part.startswith("xl/theme/")
