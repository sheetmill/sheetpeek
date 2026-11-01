"""Find a worksheet's part inside the xlsx package, from the workbook and its relationships."""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_OFFICE_DOCUMENT = "/officeDocument"


def relationships(archive: zipfile.ZipFile, part: str) -> dict[str, tuple[str, bool]]:
    """Relationship id → (target, external) for a part; targets are resolved to part names.

    A part without a relationships file has no relationships.
    """
    folder, name = posixpath.split(part)
    rels = posixpath.join(folder, "_rels", f"{name}.rels")
    try:
        root = ET.fromstring(archive.read(rels))
    except KeyError:
        return {}
    result = {}
    for rel in root.iter(f"{_PKG_REL}Relationship"):
        target = rel.get("Target", "")
        external = rel.get("TargetMode") == "External"
        if not external:
            target = _resolve(folder, target)
        result[rel.get("Id", "")] = (target, external)
    return result


def worksheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    """Sheet name → worksheet part name, in workbook order (chartsheets included)."""
    workbook = _workbook_part(archive)
    rels = relationships(archive, workbook)
    root = ET.fromstring(archive.read(workbook))
    parts = {}
    for sheet in root.iter(f"{_MAIN}sheet"):
        target = rels.get(sheet.get(f"{_REL}id", ""))
        if target is not None:
            parts[sheet.get("name", "")] = target[0]
    return parts


def _workbook_part(archive: zipfile.ZipFile) -> str:
    root = ET.fromstring(archive.read("_rels/.rels"))
    for rel in root.iter(f"{_PKG_REL}Relationship"):
        if rel.get("Type", "").endswith(_OFFICE_DOCUMENT):
            return _resolve("", rel.get("Target", ""))
    raise KeyError("the package has no workbook part")


def _resolve(folder: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(folder, target))
