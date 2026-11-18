"""The xlsx package: what a file must satisfy before it is read, and where its worksheets are.

``check_package`` runs before anything parses the file (ARCHITECTURE.md §5.5): it holds the
archive to ``ReadLimits``, makes every part's stated size true, and refuses a DTD in any part.
Afterwards every read of the package, xlio's included, sees sizes it can rely on.
"""

from __future__ import annotations

import os
import posixpath
import struct
import xml.etree.ElementTree as ET
import xml.parsers.expat
import zipfile
import zlib
from dataclasses import dataclass
from typing import IO, BinaryIO, NoReturn

import xlio

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_OFFICE_DOCUMENT = "/officeDocument"
_GIB = 1 << 30
_MIB = 1 << 20
_COMPOUND_FILE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_LOCAL_HEADER = struct.Struct("<4s22xHH")
_PROLOG_BOUND = _MIB
_CHUNK = 1 << 16
_XLS = "sheetpeek does not read the old .xls format: convert the file to .xlsx"
_XLSB = "sheetpeek does not read the binary .xlsb format: convert the file to .xlsx"


class PackageError(ValueError):
    """The file is not a readable xlsx package, or sheetpeek refuses to read it.

    Not a workbook at all, a part that is damaged or declares a DTD, sizes in the archive that are
    not the parts' sizes, or a compression method or encryption an xlsx package does not use. A
    subclass of ``ValueError``.
    """


class LimitError(PackageError):
    """The workbook, or what was asked of it, exceeds a limit set by :class:`ReadLimits`."""


@dataclass(frozen=True, slots=True)
class ReadLimits:
    """Limits on what reading one workbook may cost: protection against crafted files.

    The defaults refuse only what could hardly be read anyway, and they do not make reading safe
    on a small machine. A service that reads files from people it does not trust should lower
    them to what it will spend on one file. ``sheetpeek serve`` has lower limits of its own.

    Some parts are held in memory while they are read, so the bytes below are not free: as a
    guide, a part of markup read whole can take about 25 times its size, and shared strings about
    7 times.

    Attributes:
        max_parts: Parts (files) in the package.
        max_part_bytes: Decompressed bytes of any one part.
        max_total_bytes: Decompressed bytes of all parts together.
        max_cells: The grid area of one sheet: its rows times its columns, whether the cells hold
            anything or not. It applies to what is read, which is ``ref`` when one is given and
            otherwise the used range as the cells themselves give it (never the file's
            ``<dimension>``), and again to the layout, which adds columns for spilled text.
    """

    max_parts: int = 100_000
    max_part_bytes: int = 4 * _GIB
    max_total_bytes: int = 16 * _GIB
    max_cells: int = 100_000_000


SERVE_LIMITS = ReadLimits(
    max_parts=10_000,
    max_part_bytes=256 * _MIB,
    max_total_bytes=1 * _GIB,
    max_cells=1_048_576,
)
"""What ``sheetpeek serve`` spends on one request: one full column is allowed (§5.5)."""


def open_package(handle: IO[bytes], limits: ReadLimits, name: str | None = None) -> zipfile.ZipFile:
    """The archive of a checked package.

    Args:
        handle: The file, open for reading and seekable.
        limits: What the package may cost.
        name: The file's name, for a better message when it is not an xlsx file.

    Raises:
        PackageError: The file is not a package, or ``check_package`` refuses it.
    """
    head = handle.read(len(_COMPOUND_FILE))
    handle.seek(0)
    suffix = os.path.splitext(name or "")[1].lower()
    if head == _COMPOUND_FILE:
        if suffix == ".xls":
            raise PackageError(_XLS)
        raise PackageError(
            "this is an old .xls workbook or a password-protected one: sheetpeek reads only "
            "unprotected .xlsx files"
        )
    try:
        archive = zipfile.ZipFile(handle)
    except Exception as error:  # zipfile raises several kinds on a damaged archive
        if suffix == ".xls":
            raise PackageError(_XLS) from error
        if suffix == ".xlsb":
            raise PackageError(_XLSB) from error
        raise PackageError(f"not an xlsx file: {error}") from error
    try:
        if "xl/workbook.bin" in archive.NameToInfo:
            raise PackageError(_XLSB)
        check_package(archive, handle, limits)
    except BaseException:
        archive.close()
        raise
    return archive


def open_workbook(handle: BinaryIO, limits: ReadLimits) -> xlio.Workbook:
    """The package, checked by ``check_package`` first, opened by xlio.

    ``xlio.open`` checks the archive's directory and refuses the rest of what Excel refuses as it
    reads each part; ``check_package`` has already inflated every part (sizes, DTDs), which
    ``xlio.check`` would do only together with reading every sheet whole. xlio's refusals become
    sheetpeek's own errors, a refusal
    worded as sheetpeek words a workbook it cannot read.

    Raises:
        LimitError: The package exceeds ``limits``.
        PackageError: xlio refuses the package.
    """
    handle.seek(0)
    try:
        return xlio.open(
            handle,
            # The package's total is held to max_total_bytes by check_package; xlio counts the
            # bytes of every read it makes, a part read twice twice, so it keeps its own total.
            limits=xlio.ReadLimits(
                max_parts=limits.max_parts, max_part_bytes=limits.max_part_bytes
            ),
        )
    except xlio.LimitError as error:
        raise LimitError(str(error)) from None
    except xlio.PackageError as error:
        raise PackageError(f"cannot read this workbook: {error}") from None


def check_package(archive: zipfile.ZipFile, handle: IO[bytes], limits: ReadLimits) -> None:
    """Refuse a package before anything parses it.

    First what the central directory says, with nothing decompressed. Then every entry, whatever
    its name, is inflated once as a stream: it must produce exactly the size the directory states,
    because ``ZipFile.read`` decompresses far past a smaller stated size before cutting the result
    to it, and it must not declare a DTD.

    Raises:
        LimitError: The package exceeds ``limits``.
        PackageError: A part is refused.
    """
    infos = archive.infolist()
    if len(infos) > limits.max_parts:
        raise LimitError(f"{len(infos)} parts, more than the limit of {limits.max_parts}")
    total = 0
    names = set()
    for info in infos:
        if info.filename in names:
            raise PackageError(f"{info.filename}: the package names this part twice")
        names.add(info.filename)
        if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            raise PackageError(
                f"{info.filename}: compression method {info.compress_type}, which an xlsx "
                "package does not use"
            )
        if info.flag_bits & 1:
            raise PackageError(f"{info.filename}: encrypted")
        if info.file_size > limits.max_part_bytes:
            raise LimitError(
                f"{info.filename}: {info.file_size} bytes, more than the limit of "
                f"{limits.max_part_bytes}"
            )
        total += info.file_size
        if total > limits.max_total_bytes:
            raise LimitError(
                f"more than the limit of {limits.max_total_bytes} decompressed bytes in all"
            )
    for info in infos:
        _check_part(handle, info)


def _check_part(handle: IO[bytes], info: zipfile.ZipInfo) -> None:
    """Inflate one entry as a stream, as far as its stated size and no further."""
    handle.seek(info.header_offset)
    header = handle.read(_LOCAL_HEADER.size)
    if len(header) != _LOCAL_HEADER.size or header[:4] != b"PK\x03\x04":
        raise PackageError(f"{info.filename}: no local header where the directory says")
    _, name_length, extra_length = _LOCAL_HEADER.unpack(header)
    handle.seek(name_length + extra_length, os.SEEK_CUR)
    prolog: _Prolog | None = _Prolog(info.filename)
    produced = 0
    left = info.compress_size
    inflate = zlib.decompressobj(-15) if info.compress_type == zipfile.ZIP_DEFLATED else None
    pending = b""
    finished = inflate is None
    while True:
        if left and not pending:
            pending = handle.read(min(left, _CHUNK))
            if not pending:
                raise PackageError(f"{info.filename}: the archive ends inside this part")
            left -= len(pending)
        if inflate is None:
            data, pending = pending, b""
        else:
            # With its input used up, zlib can still hold output past max_length: empty input
            # asks for it, so the part ends only when a call produces nothing.
            try:
                data = inflate.decompress(pending, _CHUNK)
            except zlib.error as error:
                raise PackageError(f"{info.filename}: damaged: {error}") from error
            pending = inflate.unconsumed_tail
        produced += len(data)
        if produced > info.file_size:
            _false_size(info)
        if prolog is not None and data and prolog.feed(data):
            prolog = None
        if inflate is not None and inflate.eof:
            finished = True
            break
        if not (data or pending or left):
            break
    if produced != info.file_size or not (finished or info.compress_size == 0):
        _false_size(info)
    if prolog is not None:
        prolog.finish()


def _false_size(info: zipfile.ZipInfo) -> NoReturn:
    raise PackageError(
        f"{info.filename}: the archive states {info.file_size} bytes, which is not this part's size"
    )


class _RootReachedError(Exception):
    pass


class _DoctypeFoundError(Exception):
    pass


def _root_reached(*_: object) -> None:
    raise _RootReachedError


def _doctype_found(*_: object) -> None:
    raise _DoctypeFoundError


class _Prolog:
    """A part's first bytes through expat, up to its root element.

    expat reads the prolog as the parsers that come later will, whatever the encoding, however
    long the comments before the DTD, and whatever the part is called. A DTD can only come before
    the root, so nothing past it is looked at. A part expat cannot read at all (an image, say) is
    not XML to any parser that would expand an entity, and passes.
    """

    def __init__(self, name: str) -> None:
        self.name = name
        self.fed = 0
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.StartDoctypeDeclHandler = _doctype_found
        self.parser.StartElementHandler = _root_reached

    def feed(self, data: bytes) -> bool:
        """Whether the part is settled; raises if it is refused."""
        self.fed += len(data)
        settled = self._parse(data, final=False)
        if not settled and self.fed >= _PROLOG_BOUND:
            raise PackageError(f"{self.name}: no root element in the first {_PROLOG_BOUND} bytes")
        return settled

    def finish(self) -> None:
        """The part ended before the root: an empty part, or not XML."""
        self._parse(b"", final=True)

    def _parse(self, data: bytes, *, final: bool) -> bool:
        try:
            self.parser.Parse(data, final)
        except _RootReachedError:
            return True
        except xml.parsers.expat.ExpatError:
            return True
        except _DoctypeFoundError:
            raise PackageError(
                f"{self.name}: declares a DTD, which a part of an xlsx package never does"
            ) from None
        except (ValueError, LookupError) as error:  # a declared encoding expat cannot read
            raise PackageError(f"{self.name}: {error}") from error
        return False


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
