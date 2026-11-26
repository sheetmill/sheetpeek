"""``sheetpeek serve``: a small read-only web server for a folder of workbooks.

Standard library only. ``/`` lists the ``.xlsx`` files under the folder, ``/view/<path>`` shows a
workbook's default sheet; ``?sheet=`` and ``?ref=`` choose another sheet and a range, and
``?zoom=`` a zoom; ``?print_area=1`` shows the print area. Each sheet is drawn by the HTML
backend. Only regular ``.xlsx`` files inside the folder are read: every requested path is resolved,
symbolic links included, and must stay inside it.

A page can be asked for by anyone who can connect, and through a web page in a browser on this
machine, so what one request may cost is bounded (``SERVE_LIMITS``, two reads at a time), and a
server bound to a loopback address answers only to loopback names, which stops DNS rebinding
(ARCHITECTURE.md §5.5).
"""

from __future__ import annotations

import html
import io
import ipaddress
import os
import re
import sys
import threading
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlsplit

from sheetpeek._model import CellRange, one_line
from sheetpeek._package import SERVE_LIMITS, PackageError, ReadLimits
from sheetpeek._platform import DEFAULT_PLATFORM_NAME
from sheetpeek._platform import resolve as resolve_platform
from sheetpeek._read import open as open_book
from sheetpeek._recalc import RecalcOnLoad, recalc_mode
from sheetpeek.backends import _write_blocks
from sheetpeek.backends.html import CSS, HTMLBackend

_READS_AT_ONCE = 2
_SLOT_WAIT = 2.0  # seconds; a request that gets no slot in time is turned away, not queued
_RETRY_AFTER = "5"
_LOOPBACK_NAMES = frozenset({"localhost", "127.0.0.1", "[::1]"})
_HOST = re.compile(r"(\[[^\]]*\]|[^:\[\]]+)(?::\d*)?")
_PAGE_CSS = """\
body { margin: 0; font: 13px -apple-system, "Segoe UI", sans-serif; background: #ffffff; }
.sp-bar { padding: 8px 12px; border-bottom: 1px solid #d4d4d4; background: #f7f7f7; }
.sp-bar a, .sp-bar span { margin-right: 12px; }
.sp-bar .sp-current { font-weight: 600; margin-right: 12px; }
.sp-sheet { padding: 12px; }
.sp-list { padding: 12px 24px; line-height: 1.8; }
.sp-error { padding: 12px; color: #a00000; }
"""


def workbooks(root: Path) -> list[Path]:
    """The ``.xlsx`` files under ``root``, relative to it, sorted; hidden ones left out."""
    found: list[Path] = []
    for directory, folders, files in os.walk(root):
        folders[:] = sorted(f for f in folders if not f.startswith("."))
        found.extend(
            (Path(directory) / name).relative_to(root)
            for name in files
            if name.lower().endswith(".xlsx") and not name.startswith((".", "~$"))
        )
    return sorted(found, key=lambda p: p.as_posix().lower())


def resolve(root: Path, relative: str) -> Path | None:
    """The workbook a request names, or ``None``.

    Only a regular ``.xlsx`` file inside ``root`` qualifies, after resolving ``..`` and symbolic
    links; a part starting with a dot is refused.
    """
    if not relative or "\0" in relative:
        return None
    parts = Path(relative).parts
    if Path(relative).is_absolute() or any(part.startswith(".") for part in parts):
        return None
    candidate = (root / relative).resolve()
    if not candidate.is_relative_to(root) or candidate.suffix.lower() != ".xlsx":
        return None
    return candidate if candidate.is_file() else None


_KINDS = {"chartsheet": "chartsheet, not shown", "dialogsheet": "dialog sheet, not shown"}


def _page(title: str, body: str) -> bytes:
    return (
        '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{html.escape(title)}</title>\n<style>\n{CSS}{_PAGE_CSS}</style>\n</head>\n"
        f"<body>\n{body}</body>\n</html>\n"
    ).encode()


def _link(relative: Path, sheet: str | None = None) -> str:
    target = "/view/" + quote(relative.as_posix())
    if sheet is not None:
        target += "?sheet=" + quote(sheet)
    return target


class _Handler(BaseHTTPRequestHandler):
    # Set on the subclass made by make_server.
    root: Path
    limits: ReadLimits
    recalc_on_load: RecalcOnLoad
    platform: str
    slots: threading.BoundedSemaphore
    hosts: frozenset[str] | None  # the names answered to, or None to answer to any
    server_version = "sheetpeek"

    def do_HEAD(self) -> None:
        self.respond(head=True)

    def do_GET(self) -> None:
        self.respond(head=False)

    def respond(self, *, head: bool) -> None:
        url = urlsplit(self.path)
        path = unquote(url.path)
        query = parse_qs(url.query)
        if not self.host_allowed():
            status, body = HTTPStatus.FORBIDDEN, self.error("This host name is not served here.")
        elif path == "/":
            status, body = HTTPStatus.OK, self.index()
        elif path.startswith("/view/"):
            status, body = self.view(path.removeprefix("/view/"), query)
        else:
            status, body = HTTPStatus.NOT_FOUND, self.error("No such page.")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "default-src 'none'; style-src 'unsafe-inline'")
        if status == HTTPStatus.SERVICE_UNAVAILABLE:
            self.send_header("Retry-After", _RETRY_AFTER)
        self.end_headers()
        if not head:
            self.wfile.write(body)

    def host_allowed(self) -> bool:
        """Whether the request's Host names this server; the port is not compared.

        A tunnel to this machine brings its own port, and a rebound name is refused by its name.
        """
        if self.hosts is None:
            return True
        given = self.headers.get_all("Host", [])
        if len(given) != 1:
            return False
        match = _HOST.fullmatch(given[0].strip().lower())
        return match is not None and match[1] in self.hosts

    def index(self) -> bytes:
        items = "".join(
            f'<div><a href="{_link(p)}">{html.escape(p.as_posix())}</a></div>\n'
            for p in workbooks(self.root)
        )
        title = html.escape(self.root.name or "/")
        body = f'<div class="sp-bar"><span class="sp-current">{title}</span></div>\n'
        body += f'<div class="sp-list">{items or "No .xlsx files here."}</div>\n'
        return _page(f"sheetpeek: {self.root.name}", body)

    def view(self, relative: str, query: dict[str, list[str]]) -> tuple[HTTPStatus, bytes]:
        path = resolve(self.root, relative)
        if path is None:
            return HTTPStatus.NOT_FOUND, self.error("No such workbook.")
        if not self.slots.acquire(timeout=_SLOT_WAIT):
            message = "Too many workbooks are being read. Try again in a moment."
            return HTTPStatus.SERVICE_UNAVAILABLE, self.error(message)
        try:
            return self.sheet_page(path, query)
        finally:
            self.slots.release()

    def sheet_page(self, path: Path, query: dict[str, list[str]]) -> tuple[HTTPStatus, bytes]:
        name = path.relative_to(self.root)
        sheet_name = query.get("sheet", [None])[0]
        ref = query.get("ref", [None])[0]
        zoom_text = query.get("zoom", [None])[0]
        try:
            zoom = (
                int(zoom_text)
                if zoom_text and zoom_text.isascii() and zoom_text.isdigit()
                else None
            )
            if zoom_text is not None and zoom is None:
                raise ValueError("zoom must be a whole number from 10 to 400")
            backend = HTMLBackend(
                fragment=True, column_headings=True, row_headings=True, notices=True, zoom=zoom
            )
            with open_book(
                path,
                limits=self.limits,
                recalc_on_load=self.recalc_on_load,
                platform=self.platform,
            ) as book:
                infos = book.sheets
                if query.get("print_area", ["0"])[0] == "1":
                    if ref:
                        raise ValueError("print_area shows the print area; it takes no ref")
                    sheets, problem = book._print_area_sheets(sheet_name)
                    sheets = sheets or (book.sheet(sheet_name),)
                else:
                    sheets = (book.sheet(sheet_name, ref=CellRange.parse(ref) if ref else None),)
                    problem = None
            blocks = [s.layout("excel") for s in sheets]
            sheet = blocks[0].sheet
        except PackageError as error:
            return HTTPStatus.UNPROCESSABLE_ENTITY, self.error(f"Cannot read this file: {error}")
        except (KeyError, ValueError) as error:
            message = error.args[0] if error.args else str(error)
            return HTTPStatus.NOT_FOUND, self.error(str(message))
        except Exception as error:  # nothing about it is for the client: it can name local paths
            sys.stderr.write(one_line(f"sheetpeek: {name}: {error}") + "\n")
            return HTTPStatus.UNPROCESSABLE_ENTITY, self.error("Cannot read this file.")
        out = io.StringIO()
        _write_blocks(backend, blocks, problem, out)
        # The page carries the base stylesheet; the fragment keeps the sheet's own rules (its
        # fill, border, outline and freeze classes).
        fragment = out.getvalue().replace(f"<style>\n{CSS}", "<style>\n", 1)
        tabs = "".join(
            f'<span class="sp-current">{html.escape(i.name)}</span>'
            if i.name == sheet.name
            else f'<a href="{_link(name, i.name)}">{html.escape(i.name)}</a>'
            if i.readable
            else f"<span>{html.escape(i.name)} ({_KINDS[i.kind]})</span>"
            for i in infos
        )
        body = (
            f'<div class="sp-bar"><a href="/">All workbooks</a>'
            f'<span class="sp-current">{html.escape(name.as_posix())}</span> {tabs}</div>\n'
            f'<div class="sp-sheet">{fragment}</div>\n'
        )
        return HTTPStatus.OK, _page(f"{sheet.name} - {name.name}", body)

    def error(self, message: str) -> bytes:
        bar = '<div class="sp-bar"><a href="/">All workbooks</a></div>\n'
        return _page("sheetpeek", f'{bar}<div class="sp-error">{html.escape(message)}</div>\n')

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002 - http.server's name
        # http.server's own log_message escapes control characters; a request line is anyone's.
        sys.stderr.write(one_line(f"sheetpeek: {self.address_string()} {format % args}") + "\n")


def make_server(
    root: str | os.PathLike[str],
    host: str = "127.0.0.1",
    port: int = 8000,
    *,
    limits: ReadLimits = SERVE_LIMITS,
    recalc_on_load: RecalcOnLoad = "excel",
    platform: str = DEFAULT_PLATFORM_NAME,
) -> ThreadingHTTPServer:
    """A server for the workbooks under ``root``; call ``serve_forever()`` on it.

    Bound to a loopback address, it answers only to ``localhost``, ``127.0.0.1``, ``[::1]`` and
    the address itself. Bound to anything else, it answers to any name. It has no authentication:
    anyone who can reach the address can read every workbook under ``root``.

    Raises:
        NotADirectoryError: ``root`` is not a folder.
    """
    folder = Path(root).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"not a folder: {root}")

    class Handler(_Handler):
        pass

    Handler.root = folder
    Handler.limits = limits
    Handler.recalc_on_load = recalc_mode(recalc_on_load)
    Handler.platform = resolve_platform(platform).name
    Handler.slots = threading.BoundedSemaphore(_READS_AT_ONCE)
    server = ThreadingHTTPServer((host, port), Handler)
    address = str(server.server_address[0])
    loopback = ipaddress.ip_address(address).is_loopback
    Handler.hosts = _LOOPBACK_NAMES | {address} if loopback else None
    return server
