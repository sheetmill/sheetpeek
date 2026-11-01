"""``sheetpeek serve``: a small read-only web server for a folder of workbooks.

Standard library only. ``/`` lists the ``.xlsx`` files under the folder, ``/view/<path>`` shows
a workbook's default sheet, and ``?sheet=`` and ``?ref=`` choose another sheet and a range. Each
sheet is drawn by the HTML backend. Only regular ``.xlsx`` files inside the folder are read:
every requested path is resolved, symbolic links included, and must stay inside it.
"""

from __future__ import annotations

import html
import io
import os
import sys
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlsplit

from sheetpeek._model import CellRange
from sheetpeek._read import open as open_book
from sheetpeek.backends.html import CSS, HTMLBackend

_PAGE_CSS = """\
body { margin: 0; font: 13px -apple-system, "Segoe UI", sans-serif; background: #ffffff; }
.sp-bar { padding: 8px 12px; border-bottom: 1px solid #d4d4d4; background: #f7f7f7; }
.sp-bar a { margin-right: 12px; }
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
    root: Path  # set on the subclass made by make_server
    server_version = "sheetpeek"

    def do_HEAD(self) -> None:
        self.respond(head=True)

    def do_GET(self) -> None:
        self.respond(head=False)

    def respond(self, *, head: bool) -> None:
        url = urlsplit(self.path)
        path = unquote(url.path)
        query = parse_qs(url.query)
        if path == "/":
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
        self.end_headers()
        if not head:
            self.wfile.write(body)

    def index(self) -> bytes:
        items = "".join(
            f'<div><a href="{_link(p)}">{html.escape(p.as_posix())}</a></div>\n'
            for p in workbooks(self.root)
        )
        title = html.escape(str(self.root))
        body = f'<div class="sp-bar"><span class="sp-current">{title}</span></div>\n'
        body += f'<div class="sp-list">{items or "No .xlsx files here."}</div>\n'
        return _page(f"sheetpeek: {self.root.name}", body)

    def view(self, relative: str, query: dict[str, list[str]]) -> tuple[HTTPStatus, bytes]:
        path = resolve(self.root, relative)
        if path is None:
            return HTTPStatus.NOT_FOUND, self.error("No such workbook.")
        name = path.relative_to(self.root)
        sheet_name = query.get("sheet", [None])[0]
        ref = query.get("ref", [None])[0]
        try:
            with open_book(path) as book:
                names = book.sheet_names
                sheet = book.sheet(sheet_name, ref=CellRange.parse(ref) if ref else None)
        except (KeyError, ValueError) as error:
            message = error.args[0] if error.args else str(error)
            return HTTPStatus.NOT_FOUND, self.error(str(message))
        except Exception as error:  # an unreadable file: openpyxl and zipfile raise many kinds
            return HTTPStatus.UNPROCESSABLE_ENTITY, self.error(f"Cannot read this file: {error}")
        out = io.StringIO()
        HTMLBackend(fragment=True, column_headings=True, row_headings=True).render(
            sheet.layout("excel"), out
        )
        fragment = out.getvalue().split("</style>\n", 1)[1]  # the page carries the stylesheet
        tabs = "".join(
            f'<span class="sp-current">{html.escape(n)}</span>'
            if n == sheet.name
            else f'<a href="{_link(name, n)}">{html.escape(n)}</a>'
            for n in names
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
        sys.stderr.write(f"sheetpeek: {self.address_string()} {format % args}\n")


def make_server(
    root: str | os.PathLike[str], host: str = "127.0.0.1", port: int = 8000
) -> ThreadingHTTPServer:
    """A server for the workbooks under ``root``; call ``serve_forever()`` on it.

    Raises:
        NotADirectoryError: ``root`` is not a folder.
    """
    folder = Path(root).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"not a folder: {root}")
    handler = type("Handler", (_Handler,), {"root": folder})
    return ThreadingHTTPServer((host, port), handler)
