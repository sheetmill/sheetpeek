"""The ``sheetpeek`` command: ``sheetpeek view book.xlsx`` and ``sheetpeek serve FOLDER``."""

from __future__ import annotations

import argparse
import importlib.util
import sys
from collections.abc import Sequence
from typing import TextIO

from sheetpeek._read import open as open_book
from sheetpeek.backends import (
    Backend,
    HTMLBackend,
    RichBackend,
    TextBackend,
    backend_names,
    get_backend,
)


def main(argv: Sequence[str] | None = None) -> int:
    """Run the command line; returns the exit status."""
    parser = _parser()
    args = parser.parse_args(argv)
    if args.command == "serve":
        return _serve(args, sys.stdout, sys.stderr)
    if args.command != "view":  # pragma: no cover - argparse requires a command
        parser.print_help()
        return 2
    return _view(args, sys.stdout, sys.stderr)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sheetpeek", description="Show Excel workbooks the way Excel shows them."
    )
    commands = parser.add_subparsers(dest="command", required=True)
    view = commands.add_parser(
        "view",
        help="show one sheet",
        description="Show one sheet of a workbook. On a terminal the output is styled with rich "
        "when it is installed; otherwise, and when piped, it is plain text.",
    )
    view.add_argument("file", help="the .xlsx workbook")
    view.add_argument("-s", "--sheet", help="sheet name (default: the sheet Excel opens on)")
    view.add_argument("-r", "--ref", help="range to show, such as A1:H40 (default: used range)")
    view.add_argument(
        "-t", "--to", choices=backend_names(), help="output format (default: rich or text)"
    )
    view.add_argument("--column-headings", action="store_true", help="show column letters")
    view.add_argument("--row-headings", action="store_true", help="show row numbers")
    view.add_argument("--hidden", action="store_true", help="show hidden rows and columns")
    view.add_argument("--list", action="store_true", help="list the sheets and exit")
    serve = commands.add_parser(
        "serve",
        help="serve a folder of workbooks as HTML",
        description="Serve the .xlsx files under a folder as HTML pages, read-only. Binds to "
        "127.0.0.1 unless --host says otherwise, and reads only .xlsx files inside the folder.",
    )
    serve.add_argument("folder", help="the folder of workbooks")
    serve.add_argument("--host", default="127.0.0.1", help="address to bind (default: 127.0.0.1)")
    serve.add_argument("-p", "--port", type=int, default=8000, help="port (default: 8000)")
    return parser


def _serve(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    from sheetpeek._serve import make_server

    try:
        server = make_server(args.folder, args.host, args.port)
    except (NotADirectoryError, OSError) as error:
        err.write(f"sheetpeek: {error}\n")
        return 1
    host, port = str(server.server_address[0]), server.server_address[1]
    out.write(f"sheetpeek: serving {args.folder} at http://{host}:{port}/ (Ctrl-C to stop)\n")
    out.flush()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


def _view(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    try:
        with open_book(args.file) as book:
            if args.list:
                for name in book.sheet_names:
                    out.write(name + "\n")
                return 0
            sheet = book.sheet(args.sheet, ref=args.ref)
    except FileNotFoundError:
        err.write(f"sheetpeek: no such file: {args.file}\n")
        return 1
    except (KeyError, ValueError) as error:
        message = error.args[0] if error.args else str(error)
        err.write(f"sheetpeek: {message}\n")
        return 1
    except Exception as error:  # an unreadable file: openpyxl and zipfile raise many kinds
        err.write(f"sheetpeek: cannot read {args.file}: {error}\n")
        return 1
    backend = _backend(args, out)
    rendered = sheet.layout(backend.mode, show_hidden=args.hidden)
    for warning in rendered.warnings:
        err.write(f"sheetpeek: warning: {warning}\n")
    backend.render(rendered, out)
    return 0


def _backend(args: argparse.Namespace, out: TextIO) -> Backend:
    name = args.to
    if name is None:
        terminal = hasattr(out, "isatty") and out.isatty()
        name = "rich" if terminal and importlib.util.find_spec("rich") else "text"
    if name == "rich":
        return RichBackend(column_headings=args.column_headings, row_headings=args.row_headings)
    if name == "text":
        return TextBackend(column_headings=args.column_headings, row_headings=args.row_headings)
    if name == "html":
        return HTMLBackend(column_headings=args.column_headings, row_headings=args.row_headings)
    return get_backend(name)


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
