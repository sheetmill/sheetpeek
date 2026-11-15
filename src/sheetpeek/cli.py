"""The ``sheetpeek`` command: ``sheetpeek view book.xlsx`` and ``sheetpeek serve FOLDER``."""

from __future__ import annotations

import argparse
import dataclasses
import importlib.util
import sys
from collections.abc import Sequence
from typing import TextIO

from sheetpeek._model import one_line
from sheetpeek._notices import Notice
from sheetpeek._package import SERVE_LIMITS, PackageError, ReadLimits
from sheetpeek._read import Book
from sheetpeek._read import open as open_book
from sheetpeek._sheets import SheetInfo
from sheetpeek._view_model import MAX_ZOOM, MIN_ZOOM
from sheetpeek.backends import (
    Backend,
    HTMLBackend,
    RichBackend,
    TextBackend,
    backend_names,
    get_backend,
    write_blocks,
)
from sheetpeek.backends.html._workbook import render_workbook


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
        description="Show one sheet of a workbook, or with --all-sheets every sheet as one HTML "
        "page. On a terminal the output is styled with rich when it is installed; otherwise, and "
        "when piped, it is plain text.",
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
    view.add_argument(
        "--print-area",
        action="store_true",
        help="show the print area, each of its ranges as a block, instead of the used range",
    )
    view.add_argument(
        "--zoom",
        type=int,
        metavar="PERCENT",
        help="zoom from 10 to 400 (-t html only; default: the zoom saved with the sheet)",
    )
    view.add_argument(
        "--list",
        action="store_true",
        help="list the sheets and exit; a sheet that is not a visible worksheet says what it is",
    )
    view.add_argument(
        "--all-sheets",
        action="store_true",
        help="every sheet on one HTML page, with Excel's tabs (-t html only)",
    )
    view.add_argument(
        "--hidden-sheets",
        action="store_true",
        help="with --all-sheets, hidden and very hidden sheets too",
    )
    view.add_argument(
        "--notices",
        choices=("summary", "full", "none"),
        default="summary",
        help="what to say on standard error about content not shown: one line per kind "
        "(summary, the default), with the cells (full), or nothing (none)",
    )
    view.add_argument(
        "--strict", action="store_true", help="exit with status 3 when anything is not shown"
    )
    view.add_argument(
        "--suppress",
        action="append",
        default=[],
        metavar="CODE",
        help="do not report CODE or the codes under it, such as 'font' (repeatable)",
    )
    _add_limits(view, ReadLimits())
    serve = commands.add_parser(
        "serve",
        help="serve a folder of workbooks as HTML",
        description="Serve the .xlsx files under a folder as HTML pages, read-only. Binds to "
        "127.0.0.1 unless --host says otherwise, and reads only .xlsx files inside the folder. "
        "There is no authentication: with --host, anyone who can reach the address can read "
        "every workbook in the folder. Not for production: it is built on Python's http.server, "
        "which Python's documentation does not recommend for production. For a production "
        "service, call the Python API from a web framework and pass limits=, since the library's "
        "defaults are not memory-safe for untrusted files.",
    )
    serve.add_argument("folder", help="the folder of workbooks")
    serve.add_argument("--host", default="127.0.0.1", help="address to bind (default: 127.0.0.1)")
    serve.add_argument("-p", "--port", type=int, default=8000, help="port (default: 8000)")
    _add_limits(serve, SERVE_LIMITS)
    return parser


def _add_limits(command: argparse.ArgumentParser, defaults: ReadLimits) -> None:
    """The options behind ``ReadLimits``: ``--max-parts`` for ``max_parts``, and so on."""
    helps = {
        "max_parts": "parts (files) in the package",
        "max_part_bytes": "decompressed bytes of any one part",
        "max_total_bytes": "decompressed bytes of all parts",
        "max_cells": "rows times columns of a sheet read or laid out",
    }
    for field in dataclasses.fields(ReadLimits):
        default = getattr(defaults, field.name)
        command.add_argument(
            "--" + field.name.replace("_", "-"),
            dest=field.name,
            type=_count,
            default=default,
            metavar="N",
            help=f"refuse more than this many {helps[field.name]} (default: {default})",
        )


def _count(text: str) -> int:
    value = int(text)
    if value < 0:
        raise argparse.ArgumentTypeError(f"not a count: {text}")
    return value


def _limits(args: argparse.Namespace) -> ReadLimits:
    values = {field.name: getattr(args, field.name) for field in dataclasses.fields(ReadLimits)}
    return ReadLimits(**values)


def _serve(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    from sheetpeek._serve import make_server

    try:
        server = make_server(args.folder, args.host, args.port, limits=_limits(args))
    except (NotADirectoryError, OSError) as error:
        _say(err, f"sheetpeek: {error}")
        return 1
    host, port = str(server.server_address[0]), server.server_address[1]
    _say(out, f"sheetpeek: serving {args.folder} at http://{host}:{port}/ (Ctrl-C to stop)")
    out.flush()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


def _view(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    html_only = [o for o, given in (("--zoom", args.zoom is not None),) if given]
    if html_only and args.to != "html" and not (args.all_sheets and args.to is None):
        _say(err, f"sheetpeek: {html_only[0]} draws HTML only (use -t html)")
        return 2
    if args.zoom is not None and not MIN_ZOOM <= args.zoom <= MAX_ZOOM:
        _say(err, f"sheetpeek: --zoom must be from {MIN_ZOOM} to {MAX_ZOOM}")
        return 2
    if args.all_sheets:
        refused = _all_sheets_refused(args)
        if refused:
            _say(err, f"sheetpeek: --all-sheets {refused}")
            return 2
    elif args.print_area and args.ref is not None:
        _say(err, "sheetpeek: --print-area shows the print area; it takes no --ref")
        return 2
    elif args.hidden_sheets:
        _say(err, "sheetpeek: --hidden-sheets goes with --all-sheets (-s names a hidden sheet)")
        return 2
    try:
        with open_book(args.file, limits=_limits(args)) as book:
            if args.list:
                for info in book.sheets:
                    _say(out, _listed(info))
                return 0
            if args.all_sheets:
                return _all_sheets(args, book, out, err)
            if args.print_area:
                sheets, problem = book._print_area_sheets(args.sheet)
                sheets = sheets or (book.sheet(args.sheet),)
            else:
                sheets, problem = (book.sheet(args.sheet, ref=args.ref),), None
    except FileNotFoundError:
        _say(err, f"sheetpeek: no such file: {args.file}")
        return 1
    except PackageError as error:
        _say(err, f"sheetpeek: cannot read {args.file}: {error}")
        return 1
    except (KeyError, ValueError) as error:
        message = error.args[0] if error.args else str(error)
        _say(err, f"sheetpeek: {message}")
        return 1
    except Exception as error:  # an unreadable file: openpyxl and zipfile raise many kinds
        _say(err, f"sheetpeek: cannot read {args.file}: {error}")
        return 1
    backend = _backend(args, out)
    try:
        blocks = [s.layout(backend.mode, show_hidden=args.hidden) for s in sheets]
    except PackageError as error:
        _say(err, f"sheetpeek: cannot show {args.file}: {error}")
        return 1
    return _report(args, list(write_blocks(backend, blocks, problem, out)), err)


def _report(args: argparse.Namespace, notices: list[Notice], err: TextIO) -> int:
    """Write the notices as ``--notices`` asks; the exit status."""
    found = [n for n in notices if not any(n.matches(c) for c in args.suppress)]
    if args.notices != "none":
        for notice in found:
            _say(err, f"sheetpeek: {notice.code}: {notice.message}")
            if args.notices == "full" and notice.cells:
                more = " ..." if notice.context.get("truncated") else ""
                _say(err, "  cells: " + " ".join(str(r) for r in notice.cells) + more)
    return 3 if args.strict and found else 0


_KIND_LABELS = {
    "chartsheet": "chartsheet, not shown",
    "dialogsheet": "dialog sheet, not shown",
    "macrosheet": "macro sheet",
}
_STATE_LABELS = {"hidden": "hidden", "very_hidden": "very hidden"}


def _listed(info: SheetInfo) -> str:
    """A line of ``--list``: the name, and what the sheet is unless a visible worksheet."""
    notes = [n for n in (_STATE_LABELS.get(info.state), _KIND_LABELS.get(info.kind)) if n]
    return f"{info.name}  ({', '.join(notes)})" if notes else info.name


def _all_sheets_refused(args: argparse.Namespace) -> str:
    """Why ``--all-sheets`` cannot go with the other options given; empty when it can."""
    if args.to not in (None, "html"):
        return f"draws an HTML page; -t {args.to} has no workbook page (use -t html)"
    if args.sheet is not None or args.ref is not None:
        return "draws every sheet whole; it takes no --sheet or --ref"
    if args.print_area:
        return "draws every sheet whole; it takes no --print-area"
    if args.list:
        return "and --list cannot go together"
    return ""


def _all_sheets(args: argparse.Namespace, book: Book, out: TextIO, err: TextIO) -> int:
    found: list[Notice] = []
    page = render_workbook(
        book,
        hidden_sheets=args.hidden_sheets,
        show_hidden=args.hidden,
        on_notice=found.append,
        column_headings=args.column_headings,
        row_headings=args.row_headings,
        zoom=args.zoom,
    )
    assert page is not None
    out.write(page)
    return _report(args, found, err)


def _say(stream: TextIO, line: str) -> None:
    """One line for a person. It can quote a sheet, a part or a file, so it may not be trusted."""
    stream.write(one_line(line) + "\n")


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
        return HTMLBackend(
            column_headings=args.column_headings, row_headings=args.row_headings, zoom=args.zoom
        )
    return get_backend(name)


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
