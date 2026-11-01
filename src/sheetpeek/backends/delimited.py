"""Delimited text (CSV, TSV): each cell's displayed text, without widths."""

from __future__ import annotations

import csv
from typing import TextIO

from sheetpeek._rendered import Mode, RenderedSheet


class DelimitedBackend:
    r"""Rows of displayed text in ``none`` mode, written with :mod:`csv`.

    A merged range's text is in its anchor cell; the other cells of the range are empty.

    Args:
        delimiter: ``","`` for CSV, ``"\t"`` for TSV, or another single character.
    """

    mode: Mode = "none"

    def __init__(self, delimiter: str = ",") -> None:
        self.delimiter = delimiter
        self.name = {",": "csv", "\t": "tsv"}.get(delimiter, "delimited")

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write one record per visible row."""
        writer = csv.writer(out, delimiter=self.delimiter, lineterminator="\n")
        for row in sheet.rows:
            writer.writerow([cell.lines[0] for cell in row.cells])
