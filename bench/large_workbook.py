"""Write a large workbook of a common real shape, for timing and memory at scale.

    python bench/large_workbook.py [DIR] [SCALE]

Writes ``large-SCALE.xlsx`` into DIR (by default a new temporary directory) and prints its
path and shape. At scale 1 it has six sheets, 264,408 rows and 14.8 million cells, 33 MiB: a
sheet of data (3,124 rows of 12 values, strings and numbers under eight formats), two
65,523-row sheets 99 columns wide and two 65,000-row sheets 14 columns wide whose cells are
nearly all styled and empty, with a few values, a sheet of 75 rows, and nine ``SUM`` formulas
with their cached values. Most of its cells use a Times New Roman 10 style, top-aligned with
wrapped text, under an Arial 10 Normal font; there are 76 cell formats and 675 distinct strings.
This is the shape of exported reports where a style was applied to whole columns. SCALE
multiplies every sheet's rows (0.25 and 0.5 measure how time and memory grow).

The XML is streamed into the archive, so writing takes little memory and a few seconds.
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from collections.abc import Iterator
from pathlib import Path

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

STRINGS = 675
FORMATS = (  # id, code
    (164, "0.000"),
    (165, '"$"#,##0'),
    (166, "0.000%"),
    (167, "0.0"),
    (168, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'),
)
FONTS = (  # name, size, bold, italic
    ("Arial", 10, False, False),
    ("Arial", 10, True, False),
    ("Times New Roman", 14, True, False),
    ("Times New Roman", 10, False, False),
    ("Times New Roman", 10, True, False),
    ("Times New Roman", 10, False, True),
    ("Times New Roman", 14, False, False),
    ("Courier New", 10, False, False),
)
NUMBER_FORMATS = (0, 3, 4, 10, 14, 15, 164, 165, 166, 167, 168)
BLANK = 48
"""The style of the empty cells: Times New Roman 10, top-aligned, wrapped."""
DATA_STYLES = (2, 3, 4, 29, 30, 44, 45, 46, 47, 11, 12, 13)
"""The styles of a data row's twelve cells, column by column."""

SHEETS = (  # name, rows, cells a row, every how many rows a value, and the kind
    ("Data", 3124, 12, 1, "data"),
    ("Grid A", 65523, 99, 17, "sparse"),
    ("Notes", 75, 16, 3, "sparse"),
    ("Grid B", 65523, 99, 17, "sparse"),
    ("Log A", 65524, 14, 6, "sparse"),
    ("Log B", 64630, 14, 383, "sparse"),
)


def column(index: int) -> str:
    """The letters of a 1-based column number."""
    letters = ""
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def xf(index: int) -> str:
    """Cell format ``index``: a font, a number format and, every seventh, a bottom border."""
    if index == 0:
        return '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    if index == BLANK:
        return (
            '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" '
            'applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        )
    font = 3 if index < 40 else index % len(FONTS)
    number = NUMBER_FORMATS[index % len(NUMBER_FORMATS)]
    border = 1 if index % 7 == 0 else 0
    fill = 2 if index % 11 == 0 else 0
    return (
        f'<xf numFmtId="{number}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" '
        'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>'
    )


def styles() -> str:
    formats = "".join(
        f'<numFmt numFmtId="{i}" formatCode="{code.replace(chr(34), "&quot;")}"/>'
        for i, code in FORMATS
    )
    fonts = "".join(
        f'<font>{"<b/>" if b else ""}{"<i/>" if i else ""}<sz val="{size}"/>'
        f'<name val="{name}"/></font>'
        for name, size, b, i in FONTS
    )
    xfs = "".join(xf(i) for i in range(76))
    return (
        f'{DECLARATION}<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(FORMATS)}">{formats}</numFmts>'
        f'<fonts count="{len(FONTS)}">{fonts}</fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF99"/></patternFill></fill>'
        "</fills>"
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left/><right/><top/><bottom style="thin"><color auto="1"/></bottom>'
        "<diagonal/></border></borders>"
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        "</cellStyleXfs>"
        f'<cellXfs count="76">{xfs}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def shared_strings(count: int) -> str:
    words = ("Receipt", "Delivery", "Pipeline", "Meter", "Contract", "Volume", "Gas Day")
    items = "".join(
        f"<si><t>{words[i % len(words)]} {i:04d}{' adjustment' if i % 5 == 0 else ''}</t></si>"
        for i in range(STRINGS)
    )
    return f'{DECLARATION}<sst xmlns="{MAIN}" count="{count}" uniqueCount="{STRINGS}">{items}</sst>'


def data_rows(rows: int, formulas: int) -> Iterator[str]:
    """A header, then rows of four strings and eight numbers; ``formulas`` sums at the end."""
    letters = [column(c) for c in range(1, 13)]
    for r in range(1, rows + 1):
        cells = []
        for c, letter in enumerate(letters):
            style = DATA_STYLES[c]
            if c < 4:
                cells.append(
                    f'<c r="{letter}{r}" s="{style}" t="s"><v>{(r * 7 + c) % STRINGS}</v></c>'
                )
            else:
                value = ((r * 7919 + c * 104729) % 1000003) / 97.0
                cells.append(f'<c r="{letter}{r}" s="{style}"><v>{value!r}</v></c>')
        height = ' ht="13.5" customHeight="1"' if r == 1 else ""
        yield f'<row r="{r}"{height}>{"".join(cells)}</row>'
    for k in range(formulas):
        r = rows + 1 + k
        letter = letters[4 + k]
        yield (
            f'<row r="{r}"><c r="{letter}{r}" s="4"><f>SUM({letter}2:{letter}{rows})</f>'
            f"<v>{float(k)!r}</v></c></row>"
        )


def sparse_rows(rows: int, width: int, every: int, formulas: int) -> Iterator[str]:
    """Styled empty cells, the first sixteen columns in data styles, a value every ``every``."""
    letters = [column(c) for c in range(1, width + 1)]
    styles = [DATA_STYLES[c % len(DATA_STYLES)] if c < 16 else BLANK for c in range(width)]
    for r in range(1, rows + 1):
        cells = [f'<c r="{letter}{r}" s="{s}"/>' for letter, s in zip(letters, styles, strict=True)]
        if r % every == 0:
            cells[0] = f'<c r="A{r}" s="2" t="s"><v>{r % STRINGS}</v></c>'
            cells[1] = f'<c r="B{r}" s="4"><v>{r * 3.25!r}</v></c>'
        height = ' ht="13.5" customHeight="1"' if r == 1 else ""
        yield f'<row r="{r}"{height}>{"".join(cells)}</row>'
    for k in range(formulas):
        r = rows + 1 + k
        yield f'<row r="{r}"><c r="B{r}" s="4"><f>SUM(B1:B{rows})</f><v>{float(k)!r}</v></c></row>'


def worksheet(rows: Iterator[str], last: str) -> Iterator[str]:
    yield (
        f'{DECLARATION}<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><dimension ref="A1:{last}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="12.75"/>'
        '<cols><col min="1" max="1" width="13.28515625" customWidth="1"/>'
        '<col min="2" max="2" width="12.140625" customWidth="1"/>'
        '<col min="14" max="14" width="10.7109375" customWidth="1"/></cols><sheetData>'
    )
    yield from rows
    yield "</sheetData></worksheet>"


def write(path: Path, scale: float) -> dict[str, int]:
    """Write the workbook at ``scale``; its shape: sheets, rows, cells, values and formulas."""
    shape = {"sheets": 0, "rows": 0, "cells": 0, "values": 0, "formulas": 0}
    strings = 0
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for index, (_, rows, width, every, kind) in enumerate(SHEETS, 1):
            rows = max(1, round(rows * scale))
            formulas = 2 if index != 2 and index != 4 else 0
            if kind == "data":
                body = data_rows(rows, 3)
                formulas = 3
                values = rows * 12
                strings += rows * 4
                last = f"L{rows + formulas}"
            else:
                body = sparse_rows(rows, width, every, formulas)
                values = 2 * (rows // every)
                strings += rows // every
                last = f"{column(width)}{rows + formulas}"
            with archive.open(f"xl/worksheets/sheet{index}.xml", "w", force_zip64=True) as part:
                buffer: list[str] = []
                for piece in worksheet(body, last):
                    buffer.append(piece)
                    if len(buffer) >= 256:
                        part.write("".join(buffer).encode())
                        buffer.clear()
                part.write("".join(buffer).encode())
            shape["sheets"] += 1
            shape["rows"] += rows + formulas
            shape["cells"] += rows * width + formulas
            shape["values"] += values + formulas
            shape["formulas"] += formulas
        names = [name for name, *_ in SHEETS]
        sheets = "".join(
            f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(names, 1)
        )
        archive.writestr(
            "xl/workbook.xml",
            f'{DECLARATION}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{sheets}</sheets>'
            '<calcPr calcId="191029"/></workbook>',
        )
        relations = "".join(
            f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(names) + 1)
        )
        n = len(names)
        relations += (
            f'<Relationship Id="rId{n + 1}" Type="{REL}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId{n + 2}" Type="{REL}/sharedStrings" '
            'Target="sharedStrings.xml"/>'
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f'{DECLARATION}<Relationships xmlns="{PKG}">{relations}</Relationships>',
        )
        archive.writestr("xl/styles.xml", styles())
        archive.writestr("xl/sharedStrings.xml", shared_strings(strings))
        overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, n + 1)
        )
        archive.writestr(
            "[Content_Types].xml",
            f"{DECLARATION}<Types "
            'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            f"{overrides}</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            f'{DECLARATION}<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
            f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
    return shape


def main(directory: str | None, scale: float) -> None:
    folder = Path(directory) if directory else Path(tempfile.mkdtemp(prefix="sheetpeek-bench-"))
    path = folder / f"large-{scale:g}.xlsx"
    shape = write(path, scale)
    size = path.stat().st_size / 2**20
    print(f"{path}: {size:.1f} MiB, " + ", ".join(f"{k} {v:,}" for k, v in shape.items()))


if __name__ == "__main__":
    main(
        sys.argv[1] if len(sys.argv) > 1 else None,
        float(sys.argv[2]) if len(sys.argv) > 2 else 1.0,
    )
