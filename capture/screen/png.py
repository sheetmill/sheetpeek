"""A minimal PNG reader (8-bit RGB or RGBA, not interlaced), enough for ``screencapture`` output.

Standard library only. ``Image.pixel(x, y)`` gives ``(r, g, b)`` as stored: values are in the
image's color profile (``Image.profile`` names it), which need not be sRGB.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass
from pathlib import Path


@dataclass
class Image:
    width: int
    height: int
    channels: int
    rows: list[bytes]
    profile: str

    def pixel(self, x: int, y: int) -> tuple[int, int, int]:
        row = self.rows[y]
        i = x * self.channels
        return row[i], row[i + 1], row[i + 2]


def _paeth(a: int, b: int, c: int) -> int:
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    if pa <= pb and pa <= pc:
        return a
    return b if pb <= pc else c


def read(path: str | Path) -> Image:
    data = Path(path).read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    pos, idat, profile = 8, [], "none (sRGB assumed)"
    width = height = channels = 0
    while pos < len(data):
        length, kind = struct.unpack(">I4s", data[pos : pos + 8])
        chunk = data[pos + 8 : pos + 8 + length]
        pos += 12 + length
        if kind == b"IHDR":
            width, height, depth, color, _, _, interlace = struct.unpack(">IIBBBBB", chunk)
            if depth != 8 or color not in (2, 6) or interlace:
                raise ValueError(
                    f"unsupported PNG: depth {depth}, color type {color}, interlace {interlace}"
                )
            channels = 3 if color == 2 else 4
        elif kind == b"iCCP":
            profile = chunk.split(b"\0", 1)[0].decode("latin-1")
        elif kind == b"IDAT":
            idat.append(chunk)
        elif kind == b"IEND":
            break
    raw = zlib.decompress(b"".join(idat))
    stride = width * channels
    rows: list[bytes] = []
    prev = bytearray(stride)
    i = 0
    for _ in range(height):
        kind = raw[i]
        line = bytearray(raw[i + 1 : i + 1 + stride])
        i += 1 + stride
        if kind == 1:
            for x in range(channels, stride):
                line[x] = (line[x] + line[x - channels]) & 255
        elif kind == 2:
            for x in range(stride):
                line[x] = (line[x] + prev[x]) & 255
        elif kind == 3:
            for x in range(stride):
                left = line[x - channels] if x >= channels else 0
                line[x] = (line[x] + ((left + prev[x]) >> 1)) & 255
        elif kind == 4:
            for x in range(stride):
                left = line[x - channels] if x >= channels else 0
                upleft = prev[x - channels] if x >= channels else 0
                line[x] = (line[x] + _paeth(left, prev[x], upleft)) & 255
        rows.append(bytes(line))
        prev = line
    return Image(width, height, channels, rows, profile)


def write_crop(image: Image, box: tuple[int, int, int, int], path: str | Path) -> None:
    """Write ``box`` (x0, y0, x1, y1, exclusive) of ``image`` as an RGB PNG, for inspection."""
    x0, y0, x1, y1 = box
    raw = bytearray()
    for y in range(y0, y1):
        raw.append(0)
        for x in range(x0, x1):
            raw += bytes(image.pixel(x, y))

    def chunk(kind: bytes, body: bytes) -> bytes:
        return (
            struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
        )

    header = struct.pack(">IIBBBBB", x1 - x0, y1 - y0, 8, 2, 0, 0, 0)
    with open(path, "wb") as out:
        out.write(
            b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(bytes(raw)))
            + chunk(b"IEND", b"")
        )
