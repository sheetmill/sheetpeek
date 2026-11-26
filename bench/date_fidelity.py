"""How often openpyxl's date conversion changes what a date format displays (ARCHITECTURE.md §5.3).

    uv run python bench/date_fidelity.py [COUNT]

Formats random serials with xlformat twice, once from the serial and once from the
``datetime`` that openpyxl's ``from_excel`` makes of it, and counts the differences per format.
Also shows the serials around the 1900 leap day.
"""

from __future__ import annotations

import random
import sys

import xlformat
from openpyxl.utils.datetime import from_excel

FORMATS = ("hh:mm:ss", "h:mm", "hh:mm:ss.000", "yyyy-mm-dd hh:mm:ss", "[h]:mm:ss")


def main() -> None:
    count = int(sys.argv[1]) if len(sys.argv) > 1 else 200_000
    print(f"xlformat {xlformat.__version__}, {count} serials per format, seed 7")
    rng = random.Random(7)
    for fmt in FORMATS:
        elapsed = fmt.startswith("[")
        differ = 0
        example = None
        for _ in range(count):
            serial = rng.uniform(61, 80000)
            exact = xlformat.format_value(serial, fmt, platform="mac")
            converted = xlformat.format_value(
                from_excel(serial, timedelta=elapsed), fmt, platform="mac"
            )
            if exact != converted:
                differ += 1
                example = example or (serial, exact, converted)
        print(f"  {fmt!r:24s} {differ:6d} differ  e.g. {example}")
    for serial in (59, 60, 61):
        value = from_excel(serial)
        print(
            f"  serial {serial}: openpyxl {value!r}, Excel "
            f"{xlformat.format_value(serial, 'yyyy-mm-dd', platform='mac')}, "
            f"from the datetime {xlformat.format_value(value, 'yyyy-mm-dd', platform='mac')}"
        )


if __name__ == "__main__":
    main()
