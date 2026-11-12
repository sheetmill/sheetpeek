"""How a capture's header names the Excel it was made with: version, build, edition, channel.

The edition (the license, as Excel's About box names it, such as "Microsoft 365") is not in
Excel's dictionary, so the run states it in ``EXCEL_EDITION``; the build and update channel are
read from the application and from Microsoft AutoUpdate's settings.
"""

from __future__ import annotations

import os

from capture.screen.session import environment


def excel_description() -> str:
    """``Excel 16.115 (26110837, Microsoft 365, Current channel) for Mac``, as found."""
    found = environment(os.environ.get("EXCEL_EDITION", "edition not stated"))
    return (
        f"Excel {found['excel']} ({found['build']}, {found['edition']}, {found['channel']} "
        "channel) for Mac"
    )
