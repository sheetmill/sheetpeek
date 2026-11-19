"""One generated workbook open in Excel for Mac, for a screen capture.

Every request goes through ``excel.applescript``, one request a call (count, open, close, view,
scroll, window, save, read), so each call does one visible thing and can be inspected on its own.
``ExcelSession`` refuses to start unless Excel answers, holds no workbook and shows no window
of its own (About, Preferences, a dialog), since any other window spoils a window capture.
It opens the workbook, keeps the display awake, and on every exit path
closes it without saving, turns ``display alerts`` back on, deletes the file and records how many
workbooks are left open. A request past its time limit is killed (``osascript`` waiting on Excel
ignores SIGTERM, so it gets SIGKILL), evidence is saved (``sample`` of Excel, whether it still
answers), and ``ExcelTimeout`` is raised. Standard library only.
"""

from __future__ import annotations

import contextlib
import datetime as dt
import os
import platform
import subprocess
from pathlib import Path
from types import TracebackType

HERE = Path(__file__).resolve().parent
SCRIPT = HERE / "excel.applescript"
WINDOWS = HERE / "windows.js"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents"
"""Excel is sandboxed: a workbook it opens by script lives here, so no "Grant Access" shows."""
APP = Path("/Applications/Microsoft Excel.app")
US, RS = "\x1f", "\x1e"
"""Field and record separators in ``read``'s output."""


class ExcelError(RuntimeError):
    """Excel refused, or is in a state where requests must not be sent."""


class ExcelTimeout(ExcelError):  # noqa: N818 - a kind of ExcelError
    """A request passed its time limit; evidence was saved."""


def request(*arguments: str, timeout: float = 60.0) -> str:
    """Run one request of ``excel.applescript``; ``subprocess.TimeoutExpired`` past ``timeout``."""
    result = subprocess.run(
        ["osascript", str(SCRIPT), *arguments], capture_output=True, text=True, timeout=timeout
    )
    if result.returncode:
        raise ExcelError(result.stderr.strip()[:400])
    return result.stdout.rstrip("\n")


def excel_windows() -> list[dict[str, object]]:
    """Excel's windows on screen: CoreGraphics number, bounds in points, title (may be empty)."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", str(WINDOWS)], capture_output=True, text=True, timeout=15
    )
    if result.returncode:
        raise ExcelError(result.stderr.strip()[:400])
    windows: list[dict[str, object]] = []
    for line in result.stdout.splitlines():
        if not line.strip():  # no window prints an empty line
            continue
        number, x, y, w, h, title = [*line.split("\t"), ""][:6]
        windows.append(
            {"id": int(number), "x": float(x), "y": float(y), "w": float(w), "h": float(h),
             "title": title}
        )  # fmt: skip
    return windows


def environment(edition: str) -> dict[str, str]:
    """What a capture header records about the machine and Excel.

    ``edition`` is the license as Excel's About box names it (such as "Microsoft 365"); Excel's
    dictionary does not give it. The update channel comes from Microsoft AutoUpdate's settings.
    """

    def read(*command: str) -> str:
        try:
            return subprocess.run(
                command, capture_output=True, text=True, timeout=10
            ).stdout.strip()
        except (OSError, subprocess.TimeoutExpired):
            return ""

    info = APP / "Contents/Info.plist"
    channel = "unknown"
    for line in read("defaults", "read", "com.microsoft.autoupdate2").splitlines():
        key, _, value = line.strip().partition("=")
        if key.strip() == "ChannelName":
            channel = value.strip().strip(";").strip().strip('"') or "unknown"
    dark = read("defaults", "read", "-g", "AppleInterfaceStyle") == "Dark"
    return {
        "excel": read("defaults", "read", str(info), "CFBundleShortVersionString") or "unknown",
        "build": read("defaults", "read", str(info), "CFBundleVersion") or "unknown",
        "edition": edition,
        "channel": channel,
        "macos": platform.mac_ver()[0] or "unknown",
        "appearance": "dark" if dark else "light",
        "date": dt.date.today().isoformat(),
    }


class ExcelSession:
    """One generated workbook open in Excel, closed and deleted on every exit path.

    ``path`` must be inside ``CONTAINER``. Use as a context manager.
    """

    def __init__(self, path: Path, evidence: Path, timeout: float = 60.0) -> None:
        self.path = Path(path)
        self.name = self.path.name
        self.evidence = Path(evidence)
        self.timeout = timeout
        self.open_after: str | None = None
        self._awake: subprocess.Popen[bytes] | None = None
        self._copies: list[Path] = []

    def ask(self, *arguments: str, timeout: float | None = None) -> str:
        """One request of ``excel.applescript``, evidence saved if it does not answer."""
        limit = timeout or self.timeout
        try:
            return request(*arguments, timeout=limit)
        except subprocess.TimeoutExpired:
            self._collect(" ".join(arguments[:3]))
            raise ExcelTimeout(
                f"no answer within {limit:.0f} s; evidence in {self.evidence}"
            ) from None

    def _collect(self, what: str) -> None:
        self.evidence.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        notes = [f"request: {what}"]
        try:
            notes.append("count of workbooks: " + request("count", timeout=15))
        except (subprocess.TimeoutExpired, ExcelError) as exc:
            notes.append(f"count of workbooks: no answer ({type(exc).__name__})")
        sample = self.evidence / f"sample-{stamp}.txt"
        subprocess.run(
            ["sample", "Microsoft Excel", "10", "-file", str(sample)],
            capture_output=True,
            timeout=60,
        )
        (self.evidence / f"timeout-{stamp}.txt").write_text("\n".join(notes), encoding="utf-8")

    def preflight(self) -> None:
        """Refuse to start unless Excel answers, holds no workbook and shows no window."""
        try:
            count = request("count", timeout=15)
        except subprocess.TimeoutExpired:
            raise ExcelError("Excel does not answer `count of workbooks`") from None
        if count != "0":
            raise ExcelError(f"{count} workbook(s) open in Excel: close them first")
        windows = excel_windows()
        if windows:
            listed = ", ".join(f"{w['w']:.0f}x{w['h']:.0f} {w['title']!r}" for w in windows)
            raise ExcelError(f"Excel shows windows with no workbook open ({listed}): close them")

    def __enter__(self) -> ExcelSession:
        if CONTAINER not in self.path.parents:
            raise ExcelError(f"workbooks must be inside {CONTAINER}")
        self.preflight()
        # The display stays awake while Excel is driven: unexplained hangs came with it off.
        self._awake = subprocess.Popen(["caffeinate", "-d", "-i", "-w", str(os.getpid())])
        try:
            self.ask("open", str(self.path))
        except BaseException:
            self._close()
            raise
        return self

    def __exit__(
        self,
        kind: type[BaseException] | None,
        error: BaseException | None,
        trace: TracebackType | None,
    ) -> None:
        self._close()

    def _close(self) -> None:
        try:
            self.open_after = self.ask("close", self.name, timeout=30)
        except ExcelError:
            self.open_after = "no answer"
            # alerts must never stay off, whatever happened above
            with contextlib.suppress(subprocess.TimeoutExpired, ExcelError):
                request("alerts", timeout=15)
        finally:
            self.path.unlink(missing_ok=True)
            for copy in self._copies:
                copy.unlink(missing_ok=True)
            if self._awake is not None:
                self._awake.terminate()

    def save_copy(self, copy: Path) -> Path:
        """Have Excel save the open workbook as ``copy`` (inside the container), and follow it.

        After saving, the open workbook is the copy, so later requests and the closing address it
        by its new name. The copy is deleted on exit, like the original: read it before leaving.
        """
        copy = Path(copy)
        if CONTAINER not in copy.parents:
            raise ExcelError(f"copies must be inside {CONTAINER}")
        self._copies.append(copy)
        self.ask("save", self.name, str(copy))
        self.name = copy.name
        return copy

    def view(self, park: str, zoom: int = 100) -> str:
        """Bring the workbook forward at ``zoom``, the selection on ``park``, scrolled to A1.

        The selection outline and its highlighted headings then sit outside the captured area.
        """
        return self.ask("view", self.name, park, str(zoom))

    def scroll(self, row: int) -> str:
        """Scroll the window so ``row`` is its top row, column A at its left.

        Returns the address of the window's visible range (``$A$9990:$AQ$10059``), as Excel
        gives it: rows and columns shown in part count.
        """
        return self.ask("scroll", self.name, str(row))

    def window(self) -> str:
        """The workbook's active window's state as Excel reports it, nothing set first.

        ``zoom=85;view=normal view;freeze_panes=true;split_row=2;...``: the zoom, the view, freeze
        panes, the split row and column, the scroll row and column; a field Excel refuses is
        ``ERR``.
        """
        return self.ask("window", self.name)

    def read_cells(self, cells: list[str]) -> list[tuple[str, str, str]]:
        """(cell, property, value) for every cell: display format, own format, borders, geometry.

        Values are as Excel returns them; one Excel refuses is ``ERR`` and its message.
        """
        text = self.ask("read", self.name, *cells, timeout=max(60.0, 3.0 * len(cells)))
        rows = []
        for record in text.split(RS):
            if record:
                cell, key, value = [*record.split(US), "", ""][:3]
                rows.append((cell, key, value))
        return rows
