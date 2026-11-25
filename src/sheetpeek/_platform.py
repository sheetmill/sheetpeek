"""Which Excel a sheet is shown as: the platform and the regional settings, resolved.

``platform=`` takes ``"mac"``, ``"windows"`` or an :class:`xlformat.Platform`, as xlformat takes
it; ``regional_settings=`` an :class:`xlformat.RegionalSettings`, by default the platform's en-US
settings. Both come from xlformat, and sheetpeek has no settings object of its own: what it
draws differently follows from the platform. The default is Excel for Mac.

Text is measured where the layout measures it, deep in helpers that do not hold the sheet:
``measuring`` sets the platform for the length of a layout, and ``text_width`` measures with it.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from functools import lru_cache
from typing import TypeAlias

import xlformat

PlatformLike: TypeAlias = str | xlformat.Platform | None

_NAMES = {"mac": xlformat.Platform.MAC, "windows": xlformat.Platform.WINDOWS}
_MEASURING: ContextVar[xlformat.Platform] = ContextVar("measuring", default=xlformat.Platform.MAC)


def resolve(platform: PlatformLike) -> xlformat.Platform:
    """The platform a name or an object gives; ``None`` is Excel for Mac.

    Raises:
        ValueError: A name other than ``"mac"`` and ``"windows"``.
        TypeError: Neither a name nor an :class:`xlformat.Platform`.
    """
    if platform is None:
        return xlformat.Platform.MAC
    if isinstance(platform, xlformat.Platform):
        return platform
    if isinstance(platform, str):
        try:
            return _NAMES[platform]
        except KeyError:
            raise ValueError(f"platform must be 'mac' or 'windows', not {platform!r}") from None
    raise TypeError(f"platform must be 'mac', 'windows' or an xlformat.Platform, not {platform!r}")


def regional(
    settings: xlformat.RegionalSettings | None, platform: xlformat.Platform
) -> xlformat.RegionalSettings:
    """The regional settings given, or by default the platform's en-US settings.

    Raises:
        TypeError: Not an :class:`xlformat.RegionalSettings`.
    """
    if settings is None:
        if platform.name == "windows":
            return xlformat.RegionalSettings.EN_US_WINDOWS
        return xlformat.RegionalSettings.EN_US_MAC
    if isinstance(settings, xlformat.RegionalSettings):
        return settings
    raise TypeError(f"regional_settings must be an xlformat.RegionalSettings, not {settings!r}")


@contextmanager
def measuring(platform: xlformat.Platform) -> Iterator[None]:
    """Measure text as ``platform`` draws it until the block ends."""
    token = _MEASURING.set(platform)
    try:
        yield
    finally:
        _MEASURING.reset(token)


def current() -> xlformat.Platform:
    """The platform text is measured as: the layout's, or Excel for Mac outside one."""
    return _MEASURING.get()


def text_width(text: str, font: xlformat.Font | None) -> float:
    """``xlformat.text_width`` on the platform being measured."""
    return xlformat.text_width(text, font=font, platform=_MEASURING.get())


_PROBES: dict[str, Callable[[xlformat.Font, xlformat.Platform], object]] = {
    "text_width": lambda font, platform: xlformat.text_width("0", font=font, platform=platform),
    "text_insets": lambda font, platform: xlformat.text_insets(font, platform=platform),
    "row_metrics": lambda font, platform: xlformat.row_metrics(font, platform=platform),
    "vertical_metrics": lambda font, platform: xlformat.vertical_metrics(font, platform=platform),
    "rule_metrics": lambda font, platform: xlformat.rule_metrics(font, platform=platform),
}
"""The xlformat functions that take a font and measure only some families, each called at 11
points: the size every one of them covers."""


@lru_cache(maxsize=1024)
def covers(function: str, name: str, platform: xlformat.Platform) -> bool:
    """Whether xlformat's ``function`` measures the family ``name`` on ``platform``.

    Asked of xlformat, so a family it adds is measured without a change here. A face the
    platform does not have is measured in the face xlformat substitutes for it, in every
    function (on the Mac, MS Sans Serif as Times New Roman).
    """
    try:
        _PROBES[function](xlformat.Font(name, 11.0), platform)
    except LookupError:
        return False
    return True


def metrics_face(name: str, function: str, platform: xlformat.Platform) -> str:
    """The family to pass to xlformat's ``function``.

    ``name``'s face where it measures it, else Calibri, at the same size, as the layout measures
    a face it lacks (§6.2).
    """
    return name if covers(function, name, platform) else "Calibri"
