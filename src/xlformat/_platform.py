"""Where the viewer's Excel runs (``Platform``) and its regional settings (``RegionalSettings``)."""

from __future__ import annotations

from dataclasses import dataclass, fields
from typing import ClassVar, Literal

PlatformName = Literal["mac", "windows"]
# what ``platform=None`` means; all the family's packages change it together
DEFAULT_PLATFORM: PlatformName = "mac"
# the display resolutions whose Windows glyph tables are captured
WINDOWS_DPIS = (96,)

_MONTHS = (
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December",
)  # fmt: skip
_DAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")


@dataclass(frozen=True, slots=True)
class Platform:
    """The platform whose Excel shows the sheet: how it lays out and measures text.

    ``Platform.MAC`` is Excel for Mac, what ``platform=None`` means while the default is
    ``"mac"``; ``Platform.WINDOWS`` is Excel for Windows at 96 dpi. A function given a platform
    it has no captured evidence for raises ``LookupError``: Windows is measured at 96 dpi for
    the width model (``format_value`` and ``format_cell`` in Excel mode, the column and glyph
    functions); rows, text placement, rules, ``text_function`` and ``builtin_format`` know the
    Mac only. Every result is in points; Windows' fall on its pixel grid (0.75 points at 96
    dpi). Fields can be overridden with :func:`dataclasses.replace`.

    Attributes:
        name: ``"mac"`` or ``"windows"``. The Mac measures in points, glyphs hinted at the
            font's size; Windows measures in device pixels, glyphs hinted at ``size * dpi / 72``.
        dpi: Windows only: 96 at 100% display scaling, 120 at 125%. ``None`` on the Mac.
        default_column_step: The device units a column without its own width is rounded up to:
            1 on the Mac, 8 pixels on Windows.
        long_fill: A ``*`` fill that would pass 255 characters: ``"hash"`` shows a ``#`` run
            (the Mac), ``"cap"`` stops the text at 255 characters (Windows).
    """

    name: PlatformName = "mac"
    dpi: int | None = None
    default_column_step: int = 1
    long_fill: Literal["hash", "cap"] = "hash"

    MAC: ClassVar[Platform]
    WINDOWS: ClassVar[Platform]

    @property
    def scale(self) -> float:
        """Device units per point: 1 on the Mac (points), ``dpi / 72`` on Windows (pixels)."""
        return 1.0 if self.dpi is None else self.dpi / 72

    def __post_init__(self) -> None:
        if self.name not in ("mac", "windows"):
            msg = f"platform name must be 'mac' or 'windows', not {self.name!r}"
            raise ValueError(msg)
        if self.name == "mac" and self.dpi is not None:
            msg = "dpi is a Windows setting; the Mac's must be None"
            raise ValueError(msg)
        if self.name == "windows" and (self.dpi is None or self.dpi <= 0):
            msg = f"Windows needs a positive dpi, not {self.dpi!r}"
            raise ValueError(msg)


Platform.MAC = Platform()
_MAC = Platform.MAC
Platform.WINDOWS = Platform("windows", dpi=96, default_column_step=8, long_fill="cap")


@dataclass(frozen=True, slots=True)
class RegionalSettings:
    """The viewer's regional settings, which some formats follow: separators, dates, names.

    Separate from the platform (how Excel lays out text) and from the workbook (``date1904``)
    and Excel's behaviors (:class:`Quirks`). Only en-US is supported: ``short_date`` is used
    (built-in formats 14 and 22 follow the system's short date), and every other field must keep
    its en-US value, or ``ValueError``.

    Attributes:
        decimal_separator: The decimal separator shown (format codes keep ``.``).
        thousands_separator: The thousands separator shown (format codes keep ``,``).
        short_date: The system short date, which built-in format 14 is and 22 begins with.
        long_date: The system long date, shown for ``[$-F800]``.
        time: The system time, shown for ``[$-F400]``.
        am_pm: The two strings of ``AM/PM``.
        month_names: January to December.
        day_names: Sunday to Saturday.
        currency: The currency symbol of the built-in currency formats.
    """

    decimal_separator: str = "."
    thousands_separator: str = ","
    short_date: str = "m/d/yy"
    long_date: str = "dddd, mmmm d, yyyy"
    time: str = "h:mm:ss AM/PM"
    am_pm: tuple[str, str] = ("AM", "PM")
    month_names: tuple[str, ...] = _MONTHS
    day_names: tuple[str, ...] = _DAYS
    currency: str = "$"

    EN_US_MAC: ClassVar[RegionalSettings]
    EN_US_WINDOWS: ClassVar[RegionalSettings]

    def __post_init__(self) -> None:
        for f in fields(RegionalSettings):
            if f.name != "short_date" and getattr(self, f.name) != f.default:
                msg = f"only en-US regional settings are supported: {f.name} must stay its default"
                raise ValueError(msg)


# Excel follows the system short date for built-ins 14 and 22; en-US defaults: the Mac M/d/yy,
# Windows M/d/yyyy (Windows' value to be confirmed by a capture before Windows is supported)
RegionalSettings.EN_US_MAC = RegionalSettings()
RegionalSettings.EN_US_WINDOWS = RegionalSettings(short_date="m/d/yyyy")


def check_platform(
    platform: str | Platform | None, function: str, *, windows: bool = False
) -> Platform:
    """The platform a call means; ``LookupError`` where ``function`` has no evidence for it.

    ``windows`` is True for a function with Windows evidence, which holds at the display
    resolutions of ``WINDOWS_DPIS``.
    """
    if platform is _MAC or platform == "mac" or (platform is None and DEFAULT_PLATFORM == "mac"):
        return _MAC  # the common case, kept cheap: every format call passes through here
    if platform is None:
        platform = DEFAULT_PLATFORM
    if isinstance(platform, str):
        presets = {"mac": Platform.MAC, "windows": Platform.WINDOWS}
        if platform not in presets:
            msg = f"platform must be 'mac', 'windows' or a Platform, not {platform!r}"
            raise ValueError(msg)
        platform = presets[platform]
    if windows and platform.name == "windows":
        if platform.dpi not in WINDOWS_DPIS:
            msg = f"{function} has no captured evidence for Excel for Windows at {platform.dpi} dpi"
            raise LookupError(msg)
        return platform
    if platform != Platform.MAC:
        what = "Excel for Windows" if platform.name == "windows" else "a Mac with other rules"
        msg = f"{function} has no captured evidence for {what}; Platform.MAC is supported"
        raise LookupError(msg)
    return platform


def check_regional(
    regional_settings: RegionalSettings | None, platform: Platform
) -> RegionalSettings:
    """The regional settings a call means: the platform's en-US preset by default."""
    if regional_settings is None:
        return (
            RegionalSettings.EN_US_MAC if platform.name == "mac" else RegionalSettings.EN_US_WINDOWS
        )
    return regional_settings
