from __future__ import annotations
import typing as T
from pathlib import Path
import shutil
import os
import sys
import functools

__version__ = "1.4.0"

# whatever you might like to convert
SUFFIXES = (
    ".doc",
    ".docx",
    ".odt",
    ".pages",
    ".pub",
    ".rtf",
    ".wpd",
    ".wps",
    ".wri",
    ".xls",
    ".xlsx",
    ".ods",
    ".wks",
    ".ppt",
    ".pptx",
    ".odp",
)


@functools.cache
def get_lo_exe() -> str:
    name = "soffice"

    if os.name == "nt":
        path = Path(os.environ["PROGRAMFILES"]) / "LibreOffice/program"
    elif sys.platform == "darwin":
        path = Path("/Applications/LibreOffice.app/Contents/MacOS")
    else:
        path = None

    if not (exe := shutil.which(name, path=path)):
        raise FileNotFoundError("LibreOffice not found")

    return exe


@functools.cache
def get_word_exe() -> str:
    if os.name == "nt":
        name = "winword"
        path = None
    else:
        name = "Microsoft Word"
        path = "/Applications/Microsoft Word.app/Contents/MacOS/"

    if not (exe := shutil.which(name, path=path)):
        raise FileNotFoundError("Microsoft Word not found")

    return exe


@functools.cache
def get_powerpoint_exe() -> str:
    if os.name == "nt":
        name = "powerpnt"
        path = None
    else:
        name = "Microsoft PowerPoint"
        path = "/Applications/Microsoft PowerPoint.app/Contents/MacOS/"

    if not (exe := shutil.which(name, path=path)):
        raise FileNotFoundError("Microsoft PowerPoint not found")

    return exe


@functools.cache
def get_adobe_exe() -> str:
    """
    Adobe is no longer available on Linux
    """

    if os.name == "nt":
        name = "acrord32"
        path = Path(os.environ["PROGRAMFILES(x86)"]) / "Adobe/Acrobat Reader DC/Reader"
    else:
        name = "AdobeAcrobat"
        path = Path("/Applications/Adobe Acrobat DC/Adobe Acrobat.app/Contents/MacOS")

    if not (exe := shutil.which(name, path=path)):
        raise FileNotFoundError("Adobe Acrobat not found")

    return exe


def docfinder(
    path: Path, suffixes: tuple[str, ...] | None = None, exclude: str | None = None
) -> T.Iterator[Path]:
    """
    Find documents on a directory,
    optionally restricting to particular suffixes or excluding a suffix.
    """

    path = Path(path).expanduser().resolve()

    if not suffixes:
        suffixes = SUFFIXES

    if path.is_file():
        yield path
    elif exclude:
        for f in path.iterdir():
            if (
                f.is_file()
                and f.suffix.lower() in suffixes
                and not f.with_suffix(exclude).is_file()
            ):
                yield f
    else:
        for f in path.iterdir():
            if f.is_file() and f.suffix in suffixes:
                yield f
