from pathlib import Path
import subprocess
import argparse

from . import docfinder, get_lo_exe, get_word_exe, get_powerpoint_exe, get_adobe_exe


def doc2print(filein: Path, exe: str):
    """
    On macOS, use LibreOffice.
    It doesn't seem there's Microsoft Office command line on macOS.

    macOS updater command line:
    https://learn.microsoft.com/en-us/deployoffice/mac/update-office-for-mac-using-msupdate

    Command line for Microsoft Office products:
    https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6

    We don't use WordPad as it adds an extra page and has difficulty with modern Word docs.
    """

    match exe:
        case "libreoffice":
            cmd = [get_lo_exe(), "-p", str(filein)]
        case "word":
            cmd = [get_word_exe(), "/x", "/mFilePrintDefault", "/t", str(filein)]
        case "powerpoint":
            cmd = [get_powerpoint_exe(), "/P", str(filein)]
        case "notepad++":
            cmd = ["notepad++", "-quickPrint", str(filein)]
        case "acroread":
            cmd = [get_adobe_exe(), "/p", str(filein)]
        case _:
            raise ValueError(f"unknown program {exe}")

    subprocess.check_call(cmd)


if __name__ == "__main__":
    """
    prints DOC, DOCX, RTF, PPT, XLS, etc. using specified program.
    Due to limitations of programs, the DEFAULT PRINTER is used.
    This might not be the printer you expect, so do this offline / directly plugged to printer.
    """
    p = argparse.ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to print", nargs="+")
    p.add_argument(
        "-exe",
        help="printing program",
        choices=["libreoffice", "word", "notepad++", "powerpoint", "acroread"],
        default="libreoffice",
    )
    P = p.parse_args()

    for f in docfinder(P.path, P.suffix):
        yesno = input(f"press Enter to print: {f}   or 's' to skip or 'q' to quit  ")
        match yesno:
            case "q":
                break
            case "s":
                continue

        doc2print(f, P.exe)
