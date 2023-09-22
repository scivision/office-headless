from pathlib import Path
import subprocess
import argparse

from . import docfinder, get_lo_exe, get_word_exe, get_powerpoint_exe


def doc2print(filein: Path, exe: str):
    """
    On macOS, use LibreOffice.
    It doesn't seem there's Microsoft Office command line on macOS.

    macOS updater command line:
    https://learn.microsoft.com/en-us/deployoffice/mac/update-office-for-mac-using-msupdate

    Command line for Microsoft Office products:
    https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6
    """

    if exe == "libreoffice":
        cmd = [get_lo_exe(), "-p", str(filein)]
    elif "word" in exe:
        cmd = [get_word_exe(), "/x", "/mFilePrintDefault", "/t", str(filein)]
    elif exe == "powerpoint":
        cmd = [get_powerpoint_exe(), "/P", str(filein)]
    elif exe == "wordpad":
        cmd = ["write.exe", "/p", str(filein)]
    elif exe == "notepad++":
        cmd = ["notepad++", "-quickPrint", str(filein)]
    elif exe == "acroread":
        cmd = [
            "C:/Program Files (x86)/Adobe/Acrobat Reader DC/Reader/acrord32.exe",
            "/p",
            str(filein),
        ]
    else:
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
        choices=["libreoffice", "word", "wordpad", "notepad++", "powerpoint", "acroread"],
        default="libreoffice",
    )
    P = p.parse_args()

    if P.exe == "wordpad":
        print(
            "\nNOTE: WordPad sometimes wastes an extra page printed since WordPad has trouble"
            " decoding contemporary Word documents. \nConsider LibreOffice instead.\n"
        )

    for f in docfinder(P.path, P.suffix):
        yesno = input(f"press Enter to print: {f}   or 's' to skip or 'q' to quit  ")
        if yesno == "q":
            break
        elif yesno == "s":
            continue

        doc2print(f, P.exe)
