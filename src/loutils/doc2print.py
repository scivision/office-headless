from pathlib import Path
import subprocess
import argparse

from . import docfinder


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        cmd = ["soffice", "-p", str(filein)]
    elif exe == "winword":
        cmd = ["winword.exe", "/q", "/x", "/mFilePrintDefault", "/t", str(filein)]
    elif exe == "powerpoint":
        cmd = ["powerpnt.exe", "/P", str(filein)]
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
        choices=["libreoffice", "winword", "wordpad", "notepad++", "powerpoint", "acroread"],
        default="libreoffice",
    )
    p = p.parse_args()

    if p.exe == "wordpad":
        print(
            "\nNOTE: WordPad sometimes wastes an extra page printed since WordPad has trouble"
            " decoding contemporary Word documents. \nConsider LibreOffice instead.\n"
        )

    for f in docfinder(p.path, p.suffix):
        yesno = input(f"press Enter to print: {f}   or 's' to skip or 'q' to quit  ")
        if yesno == "q":
            break
        elif yesno == "s":
            continue

        doc2print(f, p.exe)
