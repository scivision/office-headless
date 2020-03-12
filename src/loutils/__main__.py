#!/usr/bin/env python
from argparse import ArgumentParser

from .convert import doc2pdf as to_pdf
from .print import doc2print as to_printer
from .find import docfinder


def doc2pdf():
    """
    Convert DOC, DOCX, ODT, RTF, etc. to PDF using LibreOffice
    It leaves the original files alone, putting the PDF in the same directory.

    NOTE: LibreOffice conversion is not thread-safe--it will
    silently fail if simultaneous processes are launched--sometimes.
    Thus, we use the "dumb" for loop, as the command line globbing doesn't work either,
    even just from Terminal!
    """
    p = ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to convert", nargs="+")
    p = p.parse_args()

    for f in docfinder(p.path, p.suffix, exclude=".pdf"):
        to_pdf(f)


def doc2print():
    """
    prints DOC, DOCX, RTF, PPT, XLS, etc. using specified program.
    Due to limitations of programs, the DEFAULT PRINTER is used.
    This might not be the printer you expect, so do this offline / directly plugged to printer.
    """
    p = ArgumentParser()
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

        to_printer(f, p.exe)
