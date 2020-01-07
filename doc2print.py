#!/usr/bin/env python
"""
prints DOC, DOCX, RTF, PPT, XLS, etc. using specified program.
Due to limitations of programs, the DEFAULT PRINTER is used.
This might not be the printer you expect, so do this offline / directly plugged to printer.
"""
from argparse import ArgumentParser

import loutils


def main():
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

    for f in loutils.docfinder(p.path, p.suffix):
        yesno = input(f"press Enter to print: {f}   or 's' to skip or 'q' to quit  ")
        if yesno == "q":
            break
        elif yesno == "s":
            continue

        loutils.doc2print(f, p.exe)


if __name__ == "__main__":
    main()
