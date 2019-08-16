#!/usr/bin/env python
"""
prints DOC, DOCX, RTF, etc. using LibreOffice.
"""
from argparse import ArgumentParser

import loutils


def main():
    p = ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to print", nargs="+")
    p.add_argument(
        "-exe",
        help="conversion program (libreoffice or winword)",
        default="libreoffice",
    )
    p = p.parse_args()

    for f in loutils.docfinder(p.path, p.suffix):
        yesno = input("Enter 'y' to print:", f)
        if not yesno == "y":
            continue

        loutils.doc2print(f, p.exe)


if __name__ == "__main__":
    main()
