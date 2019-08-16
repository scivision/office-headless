#!/usr/bin/env python
"""
Convert DOC, DOCX, ODT, RTF, etc. to PDF using LibreOffice
It leaves the original files alone, putting the PDF in the same directory.

NOTE: LibreOffice conversion is not thread-safe--it will
silently fail if simultaneous processes are launched--sometimes.
Thus, we use the "dumb" for loop, as the command line globbing doesn't work either,
even just from Terminal!
"""
from argparse import ArgumentParser

import loutils


def main():
    p = ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to convert", nargs="+")
    p = p.parse_args()

    for f in loutils.docfinder(p.path, p.suffix, exclude=".pdf"):
        loutils.doc2pdf(f)


if __name__ == "__main__":
    main()
