#!/usr/bin/env python
"""
prints DOC, DOCX, RTF, etc. using LibreOffice.
"""
from argparse import ArgumentParser
import subprocess
from loutils import EXE, docfinder


def main():
    p = ArgumentParser()
    p.add_argument('path', help='path to work in')
    p.add_argument('-s', '--suffix',
                   help='file extensions to print', nargs='+',)
    p = p.parse_args()

    flist = docfinder(p.path, p.suffix)

    for f in flist:
        print('printing', f)

        cmd = [EXE, '-p', str(f)]

        subprocess.check_call(cmd)


if __name__ == '__main__':
    main()
