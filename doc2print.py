#!/usr/bin/env python
"""
prints DOC, DOCX, RTF, etc. using LibreOffice.
"""
from pathlib import Path
from argparse import ArgumentParser
import subprocess


def main():
    p = ArgumentParser()
    p.add_argument('path', help='path to work in')
    p.add_argument('-g', '--glob', help='file extensions to print', nargs='+',
                   default=['*.doc', '*.docx', '*.rtf'])
    p = p.parse_args()

    path = Path(p.path).expanduser().resolve()

    for g in p.glob:
        for fn in path.glob(g):

            print('printing', fn)

            cmd = ['soffice', '-p', str(fn)]

            subprocess.check_call(cmd, cwd=path)


if __name__ == '__main__':
    main()
