#!/usr/bin/env python
"""
Convert DOC, DOCX, ODT, RTF, etc. to PDF using LibreOffice
It leaves the original files alone, putting the PDF in the same directory.

NOTE: LibreOffice conversion (at least through 6.2) is not thread-safe--it will
silently fail if simultaneous processes are launched--sometimes.
Thus, we use the "dumb" for loop, as the command line globbing doesn't work either,
even just from Terminal!
"""
from pathlib import Path
from argparse import ArgumentParser
import asyncio
import logging
import subprocess

from loutils import EXE, docfinder


async def arbiter_concurrent(filein: Path):
    """ concurrent subprocess """
    assert isinstance(EXE, str)

    cmd = [EXE,
           '--convert-to', 'pdf',
           '--outdir', str(filein.parent),
           str(filein)]

    proc = await asyncio.create_subprocess_exec(*cmd,
                                                stderr=subprocess.DEVNULL)

    ret = await proc.wait()
    # print(' '.join(cmd), ret)

    if ret != 0:
        logging.warning(f'FAILED: {filein}')


def arbiter_serial(filein: Path):
    """ Since Libreoffice isn't threadsaft """
    assert isinstance(EXE, str)

    cmd = [EXE,
           '--convert-to', 'pdf',
           '--outdir', str(filein.parent),
           str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)


def main():
    p = ArgumentParser()
    p.add_argument('path', help='path to work in')
    p.add_argument('-s', '--suffix',
                   help='file extensions to convert', nargs='+')
    p = p.parse_args()

    flist = docfinder(p.path, p.suffix, exclude='.pdf')
    """
        futures = [arbiter_concurrent(f) for f in flist]

        if not futures:
            raise SystemExit(f'no files found in {p.path} to convert to PDF')

        asyncio.run(asyncio.wait(futures))
    """
# %% fallback until LibreOffice is threadsafe
    for f in flist:
        arbiter_serial(f)


if __name__ == '__main__':
    main()
