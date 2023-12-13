#!/usr/bin/env python3
"""
Single file standalone script to convert a text/document file to a PDF file using LibreOiffice.
"""

import subprocess
import shutil
import os
import sys
from pathlib import Path
import argparse


def get_lo_exe() -> str:
    name = "soffice"

    if os.name == "nt":
        path = Path(os.environ["PROGRAMFILES"]) / "LibreOffice/program"
    elif sys.platform == "darwin":
        path = Path("/Applications/LibreOffice.app/Contents/MacOS")
    else:
        path = None

    if not (exe := shutil.which(name, path=path)):
        raise FileNotFoundError("LibreOffice not found")

    return exe


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Convert a document file to PDF using LibreOffice")
    p.add_argument("filein", help="Input file")
    p.add_argument("out_dir", help="Output directory", nargs="?")
    args = p.parse_args()

    filein = Path(args.filein).expanduser().resolve()

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else filein.parent

    cmd = [get_lo_exe(), "--convert-to", "pdf", "--outdir", str(out_dir), str(filein)]

    subprocess.check_call(cmd)
