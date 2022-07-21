from pathlib import Path
import subprocess
import shutil
import argparse

from . import docfinder

LOEXE = shutil.which("soffice")


def doc2pdf(filein: Path):
    if not LOEXE:
        raise EnvironmentError("LibreOffice not found")

    cmd = [LOEXE, "--convert-to", "pdf", "--outdir", str(filein.parent), str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)


if __name__ == "__main__":
    """
    Convert DOC, DOCX, ODT, RTF, etc. to PDF using LibreOffice
    It leaves the original files alone, putting the PDF in the same directory.

    NOTE: LibreOffice conversion is not thread-safe--it will
    silently fail if simultaneous processes are launched--sometimes.
    For loop, as the command line globbing doesn't work either,
    even just from Terminal!
    """
    p = argparse.ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to convert", nargs="+")
    p = p.parse_args()

    for f in docfinder(p.path, p.suffix, exclude=".pdf"):
        doc2pdf(f)
