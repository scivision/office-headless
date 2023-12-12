from pathlib import Path
import subprocess
import argparse

from . import docfinder, get_lo_exe


def doc2pdf(filein: Path, out_dir: Path | None = None):

    if not out_dir:
        out_dir = filein.parent

    cmd = [get_lo_exe(), "--convert-to", "pdf", "--outdir", str(out_dir), str(filein)]

    subprocess.check_call(cmd)


if __name__ == "__main__":
    """
    Convert plain text, DOC, DOCX, ODT, RTF, etc. to PDF using LibreOffice
    It leaves the original files alone, putting the PDF in the same directory.

    NOTE: LibreOffice conversion is not thread-safe--it will
    silently fail if simultaneous processes are launched--sometimes.
    For loop, as the command line globbing doesn't work either,
    even just from Terminal!
    """
    p = argparse.ArgumentParser()
    p.add_argument("path", help="path to work in")
    p.add_argument("-s", "--suffix", help="file extensions to convert", nargs="+")
    p.add_argument("-o", "--outdir", help="output directory")
    P = p.parse_args()

    if P.outdir:
        outdir = Path(P.outdir).expanduser().resolve()

    for f in docfinder(P.path, P.suffix, exclude=".pdf"):
        print(f"converting to PDF: {f}")
        doc2pdf(f, outdir)
