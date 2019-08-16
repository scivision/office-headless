from pathlib import Path
import subprocess


def doc2pdf(filein: Path):
    exe = "libreoffice"
    cmd = [exe, "--convert-to", "pdf", "--outdir", str(filein.parent), str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
