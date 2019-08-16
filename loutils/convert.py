from pathlib import Path
import subprocess
import shutil


def doc2pdf(filein: Path):
    exe = "libreoffice"
    if not shutil.which(exe):
        raise SystemExit(f"{exe} not found")

    cmd = [exe, "--convert-to", "pdf", "--outdir", str(filein.parent), str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
