from pathlib import Path
import subprocess
import shutil

LOEXE = shutil.which("libreoffice")


def doc2pdf(filein: Path):
    if not LOEXE:
        raise SystemExit(f"LibreOffice not found")

    cmd = [LOEXE, "--convert-to", "pdf", "--outdir", str(filein.parent), str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
