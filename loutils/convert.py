from pathlib import Path
import subprocess
from . import LOEXE


def doc2pdf(filein: Path):
    if not LOEXE:
        raise SystemExit(f"{LOEXE} not found")

    cmd = [LOEXE, "--convert-to", "pdf", "--outdir", str(filein.parent), str(filein)]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
