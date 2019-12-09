from pathlib import Path
import subprocess
from . import LOEXE


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        if not LOEXE:
            raise SystemExit(f"{LOEXE} not found")

        cmd = [LOEXE, "-p", str(filein)]
    elif exe == "winword":
        # shutil.which doesn't work for winword
        cmd = [exe, "/q", "/x", "/mFilePrintDefault", "/t", str(filein)]
    else:
        raise SystemExit("only libreoffice or winword may be used")

    subprocess.check_call(cmd)
