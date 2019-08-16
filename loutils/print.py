from pathlib import Path
import subprocess
import shutil


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        if not shutil.which(exe):
            raise SystemExit(f"{exe} not found")

        cmd = [exe, "-p", str(filein)]
    elif exe == "winword":
        # shutil.which doesn't work for winword
        cmd = [exe, "/q", "/x", "/mFilePrintDefault", "/t", str(filein)]
    else:
        raise SystemExit("only libreoffice or winword may be used")

    subprocess.check_call(cmd)
