from pathlib import Path
import subprocess


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        cmd = [exe, "-p", str(filein)]
    elif exe == "winword":
        cmd = [exe, "/q", "/mFilePrintDefault", str(filein)]
    else:
        raise SystemExit("only libreoffice or winword may be used")

    subprocess.check_call(cmd)
