from pathlib import Path
import subprocess
import shutil

LOEXE = shutil.which("libreoffice")


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        if not LOEXE:
            raise SystemExit(f"LibreOffice not found")

        cmd = [LOEXE, "-p", str(filein)]
    elif exe == "winword":
        # shutil.which doesn't work for winword
        cmd = ["winword.exe", "/q", "/x", "/mFilePrintDefault", "/t", str(filein)]
    elif exe == "wordpad":
        exe = shutil.which("write")
        if not exe:
            raise SystemExit("WordPad not found")
        cmd = [exe, "/p", str(filein)]
    else:
        raise SystemExit("only libreoffice or winword may be used")

    subprocess.check_call(cmd)
