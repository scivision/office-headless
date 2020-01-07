from pathlib import Path
import subprocess


def doc2print(filein: Path, exe: str):
    if exe == "libreoffice":
        cmd = ["libreoffice", "-p", str(filein)]
    elif exe == "winword":
        cmd = ["winword.exe", "/q", "/x", "/mFilePrintDefault", "/t", str(filein)]
    elif exe == "powerpoint":
        cmd = ["powerpnt.exe", "/P", str(filein)]
    elif exe == "wordpad":
        cmd = ["write.exe", "/p", str(filein)]
    elif exe == "notepad++":
        cmd = ["notepad++", "-quickPrint", str(filein)]
    elif exe == "acroread":
        cmd = ["C:/Program Files (x86)/Adobe/Acrobat Reader DC/Reader/acrord32.exe", "/p", str(filein)]
    else:
        raise SystemExit(f"unknown program {exe}")

    try:
        subprocess.check_call(cmd)
    except FileNotFoundError:
        raise SystemExit(f"{exe} not found")
