import shutil
from typing import List, Generator, Union
from pathlib import Path

# whatever you might like to convert
SUFFIXES = ['.doc', '.docx', '.odt', '.pages', '.pub', '.rtf',
            '.wps', '.wri']

EXE = shutil.which('libreoffice')
if not EXE:
    raise FileNotFoundError(f'LibreOffice not found')


def docfinder(path: Path,
              suffixes: List[str] = None,
              exclude: str = None) -> Union[List[Path], Generator[Path, None, None]]:

    path = Path(path).expanduser().resolve()

    if not suffixes:
        suffixes = SUFFIXES

    if path.is_file():
        return [path]
    elif path.is_dir():
        pass
    else:
        raise FileNotFoundError(f'{path} is not a path or file')

    if exclude:
        flist = (f for f in path.iterdir() if f.is_file() and
                 (f.suffix in suffixes and not f.with_suffix(exclude).is_file()))
    else:
        flist = (f for f in path.iterdir() if f.is_file() and
                 f.suffix in suffixes)

    return flist
