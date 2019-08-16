import typing
from pathlib import Path

# whatever you might like to convert
SUFFIXES = [".doc", ".docx", ".odt", ".pages", ".pub", ".rtf", ".wps", ".wri"]


def docfinder(
    path: Path, suffixes: typing.Sequence[str] = None, exclude: str = None
) -> typing.Iterator[Path]:

    path = Path(path).expanduser().resolve()

    if not suffixes:
        suffixes = SUFFIXES

    if path.is_file():
        for f in [path]:
            yield f
    elif path.is_dir():
        pass
    else:
        raise FileNotFoundError(f"{path} is not a path or file")

    if exclude:
        for f in path.iterdir():
            if (
                f.is_file()
                and f.suffix in suffixes
                and not f.with_suffix(exclude).is_file()
            ):
                yield f
    else:
        for f in path.iterdir():
            if f.is_file() and f.suffix in suffixes:
                yield f
