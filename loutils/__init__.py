import shutil

from .find import docfinder
from .convert import doc2pdf
from .print import doc2print

LOEXE = shutil.which("libreoffice")
