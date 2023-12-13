# Headless LibreOffice / Microsoft Office

[![Actions Status](https://github.com/scivision/office-headless/workflows/ci/badge.svg)](https://github.com/scivision/office-headless/actions)
[![PyPI versions](https://img.shields.io/pypi/pyversions/loutils.svg)](https://pypi.python.org/pypi/loutils)
[![PyPI Download stats](http://pepy.tech/badge/loutils)](http://pepy.tech/project/loutils)

Headless (command line) operations by LibreOffice or Microsoft Office on Word, Excel, PowerPoint and most other
[formats LibreOffice can handle](https://en.wikipedia.org/wiki/LibreOffice#Supported_file_formats)
for:

* Doc => PDF conversion  (LibreOffice only)
* printing (to the system default printer only)

from Python command line using LibreOffice or Microsoft Word

## standalone single file document to PDF

For reuse in other programs and projects, we made a
[separate standalone script doc2pdf.py](./doc2pdf.py)
to convert any document that LibreOffice can handle to PDF.

```sh
python doc2pdf.py ~/mydoc.docx
```

## .doc / .docx to PDF conversion

Convert a directory of .doc / .docx to .pdf by:

```sh
python -m loutils.doc2pdf ~/Documents
```

## Printing

CAUTION:

The doc2print.py script can print an unlimited number of pages to an unwanted printer, possibly causing great expense or violation of private documents to a public printer. Use great care with these scripts, preferably to a local non-networked printer you are sitting next to.

```sh
python -m loutils.doc2print ~/mydocs
```

The `-exe` parameter allows selecting the printing program.
The script does not check that the files can be printed appropriately, it just prints.
Thus use the `-s` parameter to select only the suffixes wanted.
For example to print all Markdown files in a directory with Notepad++:

```sh
python -m loutils.doc2print ~/mydocs -s .md -exe notepad++
```

LibreOffice 7.2 finally fixed file globbing, but we use explicit for-looping to work with older LibreOffice.
LibreOffice is not thread-safe, so documents are converted or printed one at a time.
