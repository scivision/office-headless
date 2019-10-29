[![Actions Status](https://github.com/scivision/libreoffice-utils/workflows/ci_python/badge.svg)](https://github.com/scivision/libreoffice-utils/actions)
[![PyPi versions](https://img.shields.io/pypi/pyversions/loutils.svg)](https://pypi.python.org/pypi/loutils)
[![PyPi Download stats](http://pepy.tech/badge/loutils)](http://pepy.tech/project/loutils)

# LibreOffice Utils

Headless (command line) operations on Word / rich text documents for:

* Doc => PDF conversion  (LibreOffice only)
* printing (to the system default printer only)

from Python command line using LibreOffice or Microsoft Word.

These scripts use the command-line interface of LibreOffice or Microsoft Word in a headless (no GUI) fashion.

**CAUTION**
The doc2print.py script can print an unlimited number of pages to an unwanted printer, possibly causing great expense or violation of private documents to a public printer. Use great care with these scripts, preferably to a local non-networked printer you are sitting next to.

## LibreOffice

Since the command-line globbing of LibreOffice conversion is broken, this program provides a sane workaround for mass document conversion using LibreOffice.
LibreOffice is not thread-safe, so documents are converted or printed one at a time.

## Microsoft Word

Microsoft Word can be used to print to the system default printer.
On Windows, be sure that "Let Windows manage my default printer" is unchecked to avoid surprise default printer.
