[![Build Status](https://travis-ci.com/scivision/libreoffice-utils.svg?branch=master)](https://travis-ci.com/scivision/libreoffice-utils)
[![PyPi versions](https://img.shields.io/pypi/pyversions/loutils.svg)](https://pypi.python.org/pypi/loutils)
[![PyPi Download stats](http://pepy.tech/badge/loutils)](http://pepy.tech/project/loutils)

# LibreOffice Utils

Word / rich text document conversion and printing from Python command line using LibreOffice.

These scripts use the command-line interface of LibreOffice in a headless (no GUI) fashion.

Because LibreOffice (at least through 6.2) is not threadsafe, we cannot yet use asyncio to launch numerous subprocesses for document conversion.
Further, since the command-line globbing of LibreOffice conversion is broken, this program provides a sane workaround for mass document conversion using LibreOffcie.
