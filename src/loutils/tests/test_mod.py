#!/usr/bin/env python
import pytest
import loutils


def test_pdf():
    with pytest.raises(ImportError):
        loutils.doc2pdf("")


if __name__ == "__main__":
    pytest.main([__file__])
