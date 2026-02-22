"""OCR + table-structure reconstruction package.

Converts .jpg/.jpeg/.png/.tif/.tiff and PDF files into a
layout-preserving .xlsx that mimics jpgtoexcel.com output.
"""

from .extractor import extract  # noqa: F401

__version__ = "0.1.0"
