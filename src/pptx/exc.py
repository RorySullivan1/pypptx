"""Exceptions used with pypptx.

The base exception class is PyPptxError.
"""

from __future__ import annotations


class PyPptxError(Exception):
    """Generic error class."""


class PackageNotFoundError(PyPptxError):
    """Raised when a package cannot be found at the specified path."""


class InvalidXmlError(PyPptxError):
    """Raised when invalid XML is encountered according to the OOXML schema."""


class ShapeError(PyPptxError):
    """Raised for shape operation failures (type mismatch, invalid state)."""


class SlideError(PyPptxError):
    """Raised for slide collection and management failures."""


class ChartError(PyPptxError):
    """Raised for chart-specific operation failures."""


class TableError(PyPptxError):
    """Raised for table operation failures (cell access, merge)."""


class PackageError(PyPptxError):
    """Raised for OPC package, relationship, or content-type errors."""
