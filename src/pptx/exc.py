"""Exceptions used with pypptx.

The base exception class is PyPptxError.
"""

from __future__ import annotations


class PyPptxError(Exception):
    """Generic error class."""


# Backward compatibility alias
PythonPptxError = PyPptxError


class PackageNotFoundError(PyPptxError):
    """
    Raised when a package cannot be found at the specified path.
    """


class InvalidXmlError(PyPptxError):
    """
    Raised when a value is encountered in the XML that is not valid according
    to the schema.
    """
