"""Exceptions used with pypptx.

The base exception class is |PyPptxError|; every exception pypptx raises deliberately derives
from it, so ``except PyPptxError`` catches them all.

Several classes also derive from a builtin exception. The API used to raise plain
``ValueError`` / ``TypeError`` in these cases, and callers written against that keep working:
``except ValueError`` still catches a |ShapeError|, |SlideError|, |ChartError|, |TableError|
or |InvalidValueError|, and ``except TypeError`` still catches an |InvalidTypeError|.
"""

from __future__ import annotations


class PyPptxError(Exception):
    """Generic error class."""


class PackageNotFoundError(PyPptxError):
    """Raised when a package cannot be found at the specified path."""


class InvalidXmlError(PyPptxError):
    """Raised when invalid XML is encountered according to the OOXML schema."""


class InvalidValueError(PyPptxError, ValueError):
    """Raised when a value is of an acceptable type but out of range or otherwise invalid.

    Also a ``ValueError``, which is what the API raised for these cases before.
    """


class InvalidTypeError(PyPptxError, TypeError):
    """Raised when a value is of the wrong type for the property or parameter it is given to.

    Also a ``TypeError``, which is what the API raised for these cases before.
    """


class ShapeError(PyPptxError, ValueError):
    """Raised for shape operation failures (type mismatch, invalid state).

    Also a ``ValueError``, which is what the API raised for these cases before.
    """


class SlideError(PyPptxError, ValueError):
    """Raised for slide collection and management failures.

    Also a ``ValueError``, which is what the API raised for these cases before.
    """


class ChartError(PyPptxError, ValueError):
    """Raised for chart-specific operation failures.

    Also a ``ValueError``, which is what the API raised for these cases before.
    """


class TableError(PyPptxError, ValueError):
    """Raised for table operation failures (cell access, merge).

    Also a ``ValueError``, which is what the API raised for these cases before.
    """


class PackageError(PyPptxError):
    """Raised for OPC package, relationship, or content-type errors."""


class InvalidPackageError(PackageError):
    """Raised when a file or stream exists but is not a readable `.pptx` package.

    For example, it is not a zip archive, the archive is truncated or corrupt, or it lacks the
    `[Content_Types].xml` every package must have. A path that does not exist raises
    |PackageNotFoundError| instead.
    """
