"""Exception types used across the application."""
from __future__ import annotations


class OTTError(Exception):
    """Base class for all application errors that should be shown to the user."""


class WorkbookError(OTTError):
    """The uploaded file is not a readable .xlsx workbook."""


class ColumnSpecError(OTTError, ValueError):
    """A column specification such as 'BK:BQ' could not be parsed."""


class MappingError(OTTError):
    """The question/column mapping is missing or invalid.

    ``errors`` holds one human-readable message per problem.
    """

    def __init__(self, errors):
        if isinstance(errors, str):
            errors = [errors]
        self.errors = list(errors)
        super().__init__("; ".join(self.errors))
