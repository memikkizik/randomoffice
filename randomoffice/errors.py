"""Explicit errors raised by the document processing layer."""


class DocumentParseError(Exception):
    """Base error for invalid documents and parser execution failures."""


class ParserDependencyError(DocumentParseError):
    """Raised when a parser's required third-party library is unavailable."""


class UnsupportedFormatError(DocumentParseError):
    """Raised when no registered parser supports a file extension."""
