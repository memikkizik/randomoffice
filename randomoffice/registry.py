"""Parser registration and dispatch."""

from pathlib import Path
from typing import Iterable

from .contracts import DocumentParser, ParsedDocument
from .errors import UnsupportedFormatError


class ParserRegistry:
    """Dispatches files to parsers by normalized extension.

    Input: an iterable of parser instances.
    Output: a registry that returns ParsedDocument values from parse().
    """

    def __init__(self, parsers: Iterable[DocumentParser]) -> None:
        self._parsers = {parser.supported_format.lower(): parser for parser in parsers}

    def parse(self, source_path: Path) -> ParsedDocument:
        """Parse a file or raise an explicit unsupported-format error."""
        suffix = source_path.suffix.lower()
        parser = self._parsers.get(suffix)
        if parser is None:
            supported = ", ".join(sorted(self._parsers))
            raise UnsupportedFormatError(
                f"Unsupported extension {suffix or '<none>'!r}; supported: {supported}"
            )
        return parser.parse(source_path)


def default_registry() -> ParserRegistry:
    """Build the standard registry with local DOCX, XLSX, and PPTX parsers."""
    from .parsers import DocxParser, PptxParser, XlsxParser

    return ParserRegistry((DocxParser(), XlsxParser(), PptxParser()))
