"""Public input/output contracts for local document parsing."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Protocol


class SupportedFormat:
    """Supported Office document extensions."""

    DOCX = ".docx"
    XLSX = ".xlsx"
    PPTX = ".pptx"


@dataclass(frozen=True, slots=True)
class ParsedDocument:
    """Normalized parser output.

    Input: a successfully parsed local Office file.
    Output: immutable metadata and plain-text content suitable for indexing.
    """

    source_path: Path
    format: str
    title: str | None
    text: str
    modified_at: datetime | None


class DocumentParser(Protocol):
    """Parser interface implemented by each document adapter.

    Input: an existing path with the parser's supported extension.
    Output: a ParsedDocument.
    Raises: DocumentParseError for invalid input or parser failures.
    """

    supported_format: str

    def parse(self, source_path: Path) -> ParsedDocument:
        """Parse one local Office document."""
