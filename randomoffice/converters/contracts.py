"""Contracts and interfaces for the document converter pipeline."""

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any, Protocol


class ConversionTarget(StrEnum):
    """Supported target conversion formats."""

    MARKDOWN = "markdown"
    JSON = "json"
    TEXT = "text"
    CSV = "csv"


@dataclass(frozen=True, slots=True)
class ConvertedOutput:
    """Result of a format conversion operation."""

    source_path: Path
    target_format: ConversionTarget
    content: str
    metadata: dict[str, Any]


class DocumentConverter(Protocol):
    """Interface for document format converters."""

    supported_format: str

    def convert(
        self,
        source_path: Path,
        target: ConversionTarget,
        options: dict[str, Any] | None = None,
    ) -> ConvertedOutput:
        """Convert an office document to the target format."""

