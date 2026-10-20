"""Shared parser validation helpers."""

from datetime import datetime
from pathlib import Path

from ..errors import DocumentParseError


def validate_source(source_path: Path, expected_suffix: str) -> None:
    """Validate parser input before invoking a third-party library."""
    if not isinstance(source_path, Path):
        raise DocumentParseError("source_path must be a pathlib.Path instance")
    if not source_path.is_file():
        raise DocumentParseError(f"Document does not exist or is not a file: {source_path}")
    if source_path.suffix.lower() != expected_suffix:
        raise DocumentParseError(
            f"Expected {expected_suffix} input, received {source_path.suffix or '<none>'}"
        )


def modified_at(source_path: Path) -> datetime:
    """Return the filesystem modification time as a timezone-naive datetime."""
    try:
        return datetime.fromtimestamp(source_path.stat().st_mtime)
    except OSError as exc:
        raise DocumentParseError(f"Could not read file metadata: {source_path}") from exc
