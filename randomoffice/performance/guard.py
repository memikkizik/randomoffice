"""Security and performance guardrails for Office document processing."""

import hashlib
from pathlib import Path
import zipfile
from ..errors import DocumentParseError


class SecurityLimitExceededError(DocumentParseError):
    """Raised when a document violates security limits (e.g. zip bomb, file too large)."""


class FileSecurityGuard:
    """Validates files against size limits and zip bomb attacks.
    
    Protects worker processes from Out-Of-Memory (OOM) and DoS attacks
    caused by malicious or excessively large files.
    """

    def __init__(
        self,
        max_file_size_bytes: int = 150 * 1024 * 1024,  # 150 MB
        max_uncompressed_bytes: int = 500 * 1024 * 1024,  # 500 MB
        max_compression_ratio: float = 100.0,
        max_archive_entries: int = 10_000,
    ) -> None:
        self.max_file_size_bytes = max_file_size_bytes
        self.max_uncompressed_bytes = max_uncompressed_bytes
        self.max_compression_ratio = max_compression_ratio
        self.max_archive_entries = max_archive_entries

    def validate_file(self, source_path: Path) -> None:
        """Verify file existence and physical file size."""
        if not source_path.is_file():
            raise DocumentParseError(f"File not found or is not a regular file: {source_path}")

        file_size = source_path.stat().st_size
        if file_size > self.max_file_size_bytes:
            raise SecurityLimitExceededError(
                f"File size {file_size} bytes exceeds limit of {self.max_file_size_bytes} bytes"
            )

    def validate_office_archive(self, source_path: Path) -> None:
        """Inspect the zip metadata of an OpenXML archive to prevent decompression bombs."""
        self.validate_file(source_path)

        if not zipfile.is_zipfile(source_path):
            raise DocumentParseError(f"File is not a valid zip/OpenXML archive: {source_path}")

        try:
            with zipfile.ZipFile(source_path, "r") as zf:
                infos = zf.infolist()
                if len(infos) > self.max_archive_entries:
                    raise SecurityLimitExceededError(
                        f"Archive contains {len(infos)} entries, exceeding entry limit of {self.max_archive_entries}"
                    )

                total_compressed = sum(info.compress_size for info in infos)
                total_uncompressed = sum(info.file_size for info in infos)

                if total_uncompressed > self.max_uncompressed_bytes:
                    raise SecurityLimitExceededError(
                        f"Total uncompressed size {total_uncompressed} bytes exceeds limit of {self.max_uncompressed_bytes} bytes"
                    )

                if total_compressed > 0:
                    ratio = total_uncompressed / total_compressed
                    if ratio > self.max_compression_ratio:
                        raise SecurityLimitExceededError(
                            f"Compression ratio {ratio:.1f}:1 exceeds safety threshold of {self.max_compression_ratio}:1 (Zip bomb detected)"
                        )
        except zipfile.BadZipFile as exc:
            raise DocumentParseError(f"Corrupted zip archive: {source_path}") from exc

    @staticmethod
    def compute_sha256(source_path: Path, chunk_size: int = 65536) -> str:
        """Compute SHA-256 hash of a file efficiently using streaming chunks."""
        hasher = hashlib.sha256()
        with open(source_path, "rb") as f:
            while chunk := f.read(chunk_size):
                hasher.update(chunk)
        return hasher.hexdigest()

