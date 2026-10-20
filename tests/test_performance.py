from pathlib import Path
import zipfile

import pytest

from randomoffice.errors import DocumentParseError
from randomoffice.performance.cache import DocumentCache
from randomoffice.performance.guard import FileSecurityGuard, SecurityLimitExceededError


def test_file_security_guard_validates_size(tmp_path: Path) -> None:
    guard = FileSecurityGuard(max_file_size_bytes=100)
    normal_file = tmp_path / "normal.txt"
    normal_file.write_bytes(b"hello world")
    guard.validate_file(normal_file)

    huge_file = tmp_path / "huge.txt"
    huge_file.write_bytes(b"x" * 200)
    with pytest.raises(SecurityLimitExceededError, match="exceeds limit"):
        guard.validate_file(huge_file)


def test_file_security_guard_detects_zip_bombs(tmp_path: Path) -> None:
    # Set low limits to test zip bomb trigger
    guard = FileSecurityGuard(
        max_uncompressed_bytes=1000,
        max_compression_ratio=5.0,
    )
    zip_path = tmp_path / "bomb.docx"

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 10,000 zeros compress to a few dozen bytes, exceeding 5:1 ratio and 1000 bytes
        zf.writestr("word/document.xml", b"0" * 10000)

    with pytest.raises(SecurityLimitExceededError):
        guard.validate_office_archive(zip_path)


def test_file_security_guard_rejects_corrupted_archive(tmp_path: Path) -> None:
    corrupt_file = tmp_path / "corrupt.docx"
    corrupt_file.write_bytes(b"not a real zip file header")
    guard = FileSecurityGuard()
    with pytest.raises(DocumentParseError, match="not a valid zip"):
        guard.validate_office_archive(corrupt_file)


def test_document_cache_lru_eviction() -> None:
    cache = DocumentCache(max_entries=2)
    cache.set("a", 1)
    cache.set("b", 2)
    assert cache.get("a") == 1
    assert cache.get("b") == 2

    # Inserting 3rd entry evicts 'a' if 'a' was least recently used,
    # but 'b' was accessed after 'a', so let's access 'a' then insert 'c'
    cache.get("a")
    cache.set("c", 3)
    # Now 'b' should be evicted
    assert cache.get("b") is None
    assert cache.get("a") == 1
    assert cache.get("c") == 3


def test_document_cache_invalidation_and_clear() -> None:
    cache = DocumentCache()
    cache.set("k1", "v1")
    assert cache.invalidate("k1") is True
    assert cache.get("k1") is None
    assert cache.invalidate("k1") is False

    cache.set("k2", "v2")
    cache.clear()
    assert len(cache) == 0
