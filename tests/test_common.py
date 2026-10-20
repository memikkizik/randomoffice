from pathlib import Path

import pytest

from randomoffice.errors import DocumentParseError
from randomoffice.parsers._common import validate_source


def test_validate_source_rejects_missing_file(tmp_path: Path) -> None:
    with pytest.raises(DocumentParseError, match="does not exist"):
        validate_source(tmp_path / "missing.docx", ".docx")


def test_validate_source_rejects_wrong_extension(tmp_path: Path) -> None:
    source = tmp_path / "example.txt"
    source.touch()

    with pytest.raises(DocumentParseError, match="Expected .docx"):
        validate_source(source, ".docx")
