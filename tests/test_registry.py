from pathlib import Path

import pytest

from randomoffice.contracts import ParsedDocument
from randomoffice.errors import UnsupportedFormatError
from randomoffice.registry import ParserRegistry


class StubParser:
    supported_format = ".stub"

    def parse(self, source_path: Path) -> ParsedDocument:
        return ParsedDocument(source_path, ".stub", "stub", "content", None)


def test_registry_dispatches_by_case_insensitive_extension(tmp_path: Path) -> None:
    source = tmp_path / "example.STUB"
    source.touch()

    result = ParserRegistry([StubParser()]).parse(source)

    assert result.text == "content"
    assert result.format == ".stub"


def test_registry_rejects_unknown_extension(tmp_path: Path) -> None:
    source = tmp_path / "example.pdf"
    source.touch()

    with pytest.raises(UnsupportedFormatError, match="Unsupported extension"):
        ParserRegistry([StubParser()]).parse(source)
