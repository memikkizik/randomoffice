"""DOCX parser backed by python-docx with structural ordering and fault tolerance."""

from pathlib import Path
from zipfile import BadZipFile

from ..contracts import ParsedDocument, SupportedFormat
from ..errors import DocumentParseError, ParserDependencyError
from ..performance.guard import FileSecurityGuard
from ._common import modified_at, validate_source


class DocxParser:
    """Extract paragraph and table text from a DOCX file preserving document order."""

    supported_format = SupportedFormat.DOCX

    def __init__(self, security_guard: FileSecurityGuard | None = None) -> None:
        self.security_guard = security_guard or FileSecurityGuard()

    def parse(self, source_path: Path) -> ParsedDocument:
        """Parse DOCX into normalized text and metadata."""
        validate_source(source_path, self.supported_format)
        self.security_guard.validate_office_archive(source_path)

        try:
            from docx import Document
            from docx.opc.exceptions import PackageNotFoundError
            from docx.table import Table
            from docx.text.paragraph import Paragraph
        except ImportError as exc:
            raise ParserDependencyError(
                "DOCX parsing requires python-docx; install project dependencies first"
            ) from exc

        try:
            document = Document(str(source_path))
            parts: list[str] = []

            # Preserve interleaving order of paragraphs and tables
            for child in document.element.body:
                if child.tag.endswith("p"):
                    para = Paragraph(child, document)
                    text = para.text.strip()
                    if text:
                        parts.append(text)
                elif child.tag.endswith("tbl"):
                    table = Table(child, document)
                    for row in table.rows:
                        row_cells: list[str] = []
                        seen_cells = set()
                        for cell in row.cells:
                            cell_id = id(cell._tc)
                            if cell_id not in seen_cells:
                                seen_cells.add(cell_id)
                                cell_text = cell.text.strip().replace("\n", " ")
                                if cell_text:
                                    row_cells.append(cell_text)
                        if row_cells:
                            parts.append(" | ".join(row_cells))

            text = "\n".join(parts)
            doc_title = (
                document.core_properties.title.strip()
                if document.core_properties and document.core_properties.title
                else None
            )
            title = doc_title or next((part for part in parts if part), None)

            return ParsedDocument(
                source_path=source_path,
                format=self.supported_format,
                title=title,
                text=text,
                modified_at=modified_at(source_path),
            )
        except (
            OSError,
            ValueError,
            RuntimeError,
            PackageNotFoundError,
            BadZipFile,
            KeyError,
            IndexError,
        ) as exc:
            raise DocumentParseError(f"Could not parse DOCX: {source_path}") from exc
