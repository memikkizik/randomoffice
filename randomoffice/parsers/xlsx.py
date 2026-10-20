"""XLSX parser backed by openpyxl with leak-safe resource cleanup."""

from pathlib import Path
from zipfile import BadZipFile

from ..contracts import ParsedDocument, SupportedFormat
from ..errors import DocumentParseError, ParserDependencyError
from ..performance.guard import FileSecurityGuard
from ._common import modified_at, validate_source


class XlsxParser:
    """Extract non-empty worksheet cells as tab-separated text."""

    supported_format = SupportedFormat.XLSX

    def __init__(self, security_guard: FileSecurityGuard | None = None) -> None:
        self.security_guard = security_guard or FileSecurityGuard()

    def parse(self, source_path: Path) -> ParsedDocument:
        """Parse XLSX into normalized text and metadata."""
        validate_source(source_path, self.supported_format)
        self.security_guard.validate_office_archive(source_path)

        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise ParserDependencyError(
                "XLSX parsing requires openpyxl; install project dependencies first"
            ) from exc

        workbook = None
        try:
            workbook = load_workbook(filename=source_path, read_only=True, data_only=True)
            prop_title = (
                workbook.properties.title.strip()
                if hasattr(workbook, "properties") and workbook.properties and workbook.properties.title
                else None
            )
            title = prop_title or (workbook.sheetnames[0] if workbook.sheetnames else None)

            lines: list[str] = []
            for worksheet in workbook.worksheets:
                lines.append(f"[{worksheet.title}]")
                for row in worksheet.iter_rows(values_only=True):
                    values = ["" if value is None else str(value).strip() for value in row]
                    if any(values):
                        lines.append("\t".join(values))

            text = "\n".join(lines)
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
            TypeError,
            KeyError,
            IndexError,
            BadZipFile,
        ) as exc:
            raise DocumentParseError(f"Could not parse XLSX: {source_path}") from exc
        finally:
            if workbook is not None:
                workbook.close()
