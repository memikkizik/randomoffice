"""Public input/output contracts for local document parsing."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Literal, Protocol


class SupportedFormat:
    """Supported Office document extensions."""

    DOCX = ".docx"
    XLSX = ".xlsx"
    PPTX = ".pptx"


DocumentKind = Literal["word", "excel", "powerpoint"]


@dataclass(frozen=True, slots=True)
class LossReportEntry:
    """Describe one document feature that could not be represented exactly."""

    feature: str
    message: str
    severity: Literal["info", "warning", "error"] = "warning"

    def to_dict(self) -> dict[str, str]:
        """Return the JSON-compatible representation used by API and UI clients."""
        return {
            "feature": self.feature,
            "message": self.message,
            "severity": self.severity,
        }


@dataclass(frozen=True, slots=True)
class ExcelCell:
    """Typed Excel cell data with formula and display metadata kept separate."""

    value: Any = None
    formula: str | None = None
    value_type: Literal["blank", "text", "number", "boolean", "date", "formula"] = "blank"
    number_format: str | None = None
    style_ref: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return a stable JSON-compatible cell contract."""
        return {
            "value": self.value,
            "formula": self.formula,
            "value_type": self.value_type,
            "number_format": self.number_format,
            "style_ref": self.style_ref,
        }


@dataclass(frozen=True, slots=True)
class OfficeDocument:
    """Canonical envelope shared by parsers, API, and web adapters.

    Format-specific payloads remain in ``content`` during migration so existing
    editor clients can continue consuming ``items``, ``sheets``, or ``slides``.
    """

    schema_version: Literal["office-document/1"]
    app: DocumentKind
    format: str
    title: str
    content: dict[str, Any]
    source_name: str | None = None
    source_path: Path | None = None
    modified_at: datetime | None = None
    loss_report: tuple[LossReportEntry, ...] = ()
    source_fingerprint: str | None = None

    def to_editor_data(self) -> dict[str, Any]:
        """Serialize the canonical document without exposing local paths."""
        result = dict(self.content)
        result.setdefault("app", self.app)
        result.setdefault("format", self.format)
        result.setdefault("title", self.title)
        result["schema_version"] = self.schema_version
        if self.loss_report:
            result["loss_report"] = [entry.to_dict() for entry in self.loss_report]
        if self.source_fingerprint:
            result["source_fingerprint"] = self.source_fingerprint
        return result

    @classmethod
    def from_editor_data(
        cls,
        data: dict[str, Any],
        *,
        source_path: Path | None = None,
        modified_at: datetime | None = None,
    ) -> "OfficeDocument":
        """Build the canonical envelope from an existing editor payload."""
        if not isinstance(data, dict):
            raise TypeError("Office document data must be an object")

        app = data.get("app")
        if app not in ("word", "excel", "powerpoint"):
            raise ValueError("Office document app must be word, excel, or powerpoint")
        schema_version = data.get("schema_version")
        if schema_version not in (None, "office-document/1"):
            raise ValueError("Unsupported OfficeDocument schema version")
        expected_format = {
            "word": SupportedFormat.DOCX,
            "excel": SupportedFormat.XLSX,
            "powerpoint": SupportedFormat.PPTX,
        }[app]
        document_format = data.get("format") or expected_format
        if document_format != expected_format:
            raise ValueError(f"{app} documents must use {expected_format}")

        raw_loss_report = data.get("loss_report", [])
        if not isinstance(raw_loss_report, list):
            raise ValueError("loss_report must be a list")
        entries: list[LossReportEntry] = []
        for entry in raw_loss_report:
            if not isinstance(entry, dict):
                raise ValueError("loss_report entries must be objects")
            severity = entry.get("severity", "warning")
            if severity not in ("info", "warning", "error"):
                raise ValueError("loss_report severity is invalid")
            entries.append(
                LossReportEntry(
                    feature=str(entry.get("feature", "unknown")),
                    message=str(entry.get("message", "")),
                    severity=severity,
                )
            )

        content = dict(data)
        for key in ("schema_version", "app", "format", "title", "loss_report", "source_fingerprint"):
            content.pop(key, None)
        source_name = data.get("source_name")
        if source_name is not None and not isinstance(source_name, str):
            raise TypeError("source_name must be text or null")
        source_fingerprint = data.get("source_fingerprint")
        if source_fingerprint is not None and not isinstance(source_fingerprint, str):
            raise TypeError("source_fingerprint must be text or null")
        return cls(
            schema_version="office-document/1",
            app=app,
            format=document_format,
            title=str(data.get("title") or "Untitled"),
            content=content,
            source_name=source_name,
            source_path=source_path,
            modified_at=modified_at,
            loss_report=tuple(entries),
            source_fingerprint=source_fingerprint,
        )


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
