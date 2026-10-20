"""Production-ready Office converter implementations and pipeline orchestrator."""

import csv
import io
import json
from pathlib import Path
from typing import Any, Sequence

from ..contracts import SupportedFormat
from ..errors import DocumentParseError, UnsupportedFormatError
from ..performance.cache import DocumentCache
from ..performance.guard import FileSecurityGuard
from .contracts import ConversionTarget, ConvertedOutput, DocumentConverter


class DocxConverter:
    """High-fidelity converter for DOCX documents.
    
    Preserves document structure, headings, ordered/unordered lists,
    table structures, and metadata in Markdown, JSON, or Plain Text.
    """

    supported_format = SupportedFormat.DOCX

    def convert(
        self,
        source_path: Path,
        target: ConversionTarget,
        options: dict[str, Any] | None = None,
    ) -> ConvertedOutput:
        try:
            from docx import Document
            from docx.table import Table
            from docx.text.paragraph import Paragraph
        except ImportError as exc:
            raise DocumentParseError("python-docx is required for DOCX conversion") from exc

        try:
            doc = Document(str(source_path))
        except Exception as exc:
            raise DocumentParseError(f"Failed to read DOCX file: {source_path}") from exc

        metadata: dict[str, Any] = {
            "title": doc.core_properties.title or None,
            "author": doc.core_properties.author or None,
            "created": str(doc.core_properties.created) if doc.core_properties.created else None,
            "modified": str(doc.core_properties.modified) if doc.core_properties.modified else None,
        }

        # Extract elements in exact document flow order
        elements: list[dict[str, Any]] = []
        for child in doc.element.body:
            if child.tag.endswith("p"):
                para = Paragraph(child, doc)
                text = para.text.strip()
                if text:
                    style_name = para.style.name if para.style else "Normal"
                    elements.append({
                        "type": "paragraph",
                        "text": text,
                        "style": style_name,
                    })
            elif child.tag.endswith("tbl"):
                table = Table(child, doc)
                grid_rows: list[list[str]] = []
                for row in table.rows:
                    row_cells: list[str] = []
                    seen_cells = set()
                    for cell in row.cells:
                        # De-duplicate horizontal spans in python-docx
                        cell_id = id(cell._tc)
                        if cell_id not in seen_cells:
                            seen_cells.add(cell_id)
                            row_cells.append(cell.text.strip().replace("\n", " "))
                    if any(row_cells):
                        grid_rows.append(row_cells)
                if grid_rows:
                    elements.append({
                        "type": "table",
                        "rows": grid_rows,
                    })

        if target == ConversionTarget.JSON:
            content = json.dumps({"metadata": metadata, "elements": elements}, indent=2, ensure_ascii=False)
        elif target == ConversionTarget.MARKDOWN:
            content = self._render_markdown(metadata, elements)
        elif target == ConversionTarget.TEXT:
            content = self._render_text(elements)
        else:
            raise UnsupportedFormatError(f"Target format {target} not supported for DOCX")

        return ConvertedOutput(
            source_path=source_path,
            target_format=target,
            content=content,
            metadata=metadata,
        )

    def _render_markdown(self, metadata: dict[str, Any], elements: list[dict[str, Any]]) -> str:
        lines: list[str] = []
        if metadata.get("title"):
            lines.append(f"# {metadata['title']}\n")

        for el in elements:
            if el["type"] == "paragraph":
                style = el.get("style", "").lower()
                text = el["text"]
                if "heading 1" in style:
                    lines.append(f"# {text}\n")
                elif "heading 2" in style:
                    lines.append(f"## {text}\n")
                elif "heading 3" in style:
                    lines.append(f"### {text}\n")
                elif "bullet" in style or "list" in style:
                    lines.append(f"- {text}")
                else:
                    lines.append(f"{text}\n")
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue
                header = rows[0]
                lines.append("| " + " | ".join(header) + " |")
                lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                for row in rows[1:]:
                    padded_row = row + [""] * (len(header) - len(row))
                    lines.append("| " + " | ".join(padded_row[:len(header)]) + " |")
                lines.append("")

        return "\n".join(lines).strip()

    def _render_text(self, elements: list[dict[str, Any]]) -> str:
        lines: list[str] = []
        for el in elements:
            if el["type"] == "paragraph":
                lines.append(el["text"])
            elif el["type"] == "table":
                for row in el["rows"]:
                    lines.append("\t".join(row))
        return "\n".join(lines)


class XlsxConverter:
    """Converter for XLSX workbooks with streaming & tabular formatting."""

    supported_format = SupportedFormat.XLSX

    def convert(
        self,
        source_path: Path,
        target: ConversionTarget,
        options: dict[str, Any] | None = None,
    ) -> ConvertedOutput:
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise DocumentParseError("openpyxl is required for XLSX conversion") from exc

        wb = None
        try:
            wb = load_workbook(filename=source_path, read_only=True, data_only=True)
            metadata = {
                "title": wb.properties.title if hasattr(wb, "properties") and wb.properties else None,
                "sheets": wb.sheetnames,
            }

            sheet_data: dict[str, list[list[str]]] = {}
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                rows: list[list[str]] = []
                for row in ws.iter_rows(values_only=True):
                    row_vals = ["" if v is None else str(v).strip() for v in row]
                    if any(row_vals):
                        rows.append(row_vals)
                sheet_data[sheetname] = rows

        except Exception as exc:
            raise DocumentParseError(f"Failed to read XLSX file: {source_path}") from exc
        finally:
            if wb is not None:
                wb.close()

        if target == ConversionTarget.JSON:
            content = json.dumps({"metadata": metadata, "sheets": sheet_data}, indent=2, ensure_ascii=False)
        elif target == ConversionTarget.CSV:
            # First or selected sheet to CSV
            target_sheet = (options or {}).get("sheet") or (metadata["sheets"][0] if metadata["sheets"] else None)
            rows = sheet_data.get(target_sheet, []) if target_sheet else []
            s_buf = io.StringIO()
            writer = csv.writer(s_buf)
            for row in rows:
                writer.writerow(row)
            content = s_buf.getvalue()
        elif target == ConversionTarget.MARKDOWN:
            content = self._render_markdown(sheet_data)
        elif target == ConversionTarget.TEXT:
            content = self._render_text(sheet_data)
        else:
            raise UnsupportedFormatError(f"Target format {target} not supported for XLSX")

        return ConvertedOutput(
            source_path=source_path,
            target_format=target,
            content=content,
            metadata=metadata,
        )

    def _render_markdown(self, sheet_data: dict[str, list[list[str]]]) -> str:
        lines: list[str] = []
        for sheetname, rows in sheet_data.items():
            lines.append(f"## Sheet: {sheetname}\n")
            if not rows:
                lines.append("*(Empty sheet)*\n")
                continue
            header = rows[0]
            lines.append("| " + " | ".join(header) + " |")
            lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            for row in rows[1:]:
                padded = row + [""] * (len(header) - len(row))
                lines.append("| " + " | ".join(padded[:len(header)]) + " |")
            lines.append("")
        return "\n".join(lines).strip()

    def _render_text(self, sheet_data: dict[str, list[list[str]]]) -> str:
        lines: list[str] = []
        for sheetname, rows in sheet_data.items():
            lines.append(f"[{sheetname}]")
            for row in rows:
                lines.append("\t".join(row))
        return "\n".join(lines)


class PptxConverter:
    """Converter for PPTX presentations with shape, table, and notes extraction."""

    supported_format = SupportedFormat.PPTX

    def convert(
        self,
        source_path: Path,
        target: ConversionTarget,
        options: dict[str, Any] | None = None,
    ) -> ConvertedOutput:
        try:
            from pptx import Presentation
        except ImportError as exc:
            raise DocumentParseError("python-pptx is required for PPTX conversion") from exc

        try:
            prs = Presentation(str(source_path))
        except Exception as exc:
            raise DocumentParseError(f"Failed to read PPTX file: {source_path}") from exc

        metadata = {
            "title": prs.core_properties.title if hasattr(prs, "core_properties") else None,
            "slide_count": len(prs.slides),
        }

        slides_data: list[dict[str, Any]] = []
        for idx, slide in enumerate(prs.slides, 1):
            slide_info: dict[str, Any] = {
                "slide_number": idx,
                "title": None,
                "paragraphs": [],
                "tables": [],
                "notes": None,
            }

            # Check slide title
            if slide.shapes.title and slide.shapes.title.text:
                slide_info["title"] = slide.shapes.title.text.strip()

            # Recursive shape extraction
            self._extract_shapes(slide.shapes, slide_info)

            # Presenter notes
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()
                if notes_text:
                    slide_info["notes"] = notes_text

            slides_data.append(slide_info)

        if target == ConversionTarget.JSON:
            content = json.dumps({"metadata": metadata, "slides": slides_data}, indent=2, ensure_ascii=False)
        elif target == ConversionTarget.MARKDOWN:
            content = self._render_markdown(slides_data)
        elif target == ConversionTarget.TEXT:
            content = self._render_text(slides_data)
        else:
            raise UnsupportedFormatError(f"Target format {target} not supported for PPTX")

        return ConvertedOutput(
            source_path=source_path,
            target_format=target,
            content=content,
            metadata=metadata,
        )

    def _extract_shapes(self, shapes: Any, slide_info: dict[str, Any]) -> None:
        for shape in shapes:
            # Handle Group Shapes
            if hasattr(shape, "shapes"):
                self._extract_shapes(shape.shapes, slide_info)
                continue

            # Handle Tables
            if shape.has_table:
                tbl_rows: list[list[str]] = []
                for row in shape.table.rows:
                    row_vals = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                    if any(row_vals):
                        tbl_rows.append(row_vals)
                if tbl_rows:
                    slide_info["tables"].append(tbl_rows)
                continue

            # Handle Text Frames
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t and t != slide_info["title"]:
                        slide_info["paragraphs"].append(t)

    def _render_markdown(self, slides_data: list[dict[str, Any]]) -> str:
        lines: list[str] = []
        for slide in slides_data:
            num = slide["slide_number"]
            title = slide["title"] or f"Slide {num}"
            lines.append(f"## {title} (Slide {num})\n")

            for p in slide["paragraphs"]:
                lines.append(f"- {p}")

            for table in slide["tables"]:
                if table:
                    header = table[0]
                    lines.append("\n| " + " | ".join(header) + " |")
                    lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                    for row in table[1:]:
                        padded = row + [""] * (len(header) - len(row))
                        lines.append("| " + " | ".join(padded[:len(header)]) + " |")
                    lines.append("")

            if slide.get("notes"):
                lines.append(f"\n> **Notes:** {slide['notes']}\n")

            lines.append("\n---\n")
        return "\n".join(lines).strip()

    def _render_text(self, slides_data: list[dict[str, Any]]) -> str:
        lines: list[str] = []
        for slide in slides_data:
            lines.append(f"--- Slide {slide['slide_number']}: {slide['title'] or ''} ---")
            lines.extend(slide["paragraphs"])
            for tbl in slide["tables"]:
                for row in tbl:
                    lines.append("\t".join(row))
            if slide.get("notes"):
                lines.append(f"[Note] {slide['notes']}")
        return "\n".join(lines)


class ConversionPipeline:
    """Unified document conversion pipeline with caching and security guard."""

    def __init__(
        self,
        converters: Sequence[DocumentConverter] | None = None,
        security_guard: FileSecurityGuard | None = None,
        cache: DocumentCache | None = None,
    ) -> None:
        self.security_guard = security_guard or FileSecurityGuard()
        self.cache = cache or DocumentCache()
        
        default_converters: list[DocumentConverter] = [
            DocxConverter(),
            XlsxConverter(),
            PptxConverter(),
        ]
        chosen = converters if converters is not None else default_converters
        self._converters = {c.supported_format.lower(): c for c in chosen}

    def convert(
        self,
        source_path: Path,
        target: ConversionTarget,
        options: dict[str, Any] | None = None,
        use_cache: bool = True,
    ) -> ConvertedOutput:
        """Run the conversion pipeline with security validation and caching."""
        self.security_guard.validate_office_archive(source_path)

        suffix = source_path.suffix.lower()
        converter = self._converters.get(suffix)
        if not converter:
            raise UnsupportedFormatError(f"No converter registered for {suffix}")

        cache_key = ""
        if use_cache:
            file_hash = self.security_guard.compute_sha256(source_path)
            cache_key = f"{file_hash}:{target}:{options or {}}"
            cached = self.cache.get(cache_key)
            if cached is not None:
                return cached

        result = converter.convert(source_path, target, options)

        if use_cache and cache_key:
            self.cache.set(cache_key, result)

        return result

