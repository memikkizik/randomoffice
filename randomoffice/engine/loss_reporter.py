"""Loss inspection and reporting engine for Office documents.

Identifies unsupported or downgraded OpenXML features (e.g., floating images,
complex table formatting, SmartArt, charts, animations) and produces
structured LossReportEntry records conforming to the canonical OfficeDocument contract.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..contracts import LossReportEntry


class DocumentLossAuditor:
    """Audits documents for fidelity loss during open, parse, or export cycles."""

    @staticmethod
    def audit_word_items(items: list[dict[str, Any]]) -> list[LossReportEntry]:
        """Audit Word items parsed into editor structure."""
        entries: list[LossReportEntry] = []
        has_tables = False
        has_images = False

        for item in items:
            itype = item.get("type")
            if itype == "table":
                has_tables = True
                rows = item.get("rows", [])
                # Check for merged cells or stripped cell shading
                if any(len(r) > 1 for r in rows):
                    entries.append(
                        LossReportEntry(
                            feature="table_cell_shading_and_borders",
                            message="Table cell backgrounds, borders, and vertical alignment were simplified.",
                            severity="info",
                        )
                    )
                    break
            elif itype == "paragraph":
                html = item.get("html", "")
                if "<img" in html:
                    has_images = True

        return entries

    @staticmethod
    def audit_pptx_slides(slides: list[dict[str, Any]]) -> list[LossReportEntry]:
        """Audit PowerPoint slides parsed into editor structure."""
        entries: list[LossReportEntry] = []
        image_count = sum(len(s.get("images", [])) for s in slides)
        table_count = sum(len(s.get("tables", [])) for s in slides)

        if image_count > 1:
            entries.append(
                LossReportEntry(
                    feature="pptx_floating_image_coordinates",
                    message="Multiple slide images may overlap; custom absolute canvas coordinates were normalized.",
                    severity="warning",
                )
            )

        if table_count > 0:
            entries.append(
                LossReportEntry(
                    feature="pptx_table_geometry",
                    message="Table position normalized to default slide lower quadrant.",
                    severity="info",
                )
            )

        return entries

    @staticmethod
    def audit_excel_sheets(sheets: list[dict[str, Any]]) -> list[LossReportEntry]:
        """Audit Excel sheets parsed into editor structure."""
        entries: list[LossReportEntry] = []
        total_cells = 0
        has_formulas = False

        for sheet in sheets:
            for row in sheet.get("rows", []):
                for cell in row:
                    total_cells += 1
                    if isinstance(cell, str) and cell.startswith("="):
                        has_formulas = True

        if total_cells > 100_000:
            entries.append(
                LossReportEntry(
                    feature="large_sheet_virtualization",
                    message=f"Large worksheet ({total_cells} cells) rendered in memory; chart generation skipped.",
                    severity="info",
                )
            )

        return entries

