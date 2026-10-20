"""Geometry and layout preservation engine for Word documents and PowerPoint slides.

Tracks spatial coordinates, preserves image aspect ratios, and provides
layout conversions between Office OpenXML coordinate systems (EMU/Inches/Points/Pixels).
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
import math
from typing import Any


EMU_PER_INCH = 914_400
POINTS_PER_INCH = 72.0
PX_PER_INCH = 96.0


def inches_to_emu(inches: float) -> int:
    """Convert inches to OpenXML English Metric Units (EMU)."""
    return int(round(inches * EMU_PER_INCH))


def emu_to_inches(emu: int | float) -> float:
    """Convert OpenXML English Metric Units (EMU) to inches."""
    return float(emu) / EMU_PER_INCH


def inches_to_pt(inches: float) -> float:
    """Convert inches to typographic points."""
    return float(inches) * POINTS_PER_INCH


def pt_to_inches(pt: float) -> float:
    """Convert typographic points to inches."""
    return float(pt) / POINTS_PER_INCH


def px_to_inches(px: float) -> float:
    """Convert screen pixels (standard 96 DPI) to inches."""
    return float(px) / PX_PER_INCH


@dataclass(frozen=True, slots=True)
class BoundingBox:
    """Spatial coordinates and dimensions for shapes, tables, and images."""

    x: float
    y: float
    width: float
    height: float
    unit: str = "inches"

    def __post_init__(self) -> None:
        if self.width < 0 or self.height < 0:
            raise ValueError(f"Width and height cannot be negative: {self.width}x{self.height}")

    @property
    def aspect_ratio(self) -> float:
        """Calculate width-to-height ratio. Returns 1.0 if height is zero."""
        if self.height <= 0:
            return 1.0
        return self.width / self.height

    def scale_to_fit(
        self,
        max_width: float,
        max_height: float,
        *,
        keep_aspect_ratio: bool = True,
    ) -> "BoundingBox":
        """Calculate new bounding box that fits within max bounds, preserving aspect ratio."""
        if max_width <= 0 or max_height <= 0:
            raise ValueError("Max bounds must be strictly positive")

        if not keep_aspect_ratio or self.height <= 0 or self.width <= 0:
            return BoundingBox(self.x, self.y, min(self.width, max_width), min(self.height, max_height), unit=self.unit)

        scale_w = max_width / self.width
        scale_h = max_height / self.height
        scale = min(scale_w, scale_h)

        new_width = round(self.width * scale, 4)
        new_height = round(self.height * scale, 4)
        return BoundingBox(self.x, self.y, new_width, new_height, unit=self.unit)

    def to_dict(self) -> dict[str, Any]:
        """Serialize coordinates for JSON transport."""
        return {
            "x": round(self.x, 4),
            "y": round(self.y, 4),
            "width": round(self.width, 4),
            "height": round(self.height, 4),
            "unit": self.unit,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "BoundingBox":
        """Deserialize coordinates from JSON structure."""
        return cls(
            x=float(data.get("x", 0.0)),
            y=float(data.get("y", 0.0)),
            width=float(data.get("width", 0.0)),
            height=float(data.get("height", 0.0)),
            unit=str(data.get("unit", "inches")),
        )


class SlideLayoutType(str, Enum):
    """Semantic PowerPoint slide layout classification."""

    TITLE_AND_CONTENT = "title_and_content"
    TITLE_ONLY = "title_only"
    TWO_COLUMN = "two_column"
    BLANK = "blank"
    PICTURE_WITH_CAPTION = "picture_with_caption"


def classify_slide_layout(
    *,
    has_title: bool,
    body_paragraphs_count: int,
    tables_count: int,
    images_count: int,
) -> SlideLayoutType:
    """Determine the optimal slide layout based on detected shapes."""
    if not has_title and body_paragraphs_count == 0 and tables_count == 0 and images_count == 0:
        return SlideLayoutType.BLANK

    if has_title and body_paragraphs_count == 0 and tables_count == 0 and images_count == 0:
        return SlideLayoutType.TITLE_ONLY

    if images_count >= 1 and (body_paragraphs_count <= 2 or tables_count <= 1):
        return SlideLayoutType.PICTURE_WITH_CAPTION

    if body_paragraphs_count >= 4 or (tables_count >= 1 and body_paragraphs_count >= 2):
        return SlideLayoutType.TWO_COLUMN

    return SlideLayoutType.TITLE_AND_CONTENT

