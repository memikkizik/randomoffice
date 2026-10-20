"""Engine modules for spreadsheet dependencies, layout geometry, and loss inspection."""

from .dependency_graph import (
    DependencyGraph,
    FormulaCycleError,
    RangeExpansionLimitError,
    extract_references,
)
from .layout_geometry import BoundingBox, SlideLayoutType, classify_slide_layout
from .loss_reporter import DocumentLossAuditor

__all__ = [
    "DependencyGraph",
    "FormulaCycleError",
    "RangeExpansionLimitError",
    "extract_references",
    "BoundingBox",
    "SlideLayoutType",
    "classify_slide_layout",
    "DocumentLossAuditor",
]
