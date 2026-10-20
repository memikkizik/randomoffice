"""Local Office parser adapters."""

from .docx import DocxParser
from .pptx import PptxParser
from .xlsx import XlsxParser

__all__ = ["DocxParser", "PptxParser", "XlsxParser"]
