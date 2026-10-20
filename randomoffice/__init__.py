"""RandomOffice enterprise document processing, conversion, and queue foundation."""

from .contracts import DocumentParser, ParsedDocument, SupportedFormat
from .converters import (
    ConversionPipeline,
    ConversionTarget,
    ConvertedOutput,
    DocxConverter,
    DocumentConverter,
    PptxConverter,
    XlsxConverter,
)
from .errors import DocumentParseError, ParserDependencyError, UnsupportedFormatError
from .performance import (
    CacheEntry,
    DocumentCache,
    FileSecurityGuard,
    SecurityLimitExceededError,
)
from .queue import (
    Job,
    JobQueueManager,
    TaskPriority,
    TaskStatus,
)
from .registry import ParserRegistry, default_registry
from .sync import (
    DirectorySyncManager,
    FileChange,
    SyncEvent,
)

__all__ = [
    # Contracts & Registry
    "DocumentParser",
    "DocumentParseError",
    "ParsedDocument",
    "ParserDependencyError",
    "ParserRegistry",
    "SupportedFormat",
    "UnsupportedFormatError",
    "default_registry",
    # Converters
    "ConversionPipeline",
    "ConversionTarget",
    "ConvertedOutput",
    "DocxConverter",
    "DocumentConverter",
    "PptxConverter",
    "XlsxConverter",
    # Queue
    "Job",
    "JobQueueManager",
    "TaskPriority",
    "TaskStatus",
    # Sync
    "DirectorySyncManager",
    "FileChange",
    "SyncEvent",
    # Performance & Security
    "CacheEntry",
    "DocumentCache",
    "FileSecurityGuard",
    "SecurityLimitExceededError",
]
