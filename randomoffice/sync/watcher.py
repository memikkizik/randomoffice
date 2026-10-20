"""Background file synchronization and directory monitoring."""

from dataclasses import dataclass
from datetime import datetime, timezone
from enum import StrEnum
from pathlib import Path
import threading
import time
import logging
from typing import Callable

from ..contracts import SupportedFormat
from ..performance.guard import FileSecurityGuard

logger = logging.getLogger(__name__)


class SyncEvent(StrEnum):
    """File synchronization event types."""

    CREATED = "created"
    MODIFIED = "modified"
    DELETED = "deleted"


@dataclass(frozen=True, slots=True)
class FileChange:
    """Represents a detected change on a monitored document."""

    path: Path
    event_type: SyncEvent
    timestamp: datetime
    checksum: str | None = None


class DirectorySyncManager:
    """Monitors directories for Office document changes and triggers async workflows.
    
    Includes debouncing to prevent partial-read errors while Office or users
    are actively writing or saving files.
    """

    SUPPORTED_EXTENSIONS = {
        SupportedFormat.DOCX,
        SupportedFormat.XLSX,
        SupportedFormat.PPTX,
    }

    def __init__(
        self,
        watch_dir: Path,
        on_change: Callable[[FileChange], None],
        poll_interval: float = 1.0,
        debounce_seconds: float = 0.5,
    ) -> None:
        self.watch_dir = Path(watch_dir)
        self.on_change = on_change
        self.poll_interval = poll_interval
        self.debounce_seconds = debounce_seconds

        self._snapshot: dict[Path, tuple[float, int, str]] = {}  # path -> (mtime, size, hash)
        self._pending_writes: dict[Path, tuple[FileChange, float]] = {}
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    def start(self) -> None:
        """Start the background synchronization monitor thread."""
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._stop_event.clear()
            self._scan_initial()
            self._thread = threading.Thread(
                target=self._monitor_loop,
                name=f"SyncWatcher-{self.watch_dir.name}",
                daemon=True,
            )
            self._thread.start()

    def stop(self, wait: bool = True) -> None:
        """Stop background synchronization."""
        self._stop_event.set()
        if wait and self._thread:
            self._thread.join(timeout=3.0)

    def scan_once(self) -> list[FileChange]:
        """Perform a single immediate synchronization scan and return detected changes."""
        changes: list[FileChange] = []
        if not self.watch_dir.exists():
            return changes

        current_files: set[Path] = set()
        now = datetime.now(timezone.utc)

        for path in self.watch_dir.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
                continue

            # Ignore temporary Microsoft Office lock files (e.g., ~$document.docx)
            if path.name.startswith("~$"):
                continue

            current_files.add(path)
            try:
                stat = path.stat()
                mtime = stat.st_mtime
                size = stat.st_size
            except OSError:
                continue

            prev = self._snapshot.get(path)
            if prev is None:
                # Newly discovered file
                checksum = self._safe_compute_hash(path)
                self._snapshot[path] = (mtime, size, checksum)
                changes.append(FileChange(path, SyncEvent.CREATED, now, checksum))
            elif prev[0] != mtime or prev[1] != size:
                # File modified
                checksum = self._safe_compute_hash(path)
                if checksum != prev[2]:
                    self._snapshot[path] = (mtime, size, checksum)
                    changes.append(FileChange(path, SyncEvent.MODIFIED, now, checksum))

        # Check for deleted files
        deleted_paths = set(self._snapshot.keys()) - current_files
        for path in deleted_paths:
            del self._snapshot[path]
            changes.append(FileChange(path, SyncEvent.DELETED, now, None))

        return changes

    def _scan_initial(self) -> None:
        """Record baseline state without firing events."""
        if not self.watch_dir.exists():
            return
        for path in self.watch_dir.rglob("*"):
            if path.is_file() and path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                if not path.name.startswith("~$"):
                    try:
                        stat = path.stat()
                        checksum = self._safe_compute_hash(path)
                        self._snapshot[path] = (stat.st_mtime, stat.st_size, checksum)
                    except OSError:
                        pass

    def _monitor_loop(self) -> None:
        """Periodic polling and debounced event dispatch."""
        while not self._stop_event.is_set():
            try:
                changes = self.scan_once()
                for change in changes:
                    if change.event_type == SyncEvent.DELETED:
                        self._pending_writes.pop(change.path, None)
                        self._dispatch(change)
                    else:
                        self._pending_writes[change.path] = (change, time.monotonic())
                now = time.monotonic()
                ready = [
                    path for path, (_, queued_at) in self._pending_writes.items()
                    if now - queued_at >= self.debounce_seconds
                ]
                for path in ready:
                    change, _ = self._pending_writes.pop(path)
                    self._dispatch(change)
            except (OSError, PermissionError):
                logger.exception("Directory scan failed for %s", self.watch_dir)

            self._stop_event.wait(self.poll_interval)

    def _dispatch(self, change: FileChange) -> None:
        """Invoke the callback while preserving the monitor thread."""
        try:
            self.on_change(change)
        except Exception:
            logger.exception("Sync callback failed for %s", change.path)

    @staticmethod
    def _safe_compute_hash(path: Path) -> str:
        try:
            return FileSecurityGuard.compute_sha256(path)
        except (OSError, PermissionError):
            return ""
