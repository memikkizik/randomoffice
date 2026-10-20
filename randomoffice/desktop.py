"""Native desktop runtime and a zero-socket Office document bridge."""

from __future__ import annotations

import base64
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import sys
import stat
import tempfile
from threading import RLock
from typing import Any, Callable

import webview

try:
    from .performance.guard import FileSecurityGuard
    from .web_office import export_office_file, open_office_file
    from .ui import render_office_ui
except (ImportError, ValueError):
    from randomoffice.performance.guard import FileSecurityGuard
    from randomoffice.web_office import export_office_file, open_office_file
    from randomoffice.ui import render_office_ui


_FORMATS = {"word": ".docx", "excel": ".xlsx", "powerpoint": ".pptx"}
_ALIASES = {"docx": "word", "xlsx": "excel", "sheet": "excel", "pptx": "powerpoint", "slides": "powerpoint"}
_MAX_WORKSPACE_STATE_BYTES = 16 * 1024 * 1024
_MAX_SOURCE_CACHE_BYTES = 32 * 1024 * 1024
_MAX_BASELINE_CACHE_BYTES = 16 * 1024 * 1024
_MAX_RECENT_DOCUMENTS = 20
_MAX_RECENT_STATE_BYTES = 1024 * 1024


class DocumentConflictError(ValueError):
    """The source or chosen destination changed since its native snapshot."""


@dataclass(frozen=True)
class FileSnapshot:
    sha256: str
    size: int
    mtime_ns: int

    def to_dict(self) -> dict[str, Any]:
        return {"sha256": self.sha256, "size": self.size, "mtime_ns": self.mtime_ns}


@dataclass
class SourceRecord:
    """Bounded in-memory preservation context; large native files stay on disk."""

    original_bytes: bytes | None
    baseline_editor_data: dict[str, Any] | None
    source_path: Path | None


def snapshot_file(source: Path) -> FileSnapshot:
    """Hash content as well as timestamps so same-size external edits are detected."""
    before = source.stat()
    digest = FileSecurityGuard.compute_sha256(source)
    after = source.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise DocumentConflictError("Dosya başka bir uygulamada değişiyor. Yeniden açıp tekrar deneyin.")
    return FileSnapshot(digest, after.st_size, after.st_mtime_ns)


def assert_file_snapshot(destination: Path, expected: FileSnapshot | None) -> None:
    try:
        current = snapshot_file(destination) if destination.exists() else None
    except OSError as exc:
        raise DocumentConflictError("Dosyanın güncel durumu doğrulanamadı. Yeniden açın veya farklı bir adla kaydedin.") from exc
    if current != expected:
        raise DocumentConflictError(
            "Dosya açıldıktan sonra başka bir uygulamada değiştirilmiş veya silinmiş. "
            "Değişiklikleri korumak için farklı bir adla kaydedin ya da dosyayı yeniden açın."
        )


def _bounded_baseline(data: dict[str, Any]) -> dict[str, Any] | None:
    if len(json.dumps(data, ensure_ascii=False).encode("utf-8")) > _MAX_BASELINE_CACHE_BYTES:
        return None
    return deepcopy(data)


def _same_path(first: Path, second: Path) -> bool:
    if os.path.normcase(str(first.resolve())) == os.path.normcase(str(second.resolve())):
        return True
    try:
        return first.samefile(second)
    except OSError:
        return False


def get_user_data_directory(directory: str | Path | None = None) -> Path:
    """Return stable, per-user storage, independent of the EXE extraction directory."""
    override = directory or os.environ.get("RANDOMOFFICE_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return (base / "RandomOffice").resolve()


@dataclass(frozen=True)
class DesktopWorkArea:
    """Monitor work area expressed in logical pixels, as required by pywebview."""

    width: int = 1280
    height: int = 800
    x: int = 0
    y: int = 0


def enable_dpi_awareness() -> None:
    """Request DPI scaling before obtaining screen dimensions or creating a window."""
    if sys.platform != "win32":
        return
    import ctypes

    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def detect_work_area() -> DesktopWorkArea:
    """Find the current monitor's usable area, excluding taskbar and DPI scaling."""
    if sys.platform != "win32":
        return DesktopWorkArea()
    import ctypes
    from ctypes import wintypes

    class MonitorInfo(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
        ]

    try:
        user32 = ctypes.windll.user32
        point = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(point))
        user32.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
        user32.MonitorFromPoint.restype = wintypes.HANDLE
        monitor = user32.MonitorFromPoint(point, 2)  # MONITOR_DEFAULTTONEAREST
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(info)
        user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
        if not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            return DesktopWorkArea()

        dpi = 96
        try:
            horizontal, vertical = wintypes.UINT(), wintypes.UINT()
            get_dpi = ctypes.windll.shcore.GetDpiForMonitor
            get_dpi.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.POINTER(wintypes.UINT), ctypes.POINTER(wintypes.UINT)]
            if get_dpi(monitor, 0, ctypes.byref(horizontal), ctypes.byref(vertical)) == 0:
                dpi = horizontal.value or 96
        except (AttributeError, OSError):
            try:
                dpi = user32.GetDpiForSystem() or 96
            except (AttributeError, OSError):
                pass
        scale = max(0.5, dpi / 96)
        rect = info.rcWork
        return DesktopWorkArea(
            width=max(1, round((rect.right - rect.left) / scale)),
            height=max(1, round((rect.bottom - rect.top) / scale)),
            x=round(rect.left / scale),
            y=round(rect.top / scale),
        )
    except (AttributeError, OSError, TypeError, ValueError):
        return DesktopWorkArea()


def adaptive_window_options(area: DesktopWorkArea) -> dict[str, Any]:
    """Keep restored and minimum window sizes inside even a small laptop screen."""
    width = min(1440, max(1, int(area.width * 0.94)))
    height = min(960, max(1, int(area.height * 0.94)))
    return {
        "width": width,
        "height": height,
        "x": area.x + (area.width - width) // 2,
        "y": area.y + (area.height - height) // 2,
        "min_size": (min(640, width), min(420, height)),
        "maximized": True,
    }


def locate_static_html() -> Path:
    """Find the shared UI in a source checkout, installation, or frozen bundle."""
    candidates = [Path(__file__).resolve().parent / "static" / "index.html"]
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        exe_dir = Path(sys.executable).parent
        candidates.extend([
            meipass / "randomoffice" / "static" / "index.html",
            meipass / "static" / "index.html",
            exe_dir / "_internal" / "randomoffice" / "static" / "index.html",
            exe_dir / "_internal" / "static" / "index.html",
            exe_dir / "randomoffice" / "static" / "index.html",
            exe_dir / "static" / "index.html",
        ])
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(f"Static UI 'index.html' not found. Checked: {[str(c) for c in candidates]}")


def _app_type(doc_type: str) -> str:
    value = str(doc_type).lower().lstrip(".")
    value = _ALIASES.get(value, value)
    if value not in _FORMATS:
        raise ValueError("Desteklenmeyen belge türü")
    return value


def _file_metadata(path: Path, snapshot: FileSnapshot | None = None) -> dict[str, Any]:
    try:
        read_only = not os.access(path, os.W_OK) or not bool(path.stat().st_mode & stat.S_IWRITE)
    except OSError:
        read_only = True
    return {
        "filename": path.name, "saved_path": str(path), "display_name": path.name,
        "read_only": read_only, "fingerprint": snapshot.to_dict() if snapshot is not None else None,
    }


def atomic_write_document(destination: Path, content: bytes, before_replace: Callable[[], None] | None = None) -> None:
    """Replace a document only after its complete export has reached the disk."""
    descriptor, name = tempfile.mkstemp(prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent)
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        if before_replace is not None:
            before_replace()
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


class NativeDesktopBridge:
    """In-process file operations; source identities stay on the native side."""

    def __init__(self, startup_document: str | Path | None = None, data_directory: str | Path | None = None) -> None:
        self._window: webview.Window | None = None
        self._startup_document = Path(startup_document) if startup_document is not None else None
        self._startup_consumed = False
        self._data_directory = get_user_data_directory(data_directory)
        self._current_files: dict[str, Path] = {}
        self._snapshots: dict[str, FileSnapshot] = {}
        self._sources: dict[str, SourceRecord] = {}
        self._unsaved: dict[str, bool] = {}
        self._lock = RLock()

    def set_window(self, window: webview.Window) -> None:
        self._window = window

    def is_desktop_mode(self) -> bool:
        return True

    def get_desktop_state(self) -> dict[str, Any]:
        with self._lock:
            return {
                "desktop": True,
                "current_files": {app: _file_metadata(path, self._snapshots.get(app)) for app, path in self._current_files.items()},
                "has_unsaved_changes": any(self._unsaved.values()),
                "startup_pending": self._startup_document is not None and not self._startup_consumed,
                "data_directory": str(self._data_directory),
            }

    def set_unsaved_state(self, doc_type: str, dirty: bool) -> dict[str, Any]:
        try:
            if not isinstance(dirty, bool):
                raise ValueError("Kaydedilmemiş değişiklik durumu doğru/yanlış olmalıdır")
            with self._lock:
                self._unsaved[_app_type(doc_type)] = dirty
                return {"success": True, "has_unsaved_changes": any(self._unsaved.values())}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def has_unsaved_changes(self) -> bool:
        with self._lock:
            return any(self._unsaved.values())

    def _confirm_close(self) -> bool:
        """Use native confirmation only when the editor reports unsaved changes."""
        with self._lock:
            dirty_apps = [app for app, dirty in self._unsaved.items() if dirty]
        if not dirty_apps:
            return True
        if self._window is None:
            return False
        labels = {"word": "Word belgesi", "excel": "Excel çalışma kitabı", "powerpoint": "PowerPoint sunumu"}
        documents = ", ".join(labels[app] for app in dirty_apps)
        try:
            return bool(self._window.create_confirmation_dialog(
                "Kaydedilmemiş değişiklikler",
                f"{documents} içinde kaydedilmemiş değişiklikler var.\n\n"
                "Kaydetmeden çıkmak için Tamam'a, düzenlemeye devam etmek için İptal'e basın.",
            ))
        except Exception:
            return False

    def get_windows_integration_state(self) -> dict[str, Any]:
        from .windows_integration import get_install_status
        return get_install_status()

    def open_default_apps_settings(self) -> dict[str, Any]:
        from .windows_integration import open_default_apps_settings
        return open_default_apps_settings()

    def _read_recent_documents(self) -> list[dict[str, str]]:
        path = self._data_directory / "recent-documents.json"
        if not path.exists():
            return []
        if path.stat().st_size > _MAX_RECENT_STATE_BYTES:
            raise ValueError("Son belgeler kaydı boyut sınırını aşıyor")
        records = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(records, list):
            raise ValueError("Son belgeler kaydı geçersiz")
        valid = []
        for record in records[:_MAX_RECENT_DOCUMENTS]:
            if not isinstance(record, dict) or any(not isinstance(record.get(key), str) for key in ("id", "path", "last_opened")):
                continue
            source = Path(record["path"])
            if not source.is_absolute() or source.suffix.lower() not in _FORMATS.values():
                continue
            app = next(app for app, extension in _FORMATS.items() if extension == source.suffix.lower())
            identifier = hashlib.sha256(os.path.normcase(str(source)).encode("utf-8")).hexdigest()
            if record["id"] != identifier:
                continue
            valid.append({"id": identifier, "filename": source.name, "path": str(source), "app": app, "last_opened": record["last_opened"]})
        return valid

    def _record_recent_document(self, source: Path, app: str) -> str | None:
        try:
            identifier = hashlib.sha256(os.path.normcase(str(source)).encode("utf-8")).hexdigest()
            records = [record for record in self._read_recent_documents() if record["id"] != identifier]
            records.insert(0, {
                "id": identifier, "filename": source.name, "path": str(source), "app": app,
                "last_opened": datetime.now(timezone.utc).isoformat(),
            })
            content = json.dumps(records[:_MAX_RECENT_DOCUMENTS], ensure_ascii=False).encode("utf-8")
            if len(content) > _MAX_RECENT_STATE_BYTES:
                raise ValueError("Son belgeler kaydı boyut sınırını aşıyor")
            self._data_directory.mkdir(parents=True, exist_ok=True)
            atomic_write_document(self._data_directory / "recent-documents.json", content)
            return None
        except Exception as exc:
            # Recent-history failure must not turn a successful document open/save into a failure.
            return f"Son belgeler listesi güncellenemedi: {exc}"

    def get_recent_documents(self) -> dict[str, Any]:
        with self._lock:
            try:
                return {"documents": [{**record, "exists": Path(record["path"]).is_file()} for record in self._read_recent_documents()]}
            except Exception as exc:
                return {"documents": [], "error": f"Son belgeler okunamadı: {exc}"}

    def open_recent_document(self, identifier: str) -> dict[str, Any]:
        """Resolve an opaque history ID instead of accepting a JavaScript file path."""
        with self._lock:
            try:
                record = next((record for record in self._read_recent_documents() if record["id"] == identifier), None)
                if record is None:
                    raise ValueError("Belge son belgeler listesinde bulunamadı")
                return self._load_path(Path(record["path"]))
            except Exception as exc:
                return {"error": f"Dosya açılamadı: {exc}"}

    def read_workspace_state(self) -> dict[str, Any]:
        """Read native drafts/preferences because an HTML WebView has an opaque origin."""
        try:
            with self._lock:
                source = self._data_directory / "workspace-state.json"
                if not source.exists():
                    return {"success": True, "state": {}}
                if source.stat().st_size > _MAX_WORKSPACE_STATE_BYTES:
                    raise ValueError("Yerel çalışma alanı kaydı 16 MB sınırını aşıyor")
                state = json.loads(source.read_text(encoding="utf-8"))
                if not isinstance(state, dict):
                    raise ValueError("Yerel çalışma alanı kaydı geçersiz")
                return {"success": True, "state": state}
        except Exception as exc:
            return {"success": False, "error": f"Yerel çalışma alanı okunamadı: {exc}"}

    def write_workspace_state(self, state: dict[str, Any]) -> dict[str, Any]:
        """Atomically persist a bounded workspace in the user's application directory."""
        try:
            if not isinstance(state, dict):
                raise ValueError("Çalışma alanı kaydı bir nesne olmalıdır")
            content = json.dumps(state, ensure_ascii=False, allow_nan=False).encode("utf-8")
            if len(content) > _MAX_WORKSPACE_STATE_BYTES:
                raise ValueError("Yerel çalışma alanı kaydı 16 MB sınırını aşıyor; belgenizi dosyaya kaydedin")
            with self._lock:
                self._data_directory.mkdir(parents=True, exist_ok=True)
                atomic_write_document(self._data_directory / "workspace-state.json", content)
            return {"success": True}
        except Exception as exc:
            return {"success": False, "error": f"Yerel çalışma alanı kaydedilemedi: {exc}"}

    def reset_current_file(self, doc_type: str) -> dict[str, Any]:
        """A new document must never overwrite the previously opened one."""
        try:
            with self._lock:
                app = _app_type(doc_type)
                self._current_files.pop(app, None)
                self._snapshots.pop(app, None)
                self._sources.pop(app, None)
                self._unsaved[app] = False
            return {"success": True, "current_files": self.get_desktop_state()["current_files"]}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def _load_path(self, source: Path) -> dict[str, Any]:
        chosen_path = source.expanduser().resolve()
        if chosen_path.suffix.lower() not in _FORMATS.values():
            raise ValueError("Yalnızca DOCX, XLSX ve PPTX belgeleri açılabilir")
        FileSecurityGuard().validate_office_archive(chosen_path)
        snapshot = snapshot_file(chosen_path)
        data = open_office_file(chosen_path)
        assert_file_snapshot(chosen_path, snapshot)
        app = _app_type(data["app"])
        cached = chosen_path.read_bytes() if snapshot.size <= _MAX_SOURCE_CACHE_BYTES else None
        if cached is not None and hashlib.sha256(cached).hexdigest() != snapshot.sha256:
            raise DocumentConflictError("Dosya yüklenirken değişti. Lütfen dosyayı yeniden açın.")
        self._current_files[app] = chosen_path
        self._snapshots[app] = snapshot
        self._sources[app] = SourceRecord(cached, _bounded_baseline(data), chosen_path)
        self._unsaved[app] = False
        metadata = _file_metadata(chosen_path, snapshot)
        data.update(metadata)
        data["desktop_file"] = metadata
        recent_warning = self._record_recent_document(chosen_path, app)
        if recent_warning:
            data["recent_warning"] = recent_warning
        return data

    def get_startup_document(self) -> dict[str, Any] | None:
        """Consume the optional file-open argument once, after the editor is ready."""
        with self._lock:
            if self._startup_consumed or self._startup_document is None:
                return None
            self._startup_consumed = True
            try:
                return self._load_path(self._startup_document)
            except Exception as exc:
                return {"error": f"Dosya açılamadı: {exc}"}

    def open_file_data(self, filename: str, encoded: str) -> dict[str, Any]:
        """Parse dropped content without treating a supplied filename as a path."""
        try:
            suffix = Path(filename).suffix.lower()
            if suffix not in _FORMATS.values():
                raise ValueError("Desteklenmeyen Office biçimi")
            guard = FileSecurityGuard()
            if len(encoded) > ((guard.max_file_size_bytes + 2) // 3) * 4:
                raise ValueError("Dosya yükleme sınırını aşıyor")
            raw = base64.b64decode(encoded, validate=True)
            if len(raw) > _MAX_SOURCE_CACHE_BYTES:
                raise ValueError("32 MB üzerindeki belgeleri içeriklerini korumak için Dosya Aç iletişim kutusundan açın")
            with tempfile.TemporaryDirectory(prefix="randomoffice-drop-") as directory:
                source = Path(directory) / ("document" + suffix)
                source.write_bytes(raw)
                guard.validate_office_archive(source)
                data = open_office_file(source)
            data["filename"] = Path(filename).name
            with self._lock:
                app = _app_type(data["app"])
                self._current_files.pop(app, None)
                self._snapshots.pop(app, None)
                self._sources[app] = SourceRecord(raw, _bounded_baseline(data), None)
                self._unsaved[app] = False
            return data
        except Exception as exc:
            return {"error": f"Dosya açılamadı: {exc}"}

    def open_native_file(self) -> dict[str, Any] | None:
        if self._window is None:
            return None
        try:
            selected = self._window.create_file_dialog(
                webview.FileDialog.OPEN,
                allow_multiple=False,
                file_types=(
                    "Office Belgeleri (*.docx;*.xlsx;*.pptx)",
                    "Word Belgeleri (*.docx)",
                    "Excel Çalışma Kitapları (*.xlsx)",
                    "PowerPoint Sunumları (*.pptx)",
                ),
            )
            if not selected:
                return None
            path = Path(selected if isinstance(selected, str) else selected[0])
            with self._lock:
                return self._load_path(path)
        except Exception as exc:
            return {"error": f"Dosya açılamadı: {exc}"}

    def _export_preserving_source(self, doc_type: str, payload: dict[str, Any]) -> bytes:
        source = self._sources.get(doc_type)
        if source is None:
            return export_office_file(doc_type, payload)
        original = source.original_bytes
        if original is None:
            if source.source_path is None:
                raise ValueError("Kaynak belge korunamadı. Dosya Aç ile belgeyi yeniden açın.")
            assert_file_snapshot(source.source_path, self._snapshots.get(doc_type))
            FileSecurityGuard().validate_file(source.source_path)
            original = source.source_path.read_bytes()
            if hashlib.sha256(original).hexdigest() != self._snapshots[doc_type].sha256:
                raise DocumentConflictError("Kaynak dosya okunurken değişti. Lütfen dosyayı yeniden açın.")
        baseline = source.baseline_editor_data
        if baseline is None:
            with tempfile.TemporaryDirectory(prefix="randomoffice-baseline-") as directory:
                path = Path(directory) / ("original" + _FORMATS[doc_type])
                path.write_bytes(original)
                baseline = open_office_file(path)
        from .source_preservation import export_preserving_source
        return export_preserving_source(doc_type, payload, original, baseline)

    def _save_to_path(self, doc_type: str, destination: Path, payload: dict[str, Any], *, allow_reconstruction: bool = False) -> dict[str, Any]:
        current_path = self._current_files.get(doc_type)
        expected = self._snapshots.get(doc_type) if current_path is not None and _same_path(destination, current_path) else (
            snapshot_file(destination) if destination.exists() else None
        )
        assert_file_snapshot(destination, expected)
        # Export first. A validation failure cannot truncate an existing file.
        reconstructed = False
        loss_report = []
        try:
            content = self._export_preserving_source(doc_type, payload)
        except Exception as exc:
            source = self._sources.get(doc_type)
            source_path = source.source_path if source is not None else current_path
            if not allow_reconstruction or getattr(exc, "requires_save_as", False) is not True:
                raise
            if source_path is not None and _same_path(destination, source_path):
                raise ValueError("Basitleştirilmiş kopya kaynak dosyanın üzerine kaydedilemez. Farklı bir dosya adı seçin.") from exc
            content = export_office_file(doc_type, payload)
            reconstructed = True
            loss_report = getattr(exc, "loss_report", [])
        baseline = _bounded_baseline(payload)
        atomic_write_document(destination, content, before_replace=lambda: assert_file_snapshot(destination, expected))
        snapshot = snapshot_file(destination)
        self._current_files[doc_type] = destination
        self._snapshots[doc_type] = snapshot
        self._sources[doc_type] = SourceRecord(
            content if len(content) <= _MAX_SOURCE_CACHE_BYTES else None, baseline, destination,
        )
        metadata = _file_metadata(destination, snapshot)
        result = {"success": True, **metadata, "desktop_file": metadata}
        if reconstructed:
            result.update({"reconstructed": True, "loss_report": loss_report})
        recent_warning = self._record_recent_document(destination, doc_type)
        if recent_warning:
            result["recent_warning"] = recent_warning
        return result

    @staticmethod
    def _save_error(exc: Exception) -> dict[str, Any]:
        if isinstance(exc, DocumentConflictError):
            return {"success": False, "conflict": True, "error": str(exc)}
        if callable(getattr(exc, "to_dict", None)):
            return {"success": False, "error": str(exc), **exc.to_dict()}
        return {"success": False, "error": str(exc)}

    def save_current_file(self, doc_type: str, default_filename: str, payload: dict[str, Any]) -> dict[str, Any]:
        """Save to the native identity; a fresh document receives a Save As dialog."""
        try:
            app = _app_type(doc_type)
            with self._lock:
                current_path = self._current_files.get(app)
                if current_path is not None:
                    return self._save_to_path(app, current_path, payload)
            return self.save_native_file(app, default_filename, payload)
        except Exception as exc:
            return self._save_error(exc)

    def save_native_file(self, doc_type: str, default_filename: str, payload: dict[str, Any], allow_reconstruction: bool = False) -> dict[str, Any]:
        """Choose a new destination and write atomically; cancellation preserves identity."""
        if self._window is None:
            return {"success": False, "error": "Pencere bulunamadı"}
        try:
            if not isinstance(allow_reconstruction, bool):
                raise ValueError("Basitleştirilmiş kopya tercihi doğru/yanlış olmalıdır")
            app = _app_type(doc_type)
            extension = _FORMATS[app]
            # Filename suggestions are display text, never authoritative source paths.
            suggested = Path(default_filename or "Belge").name
            if Path(suggested).suffix.lower() != extension:
                suggested = Path(suggested).stem + extension
            description = {"word": "Word Belgesi", "excel": "Excel Çalışma Kitabı", "powerpoint": "PowerPoint Sunumu"}[app]
            selected = self._window.create_file_dialog(
                webview.FileDialog.SAVE,
                save_filename=suggested,
                file_types=(f"{description} (*{extension})",),
            )
            if not selected:
                return {"success": False, "cancelled": True}
            destination = Path(selected if isinstance(selected, str) else selected[0]).expanduser()
            if not destination.suffix:
                destination = destination.with_suffix(extension)
            if destination.suffix.lower() != extension:
                raise ValueError(f"Bu belge {extension} uzantısıyla kaydedilmelidir")
            destination = destination.resolve()
            with self._lock:
                return self._save_to_path(app, destination, payload, allow_reconstruction=allow_reconstruction)
        except Exception as exc:
            return self._save_error(exc)

    def notify_editor_ready(self) -> None:
        """Signal readiness without embedding filenames or document content in JavaScript."""
        if self._window is not None:
            self._window.evaluate_js(
                "if (typeof window.ingestDesktopStartup === 'function') { void window.ingestDesktopStartup(); }"
            )

    def print_native(self) -> None:
        if self._window is not None:
            self._window.evaluate_js("window.print()")


def write_crash_log(error_text: str, data_directory: str | Path | None = None) -> Path:
    """Record startup failures where a normal user can always write."""
    directory = get_user_data_directory(data_directory) / "logs"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / "randomoffice-crash.log"
    with path.open("a", encoding="utf-8") as stream:
        stream.write(f"\n[{datetime.now(timezone.utc).isoformat()}]\n{error_text}\n")
    return path


def run_desktop_app(startup_document: str | Path | None = None, *, data_directory: str | Path | None = None) -> None:
    """Launch a maximized, persistent native window without opening an HTTP port."""
    try:
        enable_dpi_awareness()
        storage = get_user_data_directory(data_directory) / "WebView2"
        storage.mkdir(parents=True, exist_ok=True)
        bridge = NativeDesktopBridge(startup_document=startup_document, data_directory=data_directory)
        window = webview.create_window(
            title="RandomOffice — Masaüstü Ofis",
            html=render_office_ui(locate_static_html()),
            **adaptive_window_options(detect_work_area()),
            resizable=True,
            fullscreen=False,
            confirm_close=False,
            background_color="#f3f5f9",
            js_api=bridge,
        )
        bridge.set_window(window)
        window.events.loaded += bridge.notify_editor_ready
        window.events.closing += bridge._confirm_close
        webview.start(
            gui="edgechromium" if sys.platform == "win32" else None,
            debug=False,
            private_mode=False,
            storage_path=str(storage),
            localization={
                "global.quitConfirmation": "RandomOffice kapatılsın mı? Kaydedilmeyen değişiklikler kaybolabilir.",
                "global.ok": "Tamam",
                "global.cancel": "İptal",
            },
        )
    except Exception as exc:
        import traceback

        error_text = traceback.format_exc()
        try:
            log_message = f"Ayrıntılar: {write_crash_log(error_text, data_directory)}"
        except OSError:
            log_message = "Hata günlüğü yazılamadı."
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(
                    0, f"RandomOffice başlatılamadı:\n\n{exc}\n\n{log_message}", "RandomOffice Hata", 0x10,
                )
            except Exception:
                print(error_text, file=sys.stderr)
        else:
            print(error_text, file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    run_desktop_app(sys.argv[1] if len(sys.argv) > 1 else None)
