"""Production FastAPI application for RandomOffice document processing and Web Suite."""

from contextlib import asynccontextmanager
from datetime import datetime, timezone
import io
import logging
import os
from pathlib import Path
import tempfile
from typing import Any
from urllib.parse import quote

from fastapi import Body, FastAPI, File, Form, HTTPException, Query, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from ..contracts import SupportedFormat
from ..converters.contracts import ConversionTarget
from ..converters.pipeline import ConversionPipeline
from ..errors import DocumentParseError, UnsupportedFormatError
from ..performance.cache import DocumentCache
from ..performance.guard import FileSecurityGuard, SecurityLimitExceededError
from ..queue.manager import JobQueueManager
from ..registry import default_registry
from ..web_office import export_office_file, open_office_file
from ..ui import render_office_ui

logger = logging.getLogger(__name__)

# Global singletons for production runtime
queue_manager: JobQueueManager | None = None
pipeline: ConversionPipeline | None = None
cache: DocumentCache | None = None
security_guard: FileSecurityGuard | None = None
start_time: datetime | None = None

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
DEFAULT_MAX_UPLOAD_BYTES = 150 * 1024 * 1024


def _allowed_origins() -> list[str]:
    """Read an explicit CORS allowlist; local browser use is the safe default."""
    configured = os.getenv("RANDOMOFFICE_ALLOWED_ORIGINS", "").strip()
    return [origin.strip() for origin in configured.split(",") if origin.strip()] or [
        "http://127.0.0.1:8000",
        "http://localhost:8000",
    ]


def _safe_download_name(filename: str, default: str) -> str:
    """Return a basename without control characters for Content-Disposition."""
    candidate = filename.replace("\\", "/").rsplit("/", 1)[-1]
    candidate = "".join(char for char in candidate if ord(char) >= 32 and ord(char) != 127).strip()
    return candidate or default


async def _save_upload(upload: UploadFile, suffix: str, guard: FileSecurityGuard) -> Path:
    """Stream an upload to a private temporary file under the configured limit."""
    maximum = int(os.getenv("RANDOMOFFICE_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES)))
    total = 0
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_path = Path(tmp.name)
    try:
        while chunk := await upload.read(1024 * 1024):
            total += len(chunk)
            if total > maximum:
                raise SecurityLimitExceededError(
                    f"Upload exceeds limit of {maximum} bytes"
                )
            tmp.write(chunk)
        tmp.flush()
        tmp.close()
        guard.validate_office_archive(tmp_path)
        return tmp_path
    except Exception:
        tmp.close()
        tmp_path.unlink(missing_ok=True)
        raise


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Lifecycle manager: initialize background queue and caches on startup, cleanup on shutdown."""
    global queue_manager, pipeline, cache, security_guard, start_time
    start_time = datetime.now(timezone.utc)
    security_guard = FileSecurityGuard()
    cache = DocumentCache(max_entries=5000)
    pipeline = ConversionPipeline(security_guard=security_guard, cache=cache)
    queue_manager = JobQueueManager(num_workers=int(os.getenv("WORKER_CONCURRENCY", "4")))
    yield
    if queue_manager:
        queue_manager.shutdown(wait=True)


def create_app() -> FastAPI:
    """Factory creating the configured FastAPI application."""
    app = FastAPI(
        title="RandomOffice Web Suite & API",
        version="0.2.0",
        description="Free, open-source web Microsoft Office alternative and document conversion platform.",
        lifespan=lifespan,
    )
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=_allowed_origins(),
        allow_credentials=False,
        allow_methods=["GET", "POST"],
        allow_headers=["Content-Type"],
    )

    @app.middleware("http")
    async def add_security_headers(request, call_next):
        """Add browser isolation headers without changing document payloads."""
        response = await call_next(request)
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        response.headers.setdefault("Referrer-Policy", "no-referrer")
        response.headers.setdefault(
            "Permissions-Policy",
            "camera=(), microphone=(), geolocation=(), payment=()",
        )
        return response

    class HealthResponse(BaseModel):
        status: str = "healthy"
        uptime_seconds: float
        queue_stats: dict[str, int]
        supported_formats: list[str] = [
            SupportedFormat.DOCX,
            SupportedFormat.XLSX,
            SupportedFormat.PPTX,
        ]

    class ParseResponse(BaseModel):
        filename: str
        format: str
        title: str | None
        text: str
        modified_at: str | None

    class ConvertResponse(BaseModel):
        filename: str
        target_format: str
        content: str
        metadata: dict[str, Any]

    class JobStatusResponse(BaseModel):
        job_id: str
        status: str
        retry_count: int
        created_at: str
        started_at: str | None = None
        finished_at: str | None = None
        result: Any | None = None
        error: str | None = None

    # -------------------------------------------------------------------------
    # Web Office Suite UI
    # -------------------------------------------------------------------------
    @app.get("/", response_class=HTMLResponse, tags=["Web Office Suite"])
    async def serve_office_suite():
        """Serve the interactive RandomOffice Web Suite (Word, Excel, PowerPoint editor)."""
        html_file = STATIC_DIR / "index.html"
        if not html_file.exists():
            return HTMLResponse("<h1>RandomOffice Web Suite UI missing</h1>", status_code=500)
        return HTMLResponse(render_office_ui(html_file))

    @app.get("/manifest.json", tags=["Web Office Suite"])
    async def serve_manifest():
        """Serve PWA Web App Manifest."""
        manifest_file = STATIC_DIR / "manifest.json"
        if not manifest_file.exists():
            raise HTTPException(status_code=404, detail="Manifest not found")
        return FileResponse(manifest_file, media_type="application/manifest+json")

    @app.get("/api/v1/download/desktop-app", tags=["Web Office Suite"])
    async def download_desktop_app():
        """Download the standalone native Windows desktop application package (.exe in zip)."""
        zip_file = STATIC_DIR / "downloads" / "RandomOffice-Windows.zip"
        if not zip_file.exists():
            raise HTTPException(status_code=404, detail="Desktop package not built yet.")
        return FileResponse(
            zip_file,
            media_type="application/zip",
            filename="RandomOffice-Windows-Desktop.zip",
        )

    @app.post("/api/v1/office/open", tags=["Web Office Suite"])
    async def open_office_document(file: UploadFile = File(...)):
        """Open any DOCX, XLSX, or PPTX file and return structured editor data."""
        suffix = Path(file.filename or "").suffix.lower()
        if suffix not in (SupportedFormat.DOCX, SupportedFormat.XLSX, SupportedFormat.PPTX):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Unsupported format '{suffix}'. Supported: .docx, .xlsx, .pptx",
            )

        tmp_path: Path | None = None
        try:
            if security_guard is None:
                raise HTTPException(status_code=503, detail="Service is still starting")
            tmp_path = await _save_upload(file, suffix, security_guard)
            return open_office_file(tmp_path)
        except (SecurityLimitExceededError, UnsupportedFormatError) as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc))
        except DocumentParseError as exc:
            raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc))
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

    @app.post("/api/v1/office/export", tags=["Web Office Suite"])
    async def export_office_document(
        doc_type: str = Query(..., description="Document type: word, excel, powerpoint"),
        filename: str = Query("document.docx", description="Output filename"),
        data: dict[str, Any] = Body(...),
    ):
        """Export editor state into real downloadable binary DOCX, XLSX, or PPTX."""
        try:
            content_bytes = export_office_file(doc_type, data)
        except (DocumentParseError, UnsupportedFormatError, ValueError, KeyError, TypeError) as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Failed to generate {doc_type} file: {exc}",
            )

        media_types = {
            "word": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "powerpoint": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }
        media_types["sheet"] = media_types["excel"]
        media_types["slides"] = media_types["powerpoint"]
        media_type = media_types.get(doc_type.lower(), "application/octet-stream")
        download_name = _safe_download_name(filename, "document.docx")

        return StreamingResponse(
            io.BytesIO(content_bytes),
            media_type=media_type,
            headers={
                "Content-Disposition": (
                    "attachment; filename*=UTF-8''" + quote(download_name, safe="")
                )
            },
        )

    # -------------------------------------------------------------------------
    # Core REST Endpoints
    # -------------------------------------------------------------------------
    @app.get("/health", response_model=HealthResponse, tags=["System"])
    async def health_check():
        """System health and operational metrics."""
        now = datetime.now(timezone.utc)
        uptime = (now - start_time).total_seconds() if start_time else 0.0
        stats = queue_manager.get_stats() if queue_manager else {}
        return HealthResponse(
            status="healthy",
            uptime_seconds=round(uptime, 2),
            queue_stats=stats,
        )

    @app.post("/api/v1/parse", response_model=ParseResponse, tags=["Document Processing"])
    async def parse_document(file: UploadFile = File(...)):
        """Directly parse an uploaded Office document into normalized text and metadata."""
        suffix = Path(file.filename or "").suffix.lower()
        if suffix not in (SupportedFormat.DOCX, SupportedFormat.XLSX, SupportedFormat.PPTX):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Unsupported file extension '{suffix}'. Supported: .docx, .xlsx, .pptx",
            )

        tmp_path: Path | None = None
        try:
            if security_guard is None:
                raise HTTPException(status_code=503, detail="Service is still starting")
            tmp_path = await _save_upload(file, suffix, security_guard)
            registry = default_registry()
            parsed = registry.parse(tmp_path)
            return ParseResponse(
                filename=file.filename or tmp_path.name,
                format=parsed.format,
                title=parsed.title,
                text=parsed.text,
                modified_at=parsed.modified_at.isoformat() if parsed.modified_at else None,
            )
        except (SecurityLimitExceededError, UnsupportedFormatError) as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc))
        except DocumentParseError as exc:
            raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc))
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

    @app.post("/api/v1/convert", response_model=ConvertResponse, tags=["Document Processing"])
    async def convert_document(
        file: UploadFile = File(...),
        target: ConversionTarget = Form(ConversionTarget.MARKDOWN),
    ):
        """Convert an uploaded Office document into Markdown, JSON, CSV, or Text."""
        suffix = Path(file.filename or "").suffix.lower()
        if suffix not in (SupportedFormat.DOCX, SupportedFormat.XLSX, SupportedFormat.PPTX):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Unsupported file extension '{suffix}'",
            )

        tmp_path: Path | None = None
        try:
            if security_guard is None or pipeline is None:
                raise HTTPException(status_code=503, detail="Service is still starting")
            tmp_path = await _save_upload(file, suffix, security_guard)
            output = pipeline.convert(tmp_path, target=target)
            return ConvertResponse(
                filename=file.filename or tmp_path.name,
                target_format=output.target_format.value,
                content=output.content,
                metadata=output.metadata,
            )
        except (SecurityLimitExceededError, UnsupportedFormatError) as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc))
        except DocumentParseError as exc:
            raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc))
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

    @app.get("/api/v1/jobs/{job_id}", response_model=JobStatusResponse, tags=["Queue"])
    async def get_job_status(job_id: str):
        """Retrieve background job status and result."""
        assert queue_manager is not None
        job = queue_manager.get_job(job_id)
        if not job:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=f"Job {job_id} not found")

        return JobStatusResponse(
            job_id=job.job_id,
            status=job.status.value,
            retry_count=job.retry_count,
            created_at=job.created_at.isoformat(),
            started_at=job.started_at.isoformat() if job.started_at else None,
            finished_at=job.finished_at.isoformat() if job.finished_at else None,
            result=job.result,
            error=job.error,
        )

    @app.post("/api/v1/jobs/{job_id}/cancel", tags=["Queue"])
    async def cancel_job(job_id: str):
        """Cancel a pending job."""
        assert queue_manager is not None
        cancelled = queue_manager.cancel(job_id)
        if not cancelled:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Job {job_id} could not be cancelled (already running, completed, or not found)",
            )
        return {"job_id": job_id, "status": "cancelled"}

    return app


app = create_app()
