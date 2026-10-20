"""Production FastAPI application for RandomOffice document processing."""

from contextlib import asynccontextmanager
from datetime import datetime, timezone
import os
from pathlib import Path
import shutil
import tempfile
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from ..contracts import SupportedFormat
from ..converters.contracts import ConversionTarget
from ..converters.pipeline import ConversionPipeline
from ..errors import DocumentParseError, UnsupportedFormatError
from ..performance.cache import DocumentCache
from ..performance.guard import FileSecurityGuard, SecurityLimitExceededError
from ..queue.manager import JobQueueManager
from ..queue.task import TaskPriority, TaskStatus
from ..registry import default_registry

# Global singletons for production runtime
queue_manager: JobQueueManager | None = None
pipeline: ConversionPipeline | None = None
cache: DocumentCache | None = None
security_guard: FileSecurityGuard | None = None
start_time: datetime | None = None


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
        title="RandomOffice Enterprise API",
        version="0.1.0",
        description="Production microservice for parsing, converting, and queued processing of Microsoft Office documents.",
        lifespan=lifespan,
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

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

    class JobSubmitResponse(BaseModel):
        job_id: str
        status: str
        priority: str

    class JobStatusResponse(BaseModel):
        job_id: str
        status: str
        retry_count: int
        created_at: str
        started_at: str | None = None
        finished_at: str | None = None
        result: Any | None = None
        error: str | None = None

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

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = Path(tmp.name)
            try:
                shutil.copyfileobj(file.file, tmp)
                tmp.flush()
                tmp.close()

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
                if tmp_path.exists():
                    try:
                        tmp_path.unlink()
                    except OSError:
                        pass

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

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = Path(tmp.name)
            try:
                shutil.copyfileobj(file.file, tmp)
                tmp.flush()
                tmp.close()

                assert pipeline is not None
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
                if tmp_path.exists():
                    try:
                        tmp_path.unlink()
                    except OSError:
                        pass

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
