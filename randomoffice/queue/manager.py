"""Concurrent priority job queue manager."""

from datetime import datetime, timezone
import logging
from queue import Empty, PriorityQueue
import threading
import time
from typing import Any, Callable

from .task import Job, QueueItem, TaskPriority, TaskStatus

logger = logging.getLogger(__name__)


class JobQueueManager:
    """Manages thread-safe concurrent job execution with priorities and retries."""

    def __init__(self, num_workers: int = 4) -> None:
        self.num_workers = num_workers
        self._queue: PriorityQueue[QueueItem] = PriorityQueue()
        self._jobs: dict[str, Job] = {}
        self._lock = threading.RLock()
        self._workers: list[threading.Thread] = []
        self._shutdown_event = threading.Event()
        self._start_workers()

    def _start_workers(self) -> None:
        """Initialize worker threads."""
        for i in range(self.num_workers):
            thread = threading.Thread(
                target=self._worker_loop,
                name=f"RandomOffice-Worker-{i+1}",
                daemon=True,
            )
            thread.start()
            self._workers.append(thread)

    def submit(
        self,
        target: Callable[..., Any],
        *args: Any,
        priority: TaskPriority = TaskPriority.NORMAL,
        max_retries: int = 3,
        **kwargs: Any,
    ) -> Job:
        """Enqueue a new job for execution."""
        job = Job.create(
            target,
            *args,
            priority=priority,
            max_retries=max_retries,
            **kwargs,
        )
        with self._lock:
            self._jobs[job.job_id] = job
            self._queue.put(job.to_queue_item())
        return job

    def get_job(self, job_id: str) -> Job | None:
        """Retrieve job by its unique identifier."""
        with self._lock:
            return self._jobs.get(job_id)

    def cancel(self, job_id: str) -> bool:
        """Cancel a pending job if it has not already started."""
        with self._lock:
            job = self._jobs.get(job_id)
            if job and job.status == TaskStatus.PENDING:
                job.status = TaskStatus.CANCELLED
                job.finished_at = datetime.now(timezone.utc)
                return True
            return False

    def wait_for_job(self, job_id: str, timeout: float | None = None) -> Job:
        """Block until the specified job reaches a terminal state."""
        start_time = time.time()
        while True:
            with self._lock:
                job = self._jobs.get(job_id)
                if not job:
                    raise KeyError(f"Job {job_id} not found")
                if job.status in (TaskStatus.COMPLETED, TaskStatus.FAILED, TaskStatus.CANCELLED):
                    return job

            if timeout is not None and (time.time() - start_time) > timeout:
                raise TimeoutError(f"Timed out waiting for job {job_id}")
            time.sleep(0.05)

    def get_stats(self) -> dict[str, int]:
        """Return counts of jobs in each status."""
        with self._lock:
            stats = {status.value: 0 for status in TaskStatus}
            for job in self._jobs.values():
                stats[job.status.value] += 1
            stats["queued"] = self._queue.qsize()
            return stats

    def _worker_loop(self) -> None:
        """Internal worker processing loop."""
        while not self._shutdown_event.is_set():
            try:
                item = self._queue.get(timeout=0.2)
            except Empty:
                continue

            with self._lock:
                job = self._jobs.get(item.job_id)
                if not job or job.status == TaskStatus.CANCELLED:
                    self._queue.task_done()
                    continue
                job.status = TaskStatus.RUNNING
                job.started_at = datetime.now(timezone.utc)

            self._execute_job(job)
            self._queue.task_done()

    def _execute_job(self, job: Job) -> None:
        """Run the job target with retry logic."""
        try:
            result = job.target(*job.args, **job.kwargs)
            with self._lock:
                job.result = result
                job.status = TaskStatus.COMPLETED
                job.finished_at = datetime.now(timezone.utc)
        except Exception as exc:
            with self._lock:
                job.retry_count += 1
                if job.retry_count <= job.max_retries:
                    logger.warning(
                        "Job %s failed (attempt %d/%d): %s. Retrying...",
                        job.job_id,
                        job.retry_count,
                        job.max_retries,
                        exc,
                    )
                    job.status = TaskStatus.PENDING
                    self._queue.put(job.to_queue_item())
                else:
                    logger.error("Job %s permanently failed: %s", job.job_id, exc)
                    job.status = TaskStatus.FAILED
                    job.error = str(exc)
                    job.finished_at = datetime.now(timezone.utc)

    def shutdown(self, wait: bool = True) -> None:
        """Signal all workers to stop and optionally wait for termination."""
        self._shutdown_event.set()
        if wait:
            for worker in self._workers:
                worker.join(timeout=2.0)

