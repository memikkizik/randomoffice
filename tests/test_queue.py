import time

from randomoffice.queue.manager import JobQueueManager
from randomoffice.queue.task import TaskStatus


def test_queue_executes_job_successfully() -> None:
    manager = JobQueueManager(num_workers=2)
    try:
        def add(a: int, b: int) -> int:
            return a + b

        job = manager.submit(add, 10, 20)
        completed_job = manager.wait_for_job(job.job_id, timeout=3.0)

        assert completed_job.status == TaskStatus.COMPLETED
        assert completed_job.result == 30
        assert completed_job.error is None
    finally:
        manager.shutdown(wait=True)


def test_queue_retries_on_failure() -> None:
    manager = JobQueueManager(num_workers=1)
    attempts = 0

    def flaky() -> str:
        nonlocal attempts
        attempts += 1
        if attempts < 3:
            raise ValueError(f"Attempt {attempts} failed")
        return "success"

    try:
        job = manager.submit(flaky, max_retries=3)
        completed = manager.wait_for_job(job.job_id, timeout=4.0)

        assert completed.status == TaskStatus.COMPLETED
        assert completed.result == "success"
        assert attempts == 3
    finally:
        manager.shutdown(wait=True)


def test_queue_marks_as_failed_after_exceeding_max_retries() -> None:
    manager = JobQueueManager(num_workers=1)

    def always_fail() -> None:
        raise RuntimeError("Fatal error")

    try:
        job = manager.submit(always_fail, max_retries=1)
        completed = manager.wait_for_job(job.job_id, timeout=3.0)

        assert completed.status == TaskStatus.FAILED
        assert "Fatal error" in (completed.error or "")
    finally:
        manager.shutdown(wait=True)


def test_queue_cancellation() -> None:
    manager = JobQueueManager(num_workers=1)
    try:
        # Submit a slow job first
        def slow_task() -> None:
            time.sleep(0.3)

        manager.submit(slow_task)

        # Submit a second job
        second = manager.submit(lambda: "never run")
        cancelled = manager.cancel(second.job_id)

        assert cancelled is True
        assert second.status == TaskStatus.CANCELLED
    finally:
        manager.shutdown(wait=True)
