"""Task definitions and data structures for background job execution."""

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum, IntEnum
import time
from typing import Any, Callable
import uuid


class TaskPriority(IntEnum):
    """Job execution priority (lower integer indicates higher priority)."""

    URGENT = 0
    HIGH = 1
    NORMAL = 2
    LOW = 3


class TaskStatus(str, Enum):
    """Status lifecycle of a queue task."""

    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass(order=True)
class QueueItem:
    """Internal wrapper for prioritizing items in PriorityQueue."""

    priority: int
    timestamp: float
    job_id: str = field(compare=False)


@dataclass
class Job:
    """Represents a discrete executable unit of work."""

    job_id: str
    target: Callable[..., Any]
    args: tuple[Any, ...] = ()
    kwargs: dict[str, Any] = field(default_factory=dict)
    priority: TaskPriority = TaskPriority.NORMAL
    max_retries: int = 3
    retry_count: int = 0
    status: TaskStatus = TaskStatus.PENDING
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    started_at: datetime | None = None
    finished_at: datetime | None = None
    result: Any | None = None
    error: str | None = None

    @classmethod
    def create(
        cls,
        target: Callable[..., Any],
        *args: Any,
        priority: TaskPriority = TaskPriority.NORMAL,
        max_retries: int = 3,
        **kwargs: Any,
    ) -> "Job":
        """Factory method to construct a Job with a unique UUID."""
        return cls(
            job_id=str(uuid.uuid4()),
            target=target,
            args=args,
            kwargs=kwargs,
            priority=priority,
            max_retries=max_retries,
        )

    def to_queue_item(self) -> QueueItem:
        """Produce a comparable QueueItem."""
        return QueueItem(
            priority=int(self.priority),
            timestamp=time.time(),
            job_id=self.job_id,
        )

