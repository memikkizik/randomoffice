"""Performance and security optimization module."""

from .cache import CacheEntry, DocumentCache
from .guard import FileSecurityGuard, SecurityLimitExceededError

__all__ = [
    "CacheEntry",
    "DocumentCache",
    "FileSecurityGuard",
    "SecurityLimitExceededError",
]

