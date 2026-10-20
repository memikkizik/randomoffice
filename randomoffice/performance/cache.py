"""In-memory and file-backed caching for parsed and converted documents."""

from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any


@dataclass(frozen=True, slots=True)
class CacheEntry:
    """Cached representation of an operation result."""
    key: str
    data: Any
    created_at: datetime


class DocumentCache:
    """Thread-safe LRU Cache for document operations based on file hashes."""

    def __init__(self, max_entries: int = 1000) -> None:
        self.max_entries = max_entries
        self._cache: OrderedDict[str, CacheEntry] = OrderedDict()

    def get(self, key: str) -> Any | None:
        """Retrieve cached result if present, updating LRU order."""
        if key not in self._cache:
            return None
        self._cache.move_to_end(key)
        return self._cache[key].data

    def set(self, key: str, value: Any) -> None:
        """Store result in cache, evicting oldest item if over capacity."""
        if key in self._cache:
            self._cache.move_to_end(key)
        self._cache[key] = CacheEntry(
            key=key,
            data=value,
            created_at=datetime.now(timezone.utc),
        )
        if len(self._cache) > self.max_entries:
            self._cache.popitem(last=False)

    def invalidate(self, key: str) -> bool:
        """Remove an entry from the cache."""
        if key in self._cache:
            del self._cache[key]
            return True
        return False

    def clear(self) -> None:
        """Purge all cached entries."""
        self._cache.clear()

    def __len__(self) -> int:
        return len(self._cache)

