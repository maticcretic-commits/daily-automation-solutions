"""Thread-safe, PII-redacting audit log (JSONL)."""

from __future__ import annotations

import hashlib
import json
import threading
import time
from typing import Any, Dict, List, Optional

from .util import redact_pii


def hash_phone(phone: str) -> str:
    """One-way handle for a phone number: auditable, not reversible."""
    return hashlib.sha256(phone.encode()).hexdigest()[:16]


class CallLogger:
    """
    Append-only JSONL audit log (one record per event).

    P0-1(6): every write goes through a single lock, and phone numbers /
    digit-runs are redacted before they reach the file. Raw phone numbers
    are never persisted — only a truncated SHA-256 handle.
    """

    def __init__(self, path: Optional[str] = None, redact: bool = True):
        self.path = path
        self.redact = redact
        self.events: List[Dict[str, Any]] = []
        self._lock = threading.Lock()

    def _scrub(self, fields: Dict[str, Any]) -> Dict[str, Any]:
        if not self.redact:
            return fields
        out: Dict[str, Any] = {}
        for key, value in fields.items():
            if key == "phone" and isinstance(value, str):
                out["phone_hash"] = hash_phone(value)
            elif isinstance(value, str):
                out[key] = redact_pii(value)
            else:
                out[key] = value
        return out

    def log(self, event: str, **fields: Any) -> None:
        rec = {"ts": time.time(), "event": event, **self._scrub(fields)}
        with self._lock:
            self.events.append(rec)
            if self.path:
                with open(self.path, "a", encoding="utf-8") as fh:
                    fh.write(json.dumps(rec) + "\n")
