"""Call-session data model for VoiceForge v2 (voice + vision state)."""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, TYPE_CHECKING

if TYPE_CHECKING:  # avoid a runtime import cycle with vision modules
    from .vision.vision_analyze import VisionResult

# Vision lifecycle states for a call.
VISION_IDLE = "idle"
VISION_WAITING_MEDIA = "waiting_media"
VISION_ANALYZING = "analyzing"
VISION_READY = "ready"
VISION_UNAVAILABLE = "unavailable"
VISION_FAILED = "failed"


@dataclass
class Turn:
    """One spoken/written turn in a call."""
    role: str              # "caller" | "agent" | "system"
    text: str
    ts: float = field(default_factory=time.time)


@dataclass
class MediaItem:
    """One downloaded MMS attachment tied to a call."""
    media_sid: str         # "<MessageSid>_<index>"
    message_sid: str
    content_type: str
    size: int
    local_path: str
    received_ts: float = field(default_factory=time.time)


@dataclass
class CallSession:
    """State for one call, inbound or outbound (voice + vision)."""
    call_sid: str
    phone: str
    direction: str                      # "inbound" | "outbound"
    state: str = "ringing"              # ringing -> active -> ended
    turns: List[Turn] = field(default_factory=list)
    slots: Dict[str, str] = field(default_factory=dict)
    escalated: bool = False
    transfer_requested: bool = False
    transfer_done: bool = False
    started_at: float = field(default_factory=time.time)
    ended_at: Optional[float] = None
    # -- vision state (v2) ------------------------------------------------
    vision_state: str = VISION_IDLE
    media: List[MediaItem] = field(default_factory=list)
    pending_vision: Optional["VisionResult"] = None
    vision_dedupe: Set[str] = field(default_factory=set)
    vision_deadline_ts: Optional[float] = None  # 90 s budget anchor
    # -- concurrency -------------------------------------------------------
    lock: threading.Lock = field(default_factory=threading.Lock,
                                 repr=False, compare=False)

    def say(self, role: str, text: str) -> None:
        with self.lock:
            self.turns.append(Turn(role=role, text=text))

    def last_agent_text(self) -> str:
        with self.lock:
            for t in reversed(self.turns):
                if t.role == "agent":
                    return t.text
        return "Thank you for calling."

    def hangup(self, reason: str = "completed") -> None:
        with self.lock:
            self.state = "ended"
            self.ended_at = time.time()
            self.turns.append(Turn(role="system",
                                   text=f"[call ended: {reason}]"))
