"""VoiceForge v2 configuration — one place for every timeout, budget and gate."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Tuple


@dataclass(frozen=True)
class VoiceForgeConfig:
    """All tunables for the engine. Construct once, pass everywhere."""

    business: str = "Acme Services"

    # -- LLM / turn latency budget (P0-6: Twilio webhook timeout ~15 s) ----
    llm_timeout_s: float = 8.0
    llm_max_tokens: int = 120
    llm_retries: int = 1            # one retry on 429/5xx/timeout, then fallback
    llm_retry_backoff_s: float = 1.0
    turn_window: int = 8            # how many recent turns feed the LLM
    turn_deadline_s: float = 12.0   # hard per-turn budget (logged, not killed)

    # -- vision pipeline budgets -------------------------------------------
    vision_timeout_s: float = 45.0  # hard cap for vision_analyze incl. retry
    vision_budget_s: float = 90.0    # end-to-end: photo request -> inject/fallback
    min_confidence: float = 0.6      # gate: below -> ask, never assert
    max_media_bytes: int = 5_000_000
    allowed_media_types: Tuple[str, ...] = (
        "image/jpeg", "image/png", "image/webp")
    orphan_media_ttl_s: float = 600.0  # 10 min park for early-arriving MMS

    # -- booking ------------------------------------------------------------
    booking_daily_cap: int = 20

    # -- warm transfer -------------------------------------------------------
    transfer_number: str = ""       # E.164 of the human; empty = transfer disabled
    transfer_timeout_s: int = 20

    # -- compliance ----------------------------------------------------------
    ai_disclosure: bool = True      # disclose AI identity in greeting (P0-5)

    @classmethod
    def from_env(cls, business: str = "Acme Services") -> "VoiceForgeConfig":
        """Build from environment. Secrets themselves live in the adapters."""
        return cls(
            business=business,
            transfer_number=os.environ.get("VOICEFORGE_TRANSFER_NUMBER", ""),
            ai_disclosure=os.environ.get("VOICEFORGE_AI_DISCLOSURE", "1") != "0",
        )
