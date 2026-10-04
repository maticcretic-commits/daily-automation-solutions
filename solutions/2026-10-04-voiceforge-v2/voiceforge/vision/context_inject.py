"""context_inject — put a gated VisionResult into the live call.

Rules:
  * idempotency FIRST (dedupe_key) — a duplicate is dropped, never re-spoken
  * never inject mid-utterance — deferred to the next turn boundary
  * exactly-once delivery, per-session lock
  * every failure funnels to ONE spoken voice-only fallback line; the caller
    never hears technical errors
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from ..logging import CallLogger
from ..session import CallSession, Turn
from .vision_analyze import VisionResult

VOICE_ONLY_FALLBACK = ("I couldn't get a clear look at the photo, so let's "
                       "continue by voice — could you describe what you see?")


@dataclass
class InjectOutcome:
    status: str  # "injected" | "deferred" | "dropped"
    reason: str = ""


class ContextInjector:
    def __init__(self, logger: CallLogger):
        self.logger = logger

    def inject(self, session: CallSession, result: VisionResult,
               dedupe_key: str, agent_speaking: bool = False
               ) -> InjectOutcome:
        """
        Deliver a vision result into the session. `agent_speaking` models
        the mid-utterance check: when True the result is parked as pending
        and flushed at the next turn boundary by flush_pending().

        N4: the 0.6 gate is enforced HERE, not by caller convention — a
        below-gate result is never asserted into the conversation. It is
        parked so the turn boundary can ASK (clarification question) instead.
        """
        with session.lock:
            # idempotency first
            if dedupe_key in session.vision_dedupe:
                return InjectOutcome("dropped", "duplicate")
            session.vision_dedupe.add(dedupe_key)

            if session.state != "active":
                # call already over: audit-log the result, never speak it
                self.logger.log("vision_inject_dropped", reason="call_ended",
                                gate_passed=result.gate_passed)
                return InjectOutcome("dropped", "call_ended")

            if not result.gate_passed:
                # ASK, never ASSERT: park for the turn boundary, which will
                # speak the clarification question (or the voice-only
                # fallback when the result itself failed).
                if session.pending_vision is None:
                    session.pending_vision = result
                else:
                    self.logger.log("vision_pending_superseded",
                                    reason="slot_busy")
                self.logger.log("vision_inject_deferred",
                                reason="gate_not_passed")
                return InjectOutcome("deferred", "gate_not_passed")

            if agent_speaking:
                if session.pending_vision is None:
                    session.pending_vision = result
                else:
                    # N6: first result wins — a second photo must not silently
                    # overwrite the first photo's pending result.
                    self.logger.log("vision_pending_superseded",
                                    reason="slot_busy")
                self.logger.log("vision_inject_deferred",
                                reason="mid_speech")
                return InjectOutcome("deferred", "mid_speech")

            session.turns.append(Turn(role="system", text=self._context_block(result)))
            self.logger.log("vision_injected",
                            gate_passed=result.gate_passed,
                            confidence=result.confidence)
            return InjectOutcome("injected", "ok")

    def flush_pending(self, session: CallSession) -> Optional[str]:
        """
        Called at the START of each new caller turn, before the brain runs.
        Returns a string for the agent to SPEAK (clarification question or
        the voice-only fallback), or None when a gated result was injected
        silently as a system turn for the LLM to use.

        N5: exactly-once comes from the pending slot itself (swapped to None
        atomically under the session lock) — no unstable id()-based keys.
        """
        with session.lock:
            result = session.pending_vision
            session.pending_vision = None
        if result is None:
            return None
        if not result.ok:
            self.logger.log("vision_pending_fallback")
            return VOICE_ONLY_FALLBACK
        if result.gate_passed:
            # turn boundary: the gate already passed, so the context block
            # goes in as a silent system turn for the LLM to use.
            with session.lock:
                if session.state == "active":
                    session.turns.append(
                        Turn(role="system",
                             text=self._context_block(result)))
                    self.logger.log("vision_injected", gate_passed=True,
                                    confidence=result.confidence)
                else:
                    self.logger.log("vision_inject_dropped",
                                    reason="call_ended", gate_passed=True)
            return None
        # gate failed: ASK, never assert — the question goes to the caller
        self.logger.log("vision_pending_clarify")
        return result.clarification_question or VOICE_ONLY_FALLBACK

    @staticmethod
    def _context_block(result: VisionResult) -> str:
        findings = "; ".join(result.findings)
        return (f"[vision: confidence={result.confidence:.2f}] "
                f"{result.description} Findings: {findings}")
