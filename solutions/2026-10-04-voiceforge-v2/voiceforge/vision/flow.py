"""VisionFlow — orchestrates the live-call vision step end to end.

Degradation ladder (in order):
  visual-confirmed -> visual-low-confidence (ask, never assert)
  -> visual-timeout (voice-only + one spoken note)
  -> landline/no-media (voice-only, no photo mention)

Overall budget: 90 s from the moment the agent asks for a photo.
"""

from __future__ import annotations

import time
from typing import Callable, Dict, List, Optional

from ..config import VoiceForgeConfig
from ..logging import CallLogger
from ..session import (CallSession, VISION_ANALYZING, VISION_FAILED,
                       VISION_READY, VISION_UNAVAILABLE, VISION_WAITING_MEDIA)
from .context_inject import ContextInjector, VOICE_ONLY_FALLBACK
from .mms_watch import MmsWatch
from .vision_analyze import VisionAnalyzer, VisionContext, VisionResult

Clock = Callable[[], float]


class VisionFlow:
    """
    Owns the vision lifecycle for calls: photo request -> media intake ->
    analysis -> injection, with the 90 s budget and the degradation ladder.

    Wiring: the deployer creates one VisionFlow per VoiceAgent, calls
    request_photo() when the conversation needs eyes, routes the Twilio
    Messaging webhook to on_mms_webhook(), and lets check_deadlines() run
    on every turn (VoiceAgent.caller_said does this automatically).
    """

    def __init__(self, mms_watch: MmsWatch, analyzer: VisionAnalyzer,
                 injector: ContextInjector, logger: CallLogger,
                 config: Optional[VoiceForgeConfig] = None,
                 clock: Clock = time.time):
        self.mms_watch = mms_watch
        self.analyzer = analyzer
        self.injector = injector
        self.logger = logger
        self.config = config or VoiceForgeConfig()
        self.clock = clock
        self._asked: Dict[str, str] = {}  # call_sid -> what was requested

    # -- entry: agent asks for a photo --------------------------------------
    def request_photo(self, session: CallSession, what: str) -> str:
        """
        Returns the spoken ask. Landlines skip the visual flow entirely;
        the voice-only path is then identical to v1 behavior.
        """
        if self.mms_watch.is_landline(session.phone):
            session.vision_state = VISION_UNAVAILABLE
            self.logger.log("vision_skipped", reason="landline")
            return ""
        session.vision_state = VISION_WAITING_MEDIA
        session.vision_deadline_ts = self.clock() + self.config.vision_budget_s
        self._asked[session.call_sid] = what
        self.logger.log("vision_requested", what=what)
        return (f"Could you send me a photo of {what} by text message "
                f"while we talk? I'll take a look right away.")

    # -- MMS webhook ---------------------------------------------------------
    def on_mms_webhook(self, agent: "VoiceAgent", url: str,  # noqa: F821
                       form: Dict[str, str], signature: str) -> bool:
        """
        Returns True when the webhook was valid (caller answers 200),
        False when the signature failed (caller must answer 403).
        """
        if not self.mms_watch.twilio.verify_signature(url, form, signature):
            self.logger.log("mms_rejected", reason="bad_signature")
            return False
        event = self.mms_watch.handle_webhook(url, form, signature)
        if event is None:
            return True  # valid signature, nothing actionable
        call_sid = self.mms_watch.attach_to_session(agent, event)
        if call_sid is None:
            return True  # orphaned; reaped when the call starts
        session = agent.calls[call_sid]
        self._analyze_attached(agent, session)
        return True

    def _analyze_attached(self, agent: "VoiceAgent",  # noqa: F821
                          session: CallSession) -> None:
        if not session.media:
            return
        if session.vision_state not in (VISION_WAITING_MEDIA, VISION_ANALYZING):
            return
        session.vision_state = VISION_ANALYZING
        item = session.media[-1]
        try:
            with open(item.local_path, "rb") as fh:
                image_bytes = fh.read()
        except OSError as exc:
            self.logger.log("vision_media_unreadable", error=str(exc))
            session.vision_state = VISION_FAILED
            return
        context = VisionContext(
            business=agent.brain.business,
            recent_turns=list(session.turns[-6:]),
            photo_request=self._asked.get(session.call_sid, "the issue"))
        result = self.analyzer.analyze(image_bytes, item.content_type,
                                       context)
        self._deliver(agent, session, result, item.media_sid)

    def _deliver(self, agent: "VoiceAgent", session: CallSession,  # noqa: F821
                 result: VisionResult, media_sid: str) -> None:
        if result.ok and result.gate_passed:
            session.vision_state = VISION_READY
        else:
            session.vision_state = VISION_FAILED
        # never mid-utterance from the async path: always defer to the
        # next turn boundary; flush_pending() speaks or injects it.
        outcome = self.injector.inject(
            session, result, dedupe_key=f"vision:{media_sid}",
            agent_speaking=True)
        self.logger.log("vision_delivered", status=outcome.status,
                        reason=outcome.reason,
                        gate_passed=result.gate_passed)

    # -- 90 s budget enforcement ----------------------------------------------
    def check_deadlines(self, agent: "VoiceAgent") -> List[str]:  # noqa: F821
        """
        Fire expired vision timers. Returns call_sids that fell back to
        voice-only (each gets the single spoken fallback line once).
        """
        now = self.clock()
        fell_back: List[str] = []
        for sid, session in agent.calls.items():
            if session.vision_state not in (VISION_WAITING_MEDIA,
                                            VISION_ANALYZING):
                continue
            if session.vision_deadline_ts and now >= session.vision_deadline_ts:
                session.vision_state = VISION_FAILED
                session.vision_deadline_ts = None
                # park the fallback so the NEXT turn speaks it exactly once
                with session.lock:
                    session.pending_vision = VisionResult(
                        ok=False, error="vision_timeout")
                self.logger.log("vision_timeout", call_sid=sid)
                fell_back.append(sid)
        return fell_back
