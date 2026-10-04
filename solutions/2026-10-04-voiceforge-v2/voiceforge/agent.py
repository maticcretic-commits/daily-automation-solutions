"""VoiceAgent: the full call-lifecycle engine (voice + vision).

P0 fixes vs v1:
  * caller_said is idempotent: unknown SIDs are logged (no KeyError),
    retried webhooks on ended calls get the last reply (no ValueError,
    no Twilio retry storm).
  * every turn starts with flush_pending() (vision results surface at turn
    boundaries, never mid-utterance) and check_deadlines() (90 s budget).
  * escalation triggers a real warm transfer with context handoff and a
    clean fallback when the human does not answer.
"""

from __future__ import annotations

import threading
import time
from typing import Callable, Dict, List, Optional, Tuple

from .adapters_twilio import TwilioAdapter
from .brain import ConversationBrain
from .campaign import CampaignRunner
from .config import VoiceForgeConfig
from .logging import CallLogger
from .providers import TTSClient, MockTTSClient
from .session import CallSession
from .vision.context_inject import ContextInjector

# transfer_reachable(number) -> bool; deployer hook (carrier check, hours)
TransferCheck = Callable[[str], bool]


class VoiceAgent:
    """Ties telephony, brain, vision, TTS and logging into call flows."""

    def __init__(self, brain: ConversationBrain,
                 tts: Optional[TTSClient] = None,
                 logger: Optional[CallLogger] = None,
                 voice: str = "alloy",
                 config: Optional[VoiceForgeConfig] = None,
                 transfer_check: Optional[TransferCheck] = None,
                 vision_flow: Optional[object] = None):
        self.brain = brain
        self.tts = tts or MockTTSClient()
        self.logger = logger or CallLogger()
        self.voice = voice
        self.config = config or VoiceForgeConfig(business=brain.business)
        self.transfer_check = transfer_check
        self.vision_flow = vision_flow  # VisionFlow, set by the deployer
        self.calls: Dict[str, CallSession] = {}
        # B3: webhook threads mutate + iterate agent.calls concurrently;
        # every mutation/iteration goes through this lock (or a snapshot).
        self._calls_lock = threading.RLock()
        self._injector = ContextInjector(self.logger)

    def calls_snapshot(self) -> List[Tuple[str, CallSession]]:
        """Thread-safe copy of (call_sid, session) for iteration."""
        with self._calls_lock:
            return list(self.calls.items())

    # -- inbound ----------------------------------------------------------
    def inbound_call(self, call_sid: str, from_phone: str,
                     twilio: Optional[TwilioAdapter] = None) -> str:
        session = CallSession(call_sid=call_sid, phone=from_phone,
                              direction="inbound", state="active")
        with self._calls_lock:
            self.calls[call_sid] = session
        # late-arriving MMS may already be parked: try to attach it
        if self.vision_flow is not None:
            self.vision_flow.mms_watch.reap_orphans(self)
        greeting = self.brain.greet(session)
        audio = self.tts.speak(greeting, self.voice)
        self.logger.log("call_started", call_sid=call_sid, phone=from_phone,
                        direction="inbound")
        if twilio:
            return twilio.inbound_twiml(greeting)
        return greeting + f"\n[audio bytes: {len(audio)}]"

    def caller_said(self, call_sid: str, transcript: str,
                    twilio: Optional[TwilioAdapter] = None) -> str:
        """
        Idempotent per-turn handler. Unknown SIDs log and return a benign
        reply; ended sessions return the last agent reply (Twilio retries
        webhooks — this is what stops the retry storm).
        """
        started = time.time()
        # C1: Twilio may deliver an empty/missing transcript — never let a
        # None reach .lower() (500 on the voice webhook -> retry storm).
        transcript = transcript or ""
        session = self.calls.get(call_sid)
        if session is None:
            self.logger.log("turn_unknown_session", call_sid=call_sid)
            benign = "Sorry, I couldn't find that call. Goodbye."
            return twilio.continue_twiml(benign, end_call=True) \
                if twilio else benign
        if session.state != "active":
            last = session.last_agent_text()
            self.logger.log("turn_ended_session", call_sid=call_sid)
            return twilio.continue_twiml(last, end_call=True) \
                if twilio else last

        # vision: surface pending results at the turn boundary, then enforce
        # the 90 s budget before the brain runs
        vision_speech = self._injector.flush_pending(session)
        if self.vision_flow is not None:
            self.vision_flow.check_deadlines(self)
            # a deadline may have just parked the fallback line
            if vision_speech is None:
                vision_speech = self._injector.flush_pending(session)

        if vision_speech is not None:
            # B2: the caller utterance belongs in history even though
            # brain.handle() is skipped on vision-flush turns — without it,
            # LLM context, the handoff brief and the transcript lose what
            # the caller just said.
            session.say("caller", transcript)
            # clarification question or the single voice-only fallback line
            session.say("agent", vision_speech)
            audio = self.tts.speak(vision_speech, self.voice)
            self.logger.log("turn", call_sid=call_sid, caller=transcript,
                            reply=vision_speech, ended=False,
                            vision_spoken=True, audio_bytes=len(audio),
                            latency_s=round(time.time() - started, 2))
            return twilio.continue_twiml(vision_speech) \
                if twilio else vision_speech

        # B7: wire the vision step into the conversation loop — the brain
        # calls this when the caller offers a photo; without it, vision is
        # reachable only from unit tests, never from a real call.
        self.brain.request_photo_fn = (
            lambda what: self.request_photo(call_sid, what))
        reply = self.brain.handle(session, transcript)
        audio = self.tts.speak(reply, self.voice)
        ended = session.state == "ended"
        self.logger.log("turn", call_sid=call_sid, caller=transcript,
                        reply=reply, ended=ended,
                        escalated=session.escalated,
                        audio_bytes=len(audio),
                        latency_s=round(time.time() - started, 2))

        # warm transfer on escalation (P0-3)
        if session.transfer_requested and not session.transfer_done \
                and not ended:
            return self._warm_transfer(session, reply, twilio)
        if twilio:
            return twilio.continue_twiml(reply, end_call=ended)
        return reply

    # -- warm transfer ------------------------------------------------------
    def build_handoff_context(self, session: CallSession) -> str:
        """One-paragraph brief for the human agent receiving the transfer."""
        with session.lock:
            spoken = [t.text for t in session.turns
                      if t.role in ("caller", "agent")][-6:]
            slots = dict(session.slots)
            vision_bits = [
                t.text for t in session.turns
                if t.role == "system" and t.text.startswith("[vision:")]
        brief = " ".join(spoken[-4:]) if spoken else "no details yet"
        parts = [f"Incoming transfer from {self.brain.business}.",
                 f"Caller said: {brief}."]
        if slots.get("name"):
            parts.append(f"Name on file: {slots['name']}.")
        if vision_bits:
            parts.append(f"Photo analysis: {vision_bits[-1]}")
        return " ".join(parts)

    def _warm_transfer(self, session: CallSession, announcement: str,
                       twilio: Optional[TwilioAdapter]) -> str:
        session.transfer_done = True
        target = self.config.transfer_number
        # B5: a flaky deployer carrier-check must never 500 an escalation —
        # the #2 customer want cannot crash precisely when it is needed.
        try:
            reachable = (self.transfer_check(target) if self.transfer_check
                         else bool(target))
        except Exception as exc:  # noqa: BLE001 - fall back, never crash
            self.logger.log("transfer_check_failed", error=str(exc),
                            call_sid=session.call_sid)
            reachable = False
        if twilio and target and reachable:
            context = self.build_handoff_context(session)
            self.logger.log("transfer_started", call_sid=session.call_sid,
                            target=target)
            return twilio.transfer_twiml(
                announcement, target,
                whisper_url="/whisper?sid=" + session.call_sid,
                fallback_action_url="/transfer_fallback?sid="
                + session.call_sid,
                timeout_s=self.config.transfer_timeout_s)
        # fallback: no target configured or unreachable — take a message
        self.logger.log("transfer_fallback", call_sid=session.call_sid,
                        reason="no_target" if not target else "unreachable")
        fallback = ("I'm sorry, all of our teammates are busy right now. "
                    "Please tell me your name and number after the tone, "
                    "and we'll call you back within the hour.")
        session.say("agent", fallback)
        if twilio:
            return twilio.continue_twiml(fallback)
        return fallback

    # -- outbound campaign (backwards-compatible facade) --------------------
    def outbound_campaign(self, numbers_csv: str, opener: str,
                          throttle_per_min: int = 20,
                          dnc: Optional[List[str]] = None,
                          on_transcript=None,
                          sleep=time.sleep) -> List[Dict[str, str]]:
        runner = CampaignRunner(self.brain, tts=self.tts,
                                logger=self.logger)
        report = runner.run(numbers_csv, opener, throttle_per_min, dnc,
                            on_transcript, sleep)
        with self._calls_lock:
            self.calls.update(runner.sessions)
        return report

    # -- vision entry points --------------------------------------------------
    def request_photo(self, call_sid: str, what: str) -> str:
        """Ask the caller to text a photo; returns the spoken ask."""
        if self.vision_flow is None:
            return ""
        session = self.calls.get(call_sid)
        if session is None or session.state != "active":
            return ""
        ask = self.vision_flow.request_photo(session, what)
        if ask:
            session.say("agent", ask)
        return ask

    # -- reporting ----------------------------------------------------------
    def summary(self) -> Dict[str, object]:
        with self._calls_lock:
            calls = list(self.calls.values())
        return {
            "total_calls": len(calls),
            "inbound": sum(1 for c in calls if c.direction == "inbound"),
            "outbound": sum(1 for c in calls if c.direction == "outbound"),
            "escalated": sum(1 for c in calls if c.escalated),
            "transferred": sum(1 for c in calls if c.transfer_done),
            "ended": sum(1 for c in calls if c.state == "ended"),
            "total_turns": sum(len(c.turns) for c in calls),
            "log_events": len(self.logger.events),
        }

    def pilot_summary(self) -> Dict[str, object]:
        """
        Buyer-visible pilot metric behind the money-back guarantee:
        "every call answered and logged." Counts are computed from the
        session record only — no provider data needed.
        """
        with self._calls_lock:
            calls = list(self.calls.values())
        inbound = [c for c in calls if c.direction == "inbound"]
        answered = [c for c in inbound
                    if any(t.role == "agent" for t in c.turns)]
        booked = sum(1 for c in inbound if c.slots.get("day"))
        return {
            "calls_received": len(inbound),
            "calls_answered": len(answered),
            "answer_rate": (len(answered) / len(inbound)
                            if inbound else 1.0),
            "appointments_booked": booked,
            "calls_transferred": sum(1 for c in inbound
                                     if c.transfer_done),
            "guarantee_metric": "every call answered and logged",
            "guarantee_met": len(answered) == len(inbound),
        }

    def format_pilot_report(self) -> str:
        """Plain-language pilot report a business owner can read."""
        s = self.pilot_summary()
        pct = round(s["answer_rate"] * 100)
        lines = [
            f"Pilot report — {self.brain.business}",
            f"Calls received: {s['calls_received']}.",
            f"Calls answered by the AI receptionist: "
            f"{s['calls_answered']} ({pct}%).",
            f"Appointments booked: {s['appointments_booked']}.",
            f"Calls transferred to you: {s['calls_transferred']}.",
            ("Money-back metric — every call answered and logged: "
             f"{'MET' if s['guarantee_met'] else 'NOT MET'} "
             f"({s['calls_answered']}/{s['calls_received']})."),
        ]
        return "\n".join(lines)
