"""
VoiceForge - a working AI voice-agent engine for inbound & outbound calls.

Design goal: a portfolio-grade, dependency-free Python engine that mirrors the
real Vapi + Twilio stack a client would deploy:

    phone call -> telephony webhook (Twilio/Vapi) -> STT transcript
             -> conversation brain (intents + LLM) -> TTS reply
             -> action (book, escalate, hang up) -> audit log

Everything external (telephony, STT, LLM, TTS) is behind small provider
interfaces with *mock* implementations, so the full call lifecycle runs and
can be tested end-to-end with zero credentials and zero network access.
Real provider clients use only urllib from the standard library.

Run the self-contained simulation:

    python3 voice_agent.py --demo

Run the test suite:

    python3 test_voice_agent.py
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import hmac
import json
import re
import sys
import time
import urllib.request
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

# ---------------------------------------------------------------------------
# Turns, sessions, audit log
# ---------------------------------------------------------------------------

@dataclass
class Turn:
    """One spoken/written turn in a call."""
    role: str              # "caller" | "agent" | "system"
    text: str
    ts: float = field(default_factory=time.time)


@dataclass
class CallSession:
    """State for one call, inbound or outbound."""
    call_sid: str
    phone: str
    direction: str         # "inbound" | "outbound"
    state: str = "ringing"  # ringing -> active -> ended
    turns: List[Turn] = field(default_factory=list)
    slots: Dict[str, str] = field(default_factory=dict)  # captured data
    escalated: bool = False
    started_at: float = field(default_factory=time.time)
    ended_at: Optional[float] = None

    def say(self, role: str, text: str) -> None:
        self.turns.append(Turn(role=role, text=text))

    def hangup(self, reason: str = "completed") -> None:
        self.state = "ended"
        self.ended_at = time.time()
        self.turns.append(Turn(role="system", text=f"[call ended: {reason}]"))


class CallLogger:
    """Append-only JSONL audit log (one record per event)."""

    def __init__(self, path: Optional[str] = None):
        self.path = path
        self.events: List[Dict[str, Any]] = []

    def log(self, event: str, **fields: Any) -> None:
        rec = {"ts": time.time(), "event": event, **fields}
        self.events.append(rec)
        if self.path:
            with open(self.path, "a", encoding="utf-8") as fh:
                fh.write(json.dumps(rec) + "\n")

# ---------------------------------------------------------------------------
# Provider interfaces (STT / LLM / TTS) + mocks + real HTTP clients
# ---------------------------------------------------------------------------

class LLMClient:
    """Generate the agent's next reply given the call context."""

    def reply(self, session: CallSession, transcript_text: str) -> str:
        raise NotImplementedError


class MockLLMClient(LLMClient):
    """Deterministic stand-in so demos and tests run offline."""

    def __init__(self, script: Optional[Dict[str, str]] = None):
        # maps intent-name -> canned reply; matched via ConversationBrain
        self.script = script or {}

    def reply(self, session: CallSession, transcript_text: str) -> str:
        return self.script.get("fallback",
                               "Thanks, I've noted that. Is there anything else?")


class HttpLLMClient(LLMClient):
    """OpenAI-compatible chat-completions endpoint, stdlib-only."""

    def __init__(self, api_url: str, api_key: str, model: str = "gpt-4o-mini",
                 system_prompt: str = "You are a concise phone assistant."):
        self.api_url = api_url
        self.api_key = api_key
        self.model = model
        self.system_prompt = system_prompt

    def reply(self, session: CallSession, transcript_text: str) -> str:
        messages = [{"role": "system", "content": self.system_prompt}]
        for t in session.turns[-8:]:
            if t.role in ("caller", "agent"):
                messages.append({"role": "user" if t.role == "caller"
                                 else "assistant", "content": t.text})
        body = json.dumps({"model": self.model, "messages": messages,
                           "max_tokens": 120}).encode()
        req = urllib.request.Request(
            self.api_url, data=body,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {self.api_key}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode())
        return data["choices"][0]["message"]["content"].strip()


class TTSClient:
    """Render agent text to an audio payload (mocked)."""

    def speak(self, text: str, voice: str = "default") -> bytes:
        raise NotImplementedError


class MockTTSClient(TTSClient):
    """Pretends to synthesize; returns a tagged byte blob for tests/demos."""

    def speak(self, text: str, voice: str = "default") -> bytes:
        return f"[AUDIO voice={voice} chars={len(text)}] {text}".encode()

# ---------------------------------------------------------------------------
# Conversation brain: intents, booking flow, escalation
# ---------------------------------------------------------------------------

GREETING = ("Hi, this is the AI assistant for {business}. "
            "How can I help you today?")

ESCALATION_PHRASES = ("agent", "human", "representative", "manager",
                      "real person", "support")
BOOKING_PHRASES = ("book", "appointment", "schedule", "slot", "reschedule")
END_PHRASES = ("bye", "goodbye", "that's all", "no thanks", "thank you bye")

FAQ_PATTERNS: List[tuple] = [
    (re.compile(r"\b(hours?|open(ing)?|close|timing)\b", re.I),
     "We're open Monday to Saturday, 9 AM to 7 PM."),
    (re.compile(r"\b(price|cost|charge|fee|rate)\b", re.I),
     "Prices start from $49 and depend on the service. "
     "Would you like me to check a specific service?"),
    (re.compile(r"\b(where|location|address)\b", re.I),
     "We're located downtown, with parking right outside."),
    (re.compile(r"\b(service|offer|do you)\b", re.I),
     "We offer consultations, bookings, and priority support. "
     "What can I help you with?"),
]


class ConversationBrain:
    """
    Decides the agent's reply and next action for one caller utterance.

    Intent order: escalation -> goodbye -> booking -> FAQ -> LLM fallback.
    Booking is a 3-step slot-fill (name -> service -> day) driven by `slots`.
    """

    def __init__(self, business: str = "Acme Services",
                 llm: Optional[LLMClient] = None):
        self.business = business
        self.llm = llm or MockLLMClient()

    # -- public API ------------------------------------------------------
    def greet(self, session: CallSession) -> str:
        text = GREETING.format(business=self.business)
        session.say("agent", text)
        return text

    def handle(self, session: CallSession, caller_text: str) -> str:
        """Process one caller utterance; returns the agent's reply text."""
        session.say("caller", caller_text)
        lowered = caller_text.lower()

        if any(p in lowered for p in ESCALATION_PHRASES):
            session.escalated = True
            text = ("Of course — connecting you to a human teammate now. "
                    "Please hold for a moment.")
        elif self._in_booking(session) or any(p in lowered for p in BOOKING_PHRASES):
            text = self._booking_step(session, caller_text)
        elif any(p in lowered for p in END_PHRASES):
            text = "Thank you for calling. Have a great day!"
            session.say("agent", text)
            session.hangup(reason="caller goodbye")
            return text
        elif self._faq(caller_text) is not None:
            text = self._faq(caller_text)
        else:
            text = self.llm.reply(session, caller_text)

        session.say("agent", text)
        return text

    # -- booking slot-fill ------------------------------------------------
    def _in_booking(self, session: CallSession) -> bool:
        return session.slots.get("_booking_active") == "1"

    def _booking_step(self, session: CallSession, caller_text: str) -> str:
        slots = session.slots
        if not self._in_booking(session):
            slots["_booking_active"] = "1"
            slots["_booking_step"] = "name"
            return "Great, I can book that for you. What's your full name?"
        step = slots.get("_booking_step")
        if step == "name":
            slots["name"] = caller_text.strip()
            slots["_booking_step"] = "service"
            return (f"Thanks {slots['name']}. Which service would you like "
                    f"to book?")
        if step == "service":
            slots["service"] = caller_text.strip()
            slots["_booking_step"] = "day"
            return (f"{slots['service']} — got it. Which day works for you?")
        if step == "day":
            slots["day"] = caller_text.strip()
            slots.pop("_booking_active", None)
            slots.pop("_booking_step", None)
            return (f"You're booked, {slots['name']}: {slots['service']} on "
                    f"{slots['day']}. A confirmation will be sent by SMS.")
        return "Sorry, let's start over — what's your full name?"

    # -- FAQ --------------------------------------------------------------
    @staticmethod
    def _faq(caller_text: str) -> Optional[str]:
        for pattern, answer in FAQ_PATTERNS:
            if pattern.search(caller_text):
                return answer
        return None

# ---------------------------------------------------------------------------
# Telephony adapters: Twilio (signature + TwiML) and Vapi (server messages)
# ---------------------------------------------------------------------------

class TwilioAdapter:
    """
    Handles Twilio voice webhooks:
    - verifies X-Twilio-Signature (HMAC-SHA1 over URL + sorted params)
    - returns TwiML that says the reply and gathers the next utterance
    """

    def __init__(self, auth_token: str, gather_action_url: str = "/gather"):
        self.auth_token = auth_token
        self.gather_action_url = gather_action_url

    def verify_signature(self, url: str, params: Dict[str, str],
                         signature: str) -> bool:
        payload = url + "".join(k + params[k] for k in sorted(params))
        digest = hmac.new(self.auth_token.encode(), payload.encode(),
                          hashlib.sha1).digest()
        expected = base64.b64encode(digest).decode()
        return hmac.compare_digest(expected, signature)

    @staticmethod
    def _xml_escape(text: str) -> str:
        return (text.replace("&", "&amp;").replace("<", "&lt;")
                    .replace(">", "&gt;"))

    def inbound_twiml(self, greeting: str) -> str:
        return (f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Response><Say voice="alice">'
                f'{self._xml_escape(greeting)}'
                f'</Say><Gather input="speech" action="{self.gather_action_url}" '
                f'method="POST" speechTimeout="auto"/></Response>')

    def continue_twiml(self, reply: str, end_call: bool = False) -> str:
        say = (f'<Say voice="alice">{self._xml_escape(reply)}</Say>')
        if end_call:
            return (f'<?xml version="1.0" encoding="UTF-8"?>'
                    f'<Response>{say}<Hangup/></Response>')
        return (f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Response>{say}<Gather input="speech" '
                f'action="{self.gather_action_url}" method="POST" '
                f'speechTimeout="auto"/></Response>')


class VapiAdapter:
    """
    Handles Vapi server messages (webhook events). Verifies the
    X-Vapi-Secret header and turns assistant-request messages into
    conversation actions understood by the brain.
    """

    def __init__(self, webhook_secret: str):
        self.webhook_secret = webhook_secret

    def verify(self, headers: Dict[str, str]) -> bool:
        got = headers.get("x-vapi-secret", "")
        return hmac.compare_digest(got, self.webhook_secret)

    @staticmethod
    def parse_message(body: Dict[str, Any]) -> Dict[str, Any]:
        """Normalise a Vapi server message into a simple action dict."""
        msg_type = body.get("message", {}).get("type", "")
        if msg_type == "assistant-request":
            return {"action": "start_call",
                    "call_id": body.get("call", {}).get("id", "unknown")}
        if msg_type == "transcript" and body["message"].get("role") == "user":
            return {"action": "caller_said",
                    "call_id": body.get("call", {}).get("id", "unknown"),
                    "text": body["message"].get("transcript", "")}
        if msg_type == "end-of-call-report":
            return {"action": "call_ended",
                    "call_id": body.get("call", {}).get("id", "unknown"),
                    "summary": body.get("summary", "")}
        return {"action": "ignore", "type": msg_type}

# ---------------------------------------------------------------------------
# VoiceAgent: full call lifecycle engine
# ---------------------------------------------------------------------------

class VoiceAgent:
    """Ties telephony, brain, TTS and logging into runnable call flows."""

    def __init__(self, brain: ConversationBrain,
                 tts: Optional[TTSClient] = None,
                 logger: Optional[CallLogger] = None,
                 voice: str = "alloy"):
        self.brain = brain
        self.tts = tts or MockTTSClient()
        self.logger = logger or CallLogger()
        self.voice = voice
        self.calls: Dict[str, CallSession] = {}

    # -- inbound ----------------------------------------------------------
    def inbound_call(self, call_sid: str, from_phone: str,
                     twilio: Optional[TwilioAdapter] = None) -> str:
        session = CallSession(call_sid=call_sid, phone=from_phone,
                              direction="inbound", state="active")
        self.calls[call_sid] = session
        greeting = self.brain.greet(session)
        audio = self.tts.speak(greeting, self.voice)
        self.logger.log("call_started", call_sid=call_sid, phone=from_phone,
                        direction="inbound")
        if twilio:
            return twilio.inbound_twiml(greeting)
        return greeting + f"\n[audio bytes: {len(audio)}]"

    def caller_said(self, call_sid: str, transcript: str,
                    twilio: Optional[TwilioAdapter] = None) -> str:
        session = self.calls[call_sid]
        if session.state != "active":
            raise ValueError(f"call {call_sid} is not active")
        reply = self.brain.handle(session, transcript)
        audio = self.tts.speak(reply, self.voice)
        ended = session.state == "ended"
        self.logger.log("turn", call_sid=call_sid, caller=transcript,
                        reply=reply, ended=ended, escalated=session.escalated,
                        audio_bytes=len(audio))
        if twilio:
            return twilio.continue_twiml(reply, end_call=ended)
        return reply

    # -- outbound campaign -------------------------------------------------
    def outbound_campaign(self, numbers_csv: str, opener: str,
                          throttle_per_min: int = 20,
                          dnc: Optional[List[str]] = None,
                          on_transcript: Optional[Callable[[str], str]] = None
                          ) -> List[Dict[str, Any]]:
        """
        Dial through a CSV of numbers (columns: phone,name). Respects the
        do-not-call list, throttles dials, and returns a per-call report.
        `on_transcript` simulates the callee: given the agent opener it
        returns the callee's spoken reply (or None for no answer).
        """
        dnc_set = set(dnc or [])
        report: List[Dict[str, Any]] = []
        min_interval = 60.0 / max(throttle_per_min, 1)
        last_dial = 0.0

        for row in csv.DictReader(numbers_csv.splitlines()):
            phone = (row.get("phone") or "").strip()
            name = (row.get("name") or "").strip()
            if not phone:
                continue
            if phone in dnc_set:
                report.append({"phone": phone, "status": "skipped_dnc"})
                self.logger.log("campaign_skip", phone=phone,
                                reason="do_not_call")
                continue
            wait = min_interval - (time.time() - last_dial)
            if wait > 0:
                time.sleep(wait)
            last_dial = time.time()

            sid = f"CA{hashlib.md5((phone + str(last_dial)).encode()).hexdigest()[:12]}"
            session = CallSession(call_sid=sid, phone=phone,
                                  direction="outbound", state="active",
                                  slots={"name": name})
            self.calls[sid] = session
            opener_text = opener.format(name=name or "there")
            session.say("agent", opener_text)
            callee = on_transcript(opener_text) if on_transcript else None
            if callee is None:
                session.hangup(reason="no_answer")
                status = "no_answer"
            else:
                reply = self.brain.handle(session, callee)
                session.hangup(reason="outbound_complete")
                status = "completed"
            report.append({"phone": phone, "name": name, "status": status,
                           "sid": sid, "turns": len(session.turns)})
            self.logger.log("campaign_call", phone=phone, status=status,
                            sid=sid)
        return report

    # -- reporting ----------------------------------------------------------
    def summary(self) -> Dict[str, Any]:
        calls = list(self.calls.values())
        return {
            "total_calls": len(calls),
            "inbound": sum(1 for c in calls if c.direction == "inbound"),
            "outbound": sum(1 for c in calls if c.direction == "outbound"),
            "escalated": sum(1 for c in calls if c.escalated),
            "ended": sum(1 for c in calls if c.state == "ended"),
            "total_turns": sum(len(c.turns) for c in calls),
            "log_events": len(self.logger.events),
        }


def run_demo() -> None:
    """End-to-end simulated day in the life of the voice agent."""
    logger = CallLogger()
    brain = ConversationBrain(
        business="Swift Towing Co.",
        llm=MockLLMClient({"fallback":
                           "Got it. Anything else I can help with?"}))
    agent = VoiceAgent(brain, logger=logger)
    twilio = TwilioAdapter(auth_token="demo-secret", gather_action_url="/voice")

    print("=== INBOUND CALL (Twilio webhook) ===")
    twiml = agent.inbound_call("CA_INBOUND1", "+15551234567", twilio=twilio)
    print(twiml[:120], "...\n")
    for caller in ["Hi, what are your hours?",
                   "I need a tow booked please",
                   "John Carter",
                   "flatbed tow",
                   "tomorrow morning",
                   "thanks, bye"]:
        print(f"Caller : {caller}")
        agent.caller_said("CA_INBOUND1", caller, twilio=twilio)
        last_agent = next(t for t in reversed(agent.calls['CA_INBOUND1'].turns)
                          if t.role == "agent")
        print(f"Agent  : {last_agent.text}\n")

    print("=== ESCALATION CALL ===")
    agent.inbound_call("CA_ESC1", "+15557654321")
    agent.caller_said("CA_ESC1", "I want to talk to a human agent now")

    print("=== VAPI SERVER MESSAGE ===")
    vapi = VapiAdapter(webhook_secret="vapi-secret")
    ok = vapi.verify({"x-vapi-secret": "vapi-secret"})
    msg = vapi.parse_message(
        {"message": {"type": "assistant-request"},
         "call": {"id": "vapi-call-1"}})
    print("verified:", ok, "| action:", msg, "\n")

    print("=== OUTBOUND CAMPAIGN (3 numbers, 1 on DNC) ===")
    csv_data = "phone,name\n+15550000001,Alice\n+15550000002,Bob\n+15550000003,Carol\n"
    report = agent.outbound_campaign(
        csv_data,
        opener="Hi {name}, this is the AI assistant for Swift Towing Co. — "
               "quick question about your vehicle service?",
        throttle_per_min=600,  # fast for the demo
        dnc=["+15550000002"],
        on_transcript=lambda opener: "Yes, what are your prices?")
    for row in report:
        print(row)

    print("\n=== DAY SUMMARY ===")
    print(json.dumps(agent.summary(), indent=2))
    print("\nEscalated call flagged for human follow-up:",
          agent.calls["CA_ESC1"].escalated)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="VoiceForge voice agent demo")
    parser.add_argument("--demo", action="store_true",
                        help="run the end-to-end simulation")
    args = parser.parse_args()
    if args.demo:
        run_demo()
    else:
        print("Use --demo to run the end-to-end simulation, or run "
              "test_voice_agent.py for the test suite.")
        sys.exit(0)
