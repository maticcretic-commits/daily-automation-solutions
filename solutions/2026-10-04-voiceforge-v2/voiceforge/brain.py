"""Conversation brain: intents, hardened booking flow, escalation.

P0 fixes vs v1:
  1. Escalation matches on WORD BOUNDARIES ("imagine" no longer transfers).
  2. Goodbye is checked BEFORE booking, so "bye" mid-booking ends the call.
  3. Unknown booking steps reset to the name step (no infinite loop).
  4. Booking normalizes slot text to real dates, runs a confirm-before-commit
     loop, enforces per-day caps, and refuses double bookings.
  5. The FAQ catch-all no longer swallows every "do you ..." question.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Dict, List, Optional, Pattern, Tuple

from .config import VoiceForgeConfig
from .providers import LLMClient, MockLLMClient
from .session import CallSession

GREETING = ("Hi, I'm the AI assistant for {business}. "
            "This call may be recorded to help us serve you better. "
            "How can I help you today?")
GREETING_NO_DISCLOSURE = ("Hi, this is the AI assistant for {business}. "
                           "How can I help you today?")

ESCALATION_PHRASES = ("agent", "human", "representative", "manager",
                      "real person", "support")
BOOKING_PHRASES = ("book", "appointment", "schedule", "slot", "reschedule")
END_PHRASES = ("bye", "goodbye", "that's all", "no thanks")
AFFIRM = ("yes", "yeah", "yep", "correct", "right", "sure", "ok",
          "okay", "confirmed")
DENY = ("no", "nope", "nah", "wrong", "incorrect", "not correct")

_ESCALATION_RE = re.compile(
    r"\b(" + "|".join(re.escape(p) for p in ESCALATION_PHRASES) + r")\b", re.I)
_END_RE = re.compile(
    r"\b(" + "|".join(re.escape(p) for p in END_PHRASES) + r")\b", re.I)
_BOOKING_RE = re.compile(
    r"\b(" + "|".join(re.escape(p) for p in BOOKING_PHRASES) + r")\b", re.I)
_AFFIRM_RE = re.compile(r"\b(" + "|".join(AFFIRM) + r")\b", re.I)
_DENY_RE = re.compile(r"\b(" + "|".join(DENY) + r")\b", re.I)

FAQ_PATTERNS: List[Tuple[Pattern[str], str]] = [
    (re.compile(r"\b(hours?|open(ing)?|close|timing)\b", re.I),
     "We're open Monday to Saturday, 9 AM to 7 PM."),
    (re.compile(r"\b(price|cost|charge|fee|rate)\b", re.I),
     "Prices start from $49 and depend on the service. "
     "Would you like me to check a specific service?"),
    (re.compile(r"\b(where|location|address|parking)\b", re.I),
     "We're located downtown, with parking right outside."),
    (re.compile(r"\b(services?|offerings?)\b", re.I),
     "We offer consultations, bookings, and priority support. "
     "What can I help you with?"),
]

WEEKDAYS = {"monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
            "friday": 4, "saturday": 5, "sunday": 6}


def parse_day(text: str, today: dt.date) -> Optional[dt.date]:
    """
    Normalize day-ish slot text to a real date. Returns None when the text
    cannot be resolved to a valid, non-past date.
    """
    t = text.strip().lower()
    if not t:
        return None
    if t == "today":
        return today
    if t == "tomorrow":
        return today + dt.timedelta(days=1)
    if t in ("day after tomorrow",):
        return today + dt.timedelta(days=2)
    # ISO date
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        try:
            d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
        return d if d >= today else None
    # "15th", "15 oct", "october 15"
    m = re.fullmatch(
        r"(?:(\d{1,2})(?:st|nd|rd|th)?\s*)?(jan|feb|mar|apr|may|jun|jul|aug|"
        r"sep|oct|nov|dec)[a-z]*\s*(\d{1,2})(?:st|nd|rd|th)?", t)
    if not m and re.fullmatch(r"(\d{1,2})(?:st|nd|rd|th)", t):
        day = int(re.fullmatch(r"(\d{1,2})(?:st|nd|rd|th)", t).group(1))
        try:
            cand = dt.date(today.year, today.month, day)
        except ValueError:
            return None
        if cand < today:
            try:
                cand = dt.date(today.year + (1 if today.month == 12 else 0),
                               today.month % 12 + 1, day)
            except ValueError:
                return None
        return cand
    if m:
        month_name, day_s = m.group(2), m.group(3)
        month = ("jan feb mar apr may jun jul aug sep oct nov dec"
                 ).split().index(month_name[:3]) + 1
        day = int(day_s)
        year = today.year
        try:
            cand = dt.date(year, month, day)
        except ValueError:
            return None
        if cand < today:
            cand = dt.date(year + 1, month, day)
        return cand
    # weekday names, with optional "next"
    m = re.fullmatch(r"(next\s+)?(monday|tuesday|wednesday|thursday|friday|"
                     r"saturday|sunday)", t)
    if m:
        target = WEEKDAYS[m.group(2)]
        delta = (target - today.weekday()) % 7
        if m.group(1):
            # "next <day>": the occurrence a week after the plain reading
            delta = delta + 7 if delta else 7
        return today + dt.timedelta(days=delta)
    return None


class BookingStore:
    """In-memory booking ledger with per-day caps and double-book guard."""

    def __init__(self, daily_cap: int = 20):
        self.daily_cap = daily_cap
        self._bookings: Dict[str, List[Dict[str, str]]] = {}

    def add(self, day_iso: str, name: str,
            service: str) -> Tuple[bool, str]:
        day = self._bookings.setdefault(day_iso, [])
        if len(day) >= self.daily_cap:
            return False, (f"I'm sorry, {day_iso} is fully booked. "
                           f"Could I offer you another day?")
        key = (name.strip().lower(), service.strip().lower())
        if any((b["name"].strip().lower(), b["service"].strip().lower()) == key
               for b in day):
            return False, (f"You already have {service} booked on {day_iso}, "
                           f"{name}. Would you like a different day?")
        day.append({"name": name, "service": service})
        return True, ""

    def count(self, day_iso: str) -> int:
        return len(self._bookings.get(day_iso, []))


class ConversationBrain:
    """
    Decides the agent's reply and next action for one caller utterance.

    Intent order: escalation -> goodbye -> booking -> FAQ -> LLM fallback.
    Booking is a slot-fill (name -> service -> day -> confirm) driven by
    `slots`, with real date normalization and a commit loop.
    """

    def __init__(self, business: str = "Acme Services",
                 llm: Optional[LLMClient] = None,
                 config: Optional[VoiceForgeConfig] = None,
                 booking_store: Optional[BookingStore] = None,
                 today: Optional[dt.date] = None):
        self.business = business
        self.llm = llm or MockLLMClient()
        self.config = config or VoiceForgeConfig(business=business)
        self.booking_store = booking_store or BookingStore(
            daily_cap=self.config.booking_daily_cap)
        self.today = today or dt.date.today()

    # -- public API ------------------------------------------------------
    def greet(self, session: CallSession) -> str:
        template = (GREETING if self.config.ai_disclosure
                    else GREETING_NO_DISCLOSURE)
        text = template.format(business=self.business)
        session.say("agent", text)
        return text

    def handle(self, session: CallSession, caller_text: str) -> str:
        """Process one caller utterance; returns the agent's reply text."""
        session.say("caller", caller_text)
        lowered = caller_text.lower()

        if _ESCALATION_RE.search(lowered):
            session.escalated = True
            session.transfer_requested = True
            text = ("Of course — connecting you to a human teammate now. "
                    "I'll brief them on what we've discussed. "
                    "Please hold for a moment.")
        elif _END_RE.search(lowered):
            # goodbye BEFORE booking: "bye" mid-booking ends the call cleanly
            text = "Thank you for calling. Have a great day!"
            session.say("agent", text)
            session.hangup(reason="caller goodbye")
            return text
        elif self._in_booking(session) or _BOOKING_RE.search(lowered):
            text = self._booking_step(session, caller_text)
        elif self._faq(caller_text) is not None:
            text = self._faq(caller_text)
        else:
            try:
                text = self.llm.reply(session, caller_text)
            except Exception:  # noqa: BLE001 - last-resort boundary
                text = ("I'm having a little trouble on my end right now — "
                        "could you say that once more?")

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
            name = caller_text.strip()
            if not name:
                return "I didn't catch your name — could you repeat it?"
            slots["name"] = name
            slots["_booking_step"] = "service"
            return (f"Thanks {name}. Which service would you like to book?")
        if step == "service":
            service = caller_text.strip()
            if not service:
                return "Which service would you like to book?"
            slots["service"] = service
            slots["_booking_step"] = "day"
            return f"{service} — got it. Which day works for you?"
        if step == "day":
            day = parse_day(caller_text, self.today)
            if day is None:
                return ("I couldn't pin that to a calendar date — "
                        "could you give me a day like 'tomorrow' or "
                        "'next Monday'?")
            slots["day_iso"] = day.isoformat()
            slots["_booking_step"] = "confirm"
            return (f"Just to confirm: {slots['service']} on "
                    f"{day.strftime('%A, %B %d')} for {slots['name']} — "
                    f"is that correct?")
        if step == "confirm":
            lowered = caller_text.strip().lower()
            denied = (lowered == "n" or (
                _DENY_RE.search(lowered)
                and not _AFFIRM_RE.search(lowered)))
            if denied:
                slots["_booking_step"] = "day"
                return "No problem — which day would you prefer instead?"
            if lowered == "y" or _AFFIRM_RE.search(lowered):
                ok, msg = self.booking_store.add(
                    slots["day_iso"], slots["name"], slots["service"])
                if not ok:
                    slots["_booking_step"] = "day"
                    return msg  # cap hit or double-book: offer another day
                day_iso = slots.pop("day_iso")
                name = slots.pop("name")
                service = slots.pop("service")
                slots["day"] = day_iso  # normalized date kept for records
                slots.pop("_booking_active", None)
                slots.pop("_booking_step", None)
                return (f"You're booked, {name}: {service} on {day_iso}. "
                        f"A confirmation will be sent by SMS.")
            # unclear answer -> ask again, do NOT loop forever silently
            return ("Sorry, I didn't catch that — is the booking correct? "
                    "Please say yes or no.")
        # unknown/corrupt step: reset instead of trapping the caller (P0-1(4))
        slots["_booking_step"] = "name"
        return "Sorry, let's start over — what's your full name?"

    # -- FAQ --------------------------------------------------------------
    @staticmethod
    def _faq(caller_text: str) -> Optional[str]:
        for pattern, answer in FAQ_PATTERNS:
            if pattern.search(caller_text):
                return answer
        return None
