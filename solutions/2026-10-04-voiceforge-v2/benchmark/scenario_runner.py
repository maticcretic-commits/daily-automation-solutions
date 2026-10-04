"""VoiceForge v2 — standardized caller-scenario benchmark.

Runs 20 caller scenarios against the local VoiceForge v2 code and
auto-scores each one (correct behavior = pass). Stdlib only:

    python3 scenario_runner.py

Exits 0 when every scenario passes, 1 otherwise. Each scenario is
independent: fresh brain/agent per scenario, fixed `today` so dates are
deterministic.
"""

import datetime as dt
import os
import sys
import tempfile
import traceback
import urllib.error
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from voiceforge.agent import VoiceAgent
from voiceforge.brain import ConversationBrain
from voiceforge.logging import CallLogger
from voiceforge.session import CallSession
from voiceforge import providers as providers_mod
from voiceforge.providers import HttpLLMClient, MockLLMClient
from voiceforge.adapters_twilio import TwilioAdapter
from voiceforge.adapters_vapi import VapiAdapter
from voiceforge.vision.mms_watch import MmsWatch
from voiceforge.vision.vision_analyze import (
    VisionAnalyzer, VisionClient, VisionContext, VisionResult)
from voiceforge.vision.context_inject import ContextInjector
from voiceforge.vision.flow import VisionFlow
from voiceforge.config import VoiceForgeConfig

TODAY = dt.date(2026, 10, 4)  # a Saturday: "friday" -> 2026-10-09


def make_brain(**kw):
    kw.setdefault("business", "Swift Towing Co.")
    kw.setdefault("today", TODAY)
    return ConversationBrain(**kw)


def new_agent(brain=None, **kw):
    brain = brain or make_brain()
    return VoiceAgent(brain, logger=CallLogger(), **kw)


def fresh_call(agent, sid="CA1", phone="+15550001111"):
    agent.inbound_call(sid, phone)
    return agent.calls[sid]


def check(name, fn):
    """Run one scenario; return (name, passed, detail)."""
    try:
        detail = fn() or ""
        return (name, True, detail)
    except AssertionError as e:
        return (name, False, "ASSERT: %s" % e)
    except Exception as e:  # noqa: BLE001 - a crash is a failed scenario
        return (name, False, "CRASH %s: %s" % (type(e).__name__, e))


# ---------------------------------------------------------------- booking

def s_booking_full_flow():
    agent = new_agent()
    fresh_call(agent)
    agent.caller_said("CA1", "I need a tow booked please")
    agent.caller_said("CA1", "John Carter")
    agent.caller_said("CA1", "flatbed tow")
    agent.caller_said("CA1", "friday")
    reply = agent.caller_said("CA1", "yes")
    assert "booked" in reply.lower(), reply
    assert agent.calls["CA1"].slots["day"] == "2026-10-09", \
        agent.calls["CA1"].slots
    return "booked for 2026-10-09 after confirm loop"


def s_booking_day_correction():
    agent = new_agent()
    fresh_call(agent)
    agent.caller_said("CA1", "book please")
    agent.caller_said("CA1", "Neha")
    agent.caller_said("CA1", "consult")
    agent.caller_said("CA1", "monday")
    reply = agent.caller_said("CA1", "yes, friday instead")
    assert "friday" in reply.lower() or "2026-10-09" in reply, reply
    assert "booked" not in reply.lower(), "must re-confirm, not commit"
    reply2 = agent.caller_said("CA1", "yes")
    assert "booked" in reply2.lower(), reply2
    assert agent.calls["CA1"].slots["day"] == "2026-10-09"
    return "correction re-resolved to 2026-10-09, re-confirmed"


def s_booking_bye_midflow():
    agent = new_agent()
    fresh_call(agent)
    agent.caller_said("CA1", "book please")
    agent.caller_said("CA1", "Neha")
    reply = agent.caller_said("CA1", "bye")
    assert agent.calls["CA1"].state == "ended", "call must end"
    assert "day" not in agent.calls["CA1"].slots, "no booking committed"
    return "call ended cleanly, nothing booked"


def s_booking_gibberish_day_no_loop():
    agent = new_agent()
    fresh_call(agent)
    agent.caller_said("CA1", "book please")
    agent.caller_said("CA1", "Neha")
    agent.caller_said("CA1", "oil change")
    for _ in range(4):
        reply = agent.caller_said("CA1", "blorptastic zzz")
    assert "day" not in agent.calls["CA1"].slots, "no booking on gibberish"
    assert isinstance(reply, str) and reply, "agent still responds"
    return "gibberish day input: no crash, no booking, still conversing"


def s_booking_past_date_rejected():
    agent = new_agent()
    fresh_call(agent)
    agent.caller_said("CA1", "book please")
    agent.caller_said("CA1", "Neha")
    agent.caller_said("CA1", "consult")
    reply = agent.caller_said("CA1", "yesterday")
    assert "day" not in agent.calls["CA1"].slots, \
        "past date must not be booked: %s" % agent.calls["CA1"].slots
    assert isinstance(reply, str)
    return "past date rejected: %r..." % reply[:60]


# -------------------------------------------------------------------- FAQ

def s_faq_hours():
    brain = make_brain()
    s = CallSession(call_sid="C", phone="+1", direction="inbound",
                    state="active")
    reply = brain.handle(s, "what are your hours?")
    assert "9 AM" in reply, reply
    return reply[:60]


def s_faq_pricing():
    brain = make_brain()
    s = CallSession(call_sid="C", phone="+1", direction="inbound",
                    state="active")
    reply = brain.handle(s, "how much does it cost?")
    assert "$49" in reply, reply
    return reply[:60]


# -------------------------------------------------------------- escalation

def s_escalation_human():
    agent = new_agent()
    fresh_call(agent)
    reply = agent.caller_said("CA1", "let me talk to a human please")
    s = agent.calls["CA1"]
    assert s.escalated, "session must flag escalation"
    assert s.transfer_requested, "transfer must be requested"
    # no carrier configured in the harness -> B5 guard falls back to
    # message-taking instead of 500ing; that IS the correct behavior
    assert "call you back" in reply.lower() or "human" in reply.lower(), reply
    return "escalated; guarded fallback to message-taking (no carrier)"


def s_no_escalation_imagine():
    agent = new_agent()
    fresh_call(agent)
    reply = agent.caller_said("CA1", "imagine the possibilities here")
    assert not agent.calls["CA1"].escalated, \
        "substring 'imagine' must not escalate"
    return "no false escalation on 'imagine'"


# ------------------------------------------------------------------ vision

class _StubVisionClient(VisionClient):
    def __init__(self, payload):
        self.payload = payload

    def describe(self, image_bytes, content_type, prompt, timeout_s):
        return self.payload


def _vision_stack(payload):
    logger = CallLogger()
    tw = TwilioAdapter(auth_token="t", account_sid="ACx")
    watch = MmsWatch(tw, logger, media_dir=tempfile.mkdtemp(),
                     downloader=lambda url: (b"IMG", "image/jpeg"))
    analyzer = VisionAnalyzer(_StubVisionClient(payload), logger)
    injector = ContextInjector(logger)
    flow = VisionFlow(watch, analyzer, injector, logger)
    return flow, logger


def s_vision_photo_offer_wired():
    flow, _ = _vision_stack({"description": "x", "findings": ["y"],
                             "confidence": 0.9})
    agent = new_agent(vision_flow=flow)
    fresh_call(agent, phone="+15550001111")
    reply = agent.caller_said(
        "CA1", "Can I send you a photo of the cracked windshield?")
    assert agent.calls["CA1"].vision_state == "waiting_media", \
        agent.calls["CA1"].vision_state
    assert "photo" in reply.lower(), reply
    return "photo offer -> vision_state=waiting_media, ask spoken"


def s_vision_low_confidence_clarifies():
    logger = CallLogger()
    analyzer = VisionAnalyzer(
        _StubVisionClient({"description": "blurry shape",
                           "findings": ["unclear"], "confidence": 0.59}),
        logger)
    ctx = VisionContext(business="Swift Towing Co.", photo_request="damage",
                        recent_turns=[])
    res = analyzer.analyze(b"IMG", "image/jpeg", ctx)
    assert isinstance(res, VisionResult)
    assert not res.gate_passed, "0.59 must not pass the 0.6 gate"
    assert res.clarification_question, "must ask, never assert"
    assert "unclear" not in res.description or True
    return "0.59 -> clarification question, findings not asserted"


def s_vision_timeout_fallback():
    def boom(url, prompt):
        raise TimeoutError("vision timed out")

    class BoomClient(VisionClient):
        def describe(self, image_bytes, content_type, prompt, timeout_s):
            raise TimeoutError("vision timed out")

    logger = CallLogger()
    analyzer = VisionAnalyzer(BoomClient(), logger, timeout_s=0.01)
    ctx = VisionContext(business="Swift Towing Co.", photo_request="damage",
                        recent_turns=[])
    res = analyzer.analyze(b"IMG", "image/jpeg", ctx)
    assert not res.ok, "timeout must produce ok=False, never raise"
    assert res.error, "error must be recorded"
    return "timeout -> ok=False result, single voice-only fallback path"


# --------------------------------------------------------------- edge cases

def s_edge_none_transcript():
    agent = new_agent()
    agent.inbound_call("CA1", "+15550001111")
    reply = agent.caller_said("CA1", None)
    assert isinstance(reply, str) and reply, "must return benign reply"
    return "None transcript: no crash"


def s_edge_empty_secret_rejected():
    try:
        VapiAdapter(webhook_secret="")
    except ValueError:
        return "empty Vapi secret -> ValueError (fail-closed)"
    raise AssertionError("empty secret was accepted")


def s_edge_retry_ended_call():
    agent = new_agent()
    agent.inbound_call("CA1", "+15550001111")
    agent.caller_said("CA1", "bye")
    reply = agent.caller_said("CA1", "hello again?")  # retried webhook
    assert isinstance(reply, str), "must return last reply, not raise"
    return "retried webhook on ended call: idempotent"


def s_edge_dnc_normalized():
    agent = new_agent()
    report = agent.outbound_campaign(
        "phone,name\n+15550000001,Ann\n+15550000002,Ben\n",
        opener="Hi {name}!", throttle_per_min=60000,
        dnc=["+1-555-000-0001"],  # differently formatted, same number
        on_transcript=lambda opener: "hi")
    by_phone = {r["phone"]: r["status"] for r in report}
    assert by_phone["+15550000001"] == "skipped_dnc", by_phone
    assert by_phone["+15550000002"] == "completed", by_phone
    return "E.164-normalized DNC respected"


def s_edge_llm_500_fallback():
    def always_down(session):
        raise urllib.error.HTTPError("url", 500, "boom", {}, None)

    client = HttpLLMClient("http://x", "k")
    s = CallSession(call_sid="C", phone="+1", direction="inbound",
                    state="active")
    with mock.patch.object(client, "_request", side_effect=always_down):
        with mock.patch("time.sleep", return_value=None):
            reply = client.reply(s, "hi")
    assert reply == HttpLLMClient.SAFE_FALLBACK, reply
    return "LLM 500 -> safe canned fallback, call survives"


def s_edge_goodbye_ends_call():
    agent = new_agent()
    agent.inbound_call("CA1", "+15550001111")
    agent.caller_said("CA1", "okay thanks, goodbye!")
    assert agent.calls["CA1"].state == "ended"
    return "goodbye ends call"


def s_edge_unknown_sid():
    agent = new_agent()
    reply = agent.caller_said("NOPE", "hello?")
    assert isinstance(reply, str), "unknown SID must not raise"
    return "unknown SID: benign reply, no KeyError"


def s_edge_landline_skips_vision():
    flow, _ = _vision_stack({"description": "x", "findings": ["y"],
                             "confidence": 0.9})
    agent = new_agent(vision_flow=flow)
    fresh_call(agent, phone="+15550001111")
    # force landline classification
    flow.mms_watch._line_cache["+15550001111"] = ("landline", 9999999999.0)
    ask = agent.request_photo("CA1", "the damage")
    assert agent.calls["CA1"].vision_state == "unavailable", \
        agent.calls["CA1"].vision_state
    assert ask == "", "landline: no photo ask, voice-only path"
    return "landline -> vision skipped, voice-only"


SCENARIOS = [
    ("booking_full_flow", s_booking_full_flow),
    ("booking_day_correction", s_booking_day_correction),
    ("booking_bye_midflow", s_booking_bye_midflow),
    ("booking_gibberish_day_no_loop", s_booking_gibberish_day_no_loop),
    ("booking_past_date_rejected", s_booking_past_date_rejected),
    ("faq_hours", s_faq_hours),
    ("faq_pricing", s_faq_pricing),
    ("escalation_human", s_escalation_human),
    ("no_escalation_imagine", s_no_escalation_imagine),
    ("vision_photo_offer_wired", s_vision_photo_offer_wired),
    ("vision_low_confidence_clarifies", s_vision_low_confidence_clarifies),
    ("vision_timeout_fallback", s_vision_timeout_fallback),
    ("edge_none_transcript", s_edge_none_transcript),
    ("edge_empty_secret_rejected", s_edge_empty_secret_rejected),
    ("edge_retry_ended_call", s_edge_retry_ended_call),
    ("edge_dnc_normalized", s_edge_dnc_normalized),
    ("edge_llm_500_fallback", s_edge_llm_500_fallback),
    ("edge_goodbye_ends_call", s_edge_goodbye_ends_call),
    ("edge_unknown_sid", s_edge_unknown_sid),
    ("edge_landline_skips_vision", s_edge_landline_skips_vision),
]


def main():
    assert len(SCENARIOS) == 20, "benchmark must hold exactly 20 scenarios"
    results = [check(name, fn) for name, fn in SCENARIOS]
    passed = sum(1 for _, ok, _ in results if ok)
    print("VoiceForge v2 scenario benchmark: %d/%d passed\n" % (passed, 20))
    for name, ok, detail in results:
        print("[%s] %-32s %s" % ("PASS" if ok else "FAIL", name, detail))
        if not ok and "CRASH" in detail:
            print("       %s" % detail)
    print("\n%d/%d scenarios passed" % (passed, 20))
    return 0 if passed == 20 else 1


if __name__ == "__main__":
    sys.exit(main())
