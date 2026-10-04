"""Tests for VoiceForge v2. Stdlib only: run with `python3 test_voiceforge_v2.py`.

Covers: ported v1 behaviors (adapted to the hardened booking flow), a
regression test for EACH of the 7 pilot-killers, booking correctness, warm
transfer, and the three vision modules with mocked providers.
"""

import base64
import datetime as dt
import hashlib
import hmac
import json
import os
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from unittest import mock

from voiceforge import (
    CallLogger, CallSession, CampaignRunner, ConversationBrain,
    HttpLLMClient, MockLLMClient, MockTTSClient, MockVisionClient,
    TwilioAdapter, VapiAdapter, VoiceAgent, VoiceForgeConfig,
    normalize_phone, BookingStore, parse_day,
)
from voiceforge import providers as providers_mod
from voiceforge.vision import (
    ContextInjector, MmsWatch, VisionAnalyzer, VisionContext, VisionResult,
    VOICE_ONLY_FALLBACK,
)
from voiceforge.vision.flow import VisionFlow

TODAY = dt.date(2026, 10, 4)  # a Sunday


def make_brain(**kw):
    kw.setdefault("today", TODAY)
    return ConversationBrain(business="Test Co.", **kw)


def new_session(sid="CA1"):
    return CallSession(call_sid=sid, phone="+1000", direction="inbound",
                       state="active")


# ---------------------------------------------------------------------------
# Ported v1 behaviors
# ---------------------------------------------------------------------------

class TestBrainBasics(unittest.TestCase):
    def setUp(self):
        self.brain = make_brain()

    def test_greeting_discloses_ai(self):
        s = new_session()
        text = self.brain.greet(s)
        self.assertIn("Test Co.", text)
        self.assertIn("AI assistant", text)
        self.assertIn("recorded", text)
        self.assertEqual(s.turns[-1].role, "agent")

    def test_greeting_without_disclosure(self):
        brain = make_brain(config=VoiceForgeConfig(ai_disclosure=False))
        s = new_session()
        text = brain.greet(s)
        self.assertIn("Test Co.", text)
        self.assertNotIn("recorded", text)

    def test_faq_hours(self):
        s = new_session()
        self.assertIn("9 AM", self.brain.handle(s, "What are your hours?"))

    def test_faq_pricing(self):
        s = new_session()
        self.assertIn("$49", self.brain.handle(s, "How much does it cost?"))

    def test_faq_scoped_no_catchall(self):
        # v1 answered "do you have parking?" with the generic services blurb
        s = new_session()
        reply = self.brain.handle(s, "Do you have parking?")
        self.assertIn("parking", reply.lower())
        self.assertNotIn("consultations", reply)

    def test_escalation_flags_session(self):
        s = new_session()
        reply = self.brain.handle(s, "Let me talk to a human please")
        self.assertTrue(s.escalated)
        self.assertTrue(s.transfer_requested)
        self.assertIn("human", reply.lower())

    def test_imagine_does_not_escalate(self):
        s = new_session()
        reply = self.brain.handle(s, "Imagine my car broke down downtown")
        self.assertFalse(s.escalated)
        self.assertFalse(s.transfer_requested)

    def test_goodbye_hangs_up(self):
        s = new_session()
        reply = self.brain.handle(s, "Okay thanks, goodbye!")
        self.assertEqual(s.state, "ended")
        self.assertIsNotNone(s.ended_at)
        self.assertIn("great day", reply.lower())

    def test_fallback_uses_llm(self):
        brain = make_brain(
            llm=MockLLMClient({"fallback": "CUSTOM FALLBACK REPLY"}))
        s = new_session()
        self.assertEqual(brain.handle(s, "tell me about quantum tunneling"),
                         "CUSTOM FALLBACK REPLY")

    def test_llm_exception_boundary(self):
        class Boom(MockLLMClient):
            def reply(self, session, text):
                raise RuntimeError("dead")
        brain = make_brain(llm=Boom())
        s = new_session()
        reply = brain.handle(s, "something unanswerable")
        self.assertIn("trouble", reply)
        self.assertEqual(s.state, "active")  # call survives


class TestBookingFlow(unittest.TestCase):
    def setUp(self):
        self.brain = make_brain()

    def run_flow(self, s, *utterances):
        return [self.brain.handle(s, u) for u in utterances]

    def test_full_flow_with_confirm(self):
        s = new_session()
        r = self.run_flow(s, "I want to book an appointment", "Ravi Sharma",
                          "oil change", "tomorrow", "yes")
        self.assertIn("name", r[0].lower())
        self.assertIn("service", r[1].lower())
        self.assertIn("day", r[2].lower())
        self.assertIn("confirm", r[3].lower())
        self.assertIn("October 05", r[3])  # tomorrow normalized to a date
        self.assertIn("booked", r[4].lower())
        self.assertIn("2026-10-05", r[4])
        self.assertEqual(s.slots["day"], "2026-10-05")
        self.assertNotIn("_booking_active", s.slots)

    def test_deny_restarts_at_day(self):
        s = new_session()
        self.run_flow(s, "book please", "Ravi", "oil change", "tomorrow")
        r = self.brain.handle(s, "no, friday please")
        self.assertIn("day", r.lower())
        self.assertEqual(s.slots["_booking_step"], "day")

    def test_bye_mid_booking_ends_call(self):
        s = new_session()
        self.run_flow(s, "book please", "Ravi")
        reply = self.brain.handle(s, "bye")
        self.assertEqual(s.state, "ended")
        self.assertIn("great day", reply.lower())
        self.assertNotEqual(s.slots.get("service"), "bye")

    def test_unknown_step_resets(self):
        s = new_session()
        s.slots["_booking_active"] = "1"
        s.slots["_booking_step"] = "corrupted"
        reply = self.brain.handle(s, "hello?")
        self.assertIn("start over", reply.lower())
        self.assertEqual(s.slots["_booking_step"], "name")
        # and the next utterance proceeds normally (no infinite loop)
        reply2 = self.brain.handle(s, "Ravi")
        self.assertIn("service", reply2.lower())

    def test_past_date_rejected(self):
        s = new_session()
        self.run_flow(s, "book", "Ravi", "cut", "2026-10-01")
        # 2026-10-01 is before TODAY -> ask again
        self.assertIn("calendar date", s.turns[-1].text)

    def test_double_booking_refused(self):
        s1, s2 = new_session("A"), new_session("B")
        self.run_flow(s1, "book", "Ravi", "cut", "tomorrow", "yes")
        r = self.run_flow(s2, "book", "Ravi", "cut", "tomorrow", "yes")
        self.assertIn("already have", r[-1].lower())

    def test_daily_cap(self):
        brain = make_brain(booking_store=BookingStore(daily_cap=1))
        s1, s2 = new_session("A"), new_session("B")
        for u in ("book", "Ann", "cut", "tomorrow", "yes"):
            brain.handle(s1, u)
        replies = [brain.handle(s2, u)
                   for u in ("book", "Bob", "cut", "tomorrow", "yes")]
        self.assertIn("fully booked", replies[-1].lower())

    def test_unclear_confirm_asks_again(self):
        s = new_session()
        self.run_flow(s, "book", "Ravi", "cut", "tomorrow")
        r = self.brain.handle(s, "maybe later")
        self.assertIn("yes or no", r.lower())
        self.assertEqual(s.slots["_booking_step"], "confirm")


class TestParseDay(unittest.TestCase):
    def test_today_tomorrow(self):
        self.assertEqual(parse_day("today", TODAY), TODAY)
        self.assertEqual(parse_day("tomorrow", TODAY),
                         dt.date(2026, 10, 5))

    def test_weekday(self):
        self.assertEqual(parse_day("monday", TODAY), dt.date(2026, 10, 5))
        self.assertEqual(parse_day("friday", TODAY), dt.date(2026, 10, 9))

    def test_next_weekday(self):
        self.assertEqual(parse_day("next monday", TODAY),
                         dt.date(2026, 10, 12))

    def test_iso(self):
        self.assertEqual(parse_day("2026-10-20", TODAY),
                         dt.date(2026, 10, 20))

    def test_past_iso_rejected(self):
        self.assertIsNone(parse_day("2026-10-01", TODAY))

    def test_garbage_rejected(self):
        self.assertIsNone(parse_day("someday never", TODAY))
        self.assertIsNone(parse_day("", TODAY))


# ---------------------------------------------------------------------------
# The 7 pilot-killer regressions
# ---------------------------------------------------------------------------

class TestPilotKillers(unittest.TestCase):
    def test_k1_empty_vapi_secret_rejected(self):
        with mock.patch.dict(os.environ, {"VAPI_WEBHOOK_SECRET": ""}):
            with self.assertRaises(ValueError):
                VapiAdapter(webhook_secret="")
        # and a missing header never verifies
        v = VapiAdapter(webhook_secret="s3cret")
        self.assertFalse(v.verify({}))

    def test_k1b_empty_twilio_token_rejected(self):
        with mock.patch.dict(os.environ, {"TWILIO_AUTH_TOKEN": ""}):
            with self.assertRaises(ValueError):
                TwilioAdapter(auth_token="")

    def test_k2_llm_429_then_fallback(self):
        calls = {"n": 0}

        def flaky():
            calls["n"] += 1
            if calls["n"] == 1:
                raise urllib.error.HTTPError(
                    "url", 429, "rate limited", {}, None)
            return "recovered"

        ok, _ = providers_mod.call_with_retry(flaky, retries=1,
                                              backoff_s=0)
        self.assertTrue(ok)

        def always_down(session):
            raise urllib.error.HTTPError("url", 500, "boom", {}, None)

        client = HttpLLMClient("http://x", "k")
        with mock.patch.object(client, "_request", side_effect=always_down):
            with mock.patch("time.sleep", return_value=None):
                reply = client.reply(new_session(), "hi")
        self.assertEqual(reply, HttpLLMClient.SAFE_FALLBACK)

    def test_k2b_llm_malformed_response_falls_back(self):
        class FakeResp:
            def read(self):
                return json.dumps({"choices": []}).encode()

            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

        client = HttpLLMClient("http://x", "k")
        with mock.patch.object(urllib.request, "urlopen",
                               return_value=FakeResp()):
            reply = client.reply(new_session(), "hi")
        self.assertEqual(reply, HttpLLMClient.SAFE_FALLBACK)

    def test_k3_unknown_sid_no_raise(self):
        agent = VoiceAgent(make_brain())
        out = agent.caller_said("NOPE", "hello?")
        self.assertIsInstance(out, str)

    def test_k3b_ended_session_returns_last_reply(self):
        agent = VoiceAgent(make_brain())
        agent.inbound_call("CA1", "+1000")
        agent.caller_said("CA1", "bye")
        out = agent.caller_said("CA1", "are you still there?")
        self.assertIn("great day", out.lower())  # last agent reply, no raise

    def test_k4_bye_mid_booking(self):
        brain = make_brain()
        s = new_session()
        brain.handle(s, "book please")
        brain.handle(s, "Ravi")
        reply = brain.handle(s, "bye")
        self.assertEqual(s.state, "ended")
        self.assertIn("great day", reply.lower())

    def test_k4b_unknown_step_no_loop(self):
        brain = make_brain()
        s = new_session()
        s.slots.update({"_booking_active": "1", "_booking_step": "bogus"})
        r1 = brain.handle(s, "huh")   # resets to name step
        self.assertIn("start over", r1.lower())
        r2 = brain.handle(s, "Ravi")  # name accepted -> asks service
        self.assertIn("service", r2.lower())
        r3 = brain.handle(s, "cut")   # service accepted -> asks day
        self.assertIn("day", r3.lower())

    def test_k5_imagine_no_escalation(self):
        brain = make_brain()
        s = new_session()
        # "reagent" contains "agent" as a substring: v1 escalated on this
        brain.handle(s, "The reagent levels look fine to me")
        self.assertFalse(s.escalated)
        self.assertFalse(s.transfer_requested)
        # ...but a real standalone request still escalates
        s2 = new_session()
        brain.handle(s2, "Please get me the manager now")
        self.assertTrue(s2.escalated)

    def test_k6_log_is_thread_safe_and_redacts(self):
        logger = CallLogger()
        logger.log("turn", phone="+1-555-000-1234",
                   caller="call me at 555 000 1234 please")
        rec = logger.events[-1]
        self.assertNotIn("+1-555-000-1234", json.dumps(rec))
        self.assertIn("phone_hash", rec)
        self.assertIn("[REDACTED_PHONE]", rec["caller"])

        def hammer(n):
            for i in range(50):
                logger.log("t", i=i, worker=n)

        threads = [threading.Thread(target=hammer, args=(n,))
                   for n in range(8)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self.assertEqual(len(logger.events), 1 + 8 * 50)

    def test_k6b_log_file_lines_are_valid_json(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "audit.jsonl")
            logger = CallLogger(path=path)
            logger.log("call_started", phone="+15550001111")
            with open(path, encoding="utf-8") as fh:
                lines = fh.read().strip().split("\n")
            self.assertEqual(len(lines), 1)
            rec = json.loads(lines[0])
            self.assertNotIn("15550001111", lines[0])
            self.assertIn("phone_hash", rec)

    def test_k7_dnc_normalized(self):
        brain = make_brain()
        runner = CampaignRunner(brain)
        csv_data = "phone,name\n+1-555-000-0001,Alice\n+15550000002,Bob\n"
        report = runner.run(csv_data, opener="Hi {name}",
                            throttle_per_min=60000,
                            dnc=["+15550000001", "1 (555) 000-0002"],
                            on_transcript=lambda o: "hi",
                            sleep=lambda s: None)
        by_phone = {r["phone"]: r["status"] for r in report}
        self.assertEqual(by_phone["+1-555-000-0001"], "skipped_dnc")
        self.assertEqual(by_phone["+15550000002"], "skipped_dnc")


# ---------------------------------------------------------------------------
# Twilio / Vapi adapters
# ---------------------------------------------------------------------------

def twilio_sig(token, url, params):
    payload = url + "".join(k + params[k] for k in sorted(params))
    return base64.b64encode(
        hmac.new(token.encode(), payload.encode(),
                 hashlib.sha1).digest()).decode()


class TestTwilioAdapter(unittest.TestCase):
    def setUp(self):
        self.tw = TwilioAdapter(auth_token="secret123",
                                gather_action_url="/voice")

    def test_signature_verifies(self):
        url, params = "https://example.com/voice", {"From": "+1555"}
        self.assertTrue(self.tw.verify_signature(
            url, params, twilio_sig("secret123", url, params)))

    def test_signature_rejects_tampered(self):
        self.assertFalse(self.tw.verify_signature(
            "https://example.com/voice", {"From": "+1555"}, "AAAA"))

    def test_twiml_is_valid_xml(self):
        xml = self.tw.inbound_twiml("Hello & welcome")
        root = ET.fromstring(xml)
        self.assertEqual(root.tag, "Response")
        self.assertIn("Hello &amp; welcome", xml)

    def test_transfer_twiml_has_dial_and_whisper(self):
        xml = self.tw.transfer_twiml("One moment please", "+1999",
                                     "/whisper?sid=CA1", "/fb?sid=CA1")
        root = ET.fromstring(xml)
        dial = root.find("Dial")
        self.assertIsNotNone(dial)
        number = dial.find("Number")
        self.assertEqual(number.get("url"), "/whisper?sid=CA1")
        self.assertEqual(dial.get("action"), "/fb?sid=CA1")

    def test_transfer_fallback_twiml(self):
        xml = self.tw.transfer_fallback_twiml()
        ET.fromstring(xml)  # must parse
        self.assertIn("call you back", xml)


class TestVapiAdapter(unittest.TestCase):
    def test_verify_ok_and_wrong(self):
        v = VapiAdapter(webhook_secret="s3cret")
        self.assertTrue(v.verify({"x-vapi-secret": "s3cret"}))
        self.assertFalse(v.verify({"x-vapi-secret": "wrong"}))

    def test_parse_variants(self):
        v = VapiAdapter(webhook_secret="x")
        self.assertEqual(v.parse_message(
            {"message": {"type": "assistant-request"},
             "call": {"id": "c1"}})["action"], "start_call")
        out = v.parse_message(
            {"message": {"type": "transcript", "role": "user",
                         "transcript": "hello"}, "call": {"id": "c2"}})
        self.assertEqual((out["action"], out["text"]),
                         ("caller_said", "hello"))

    def test_parse_malformed_never_raises(self):
        v = VapiAdapter(webhook_secret="x")
        for bad in ({"message": {"type": "transcript"}},  # no role chain
                    {"message": "nope"},
                    {},
                    None,
                    {"message": {"type": "transcript", "role": "user"}}):
            out = v.parse_message(bad)
            self.assertEqual(out["action"], "ignore")


# ---------------------------------------------------------------------------
# Agent flows incl. warm transfer
# ---------------------------------------------------------------------------

class TestAgentFlows(unittest.TestCase):
    def setUp(self):
        self.brain = make_brain()
        self.agent = VoiceAgent(self.brain, logger=CallLogger())

    def test_inbound_full_flow(self):
        g = self.agent.inbound_call("CA100", "+1999")
        self.assertIn("Test Co.", g)
        self.agent.caller_said("CA100", "what are your hours?")
        self.agent.caller_said("CA100", "book please")
        self.agent.caller_said("CA100", "Neha")
        self.agent.caller_said("CA100", "consult")
        self.agent.caller_said("CA100", "friday")
        reply = self.agent.caller_said("CA100", "yes")
        self.assertIn("booked", reply.lower())
        self.assertEqual(
            self.agent.calls["CA100"].slots["day"], "2026-10-09")
        self.agent.caller_said("CA100", "bye!")
        self.assertEqual(self.agent.calls["CA100"].state, "ended")

    def test_warm_transfer_twiml(self):
        tw = TwilioAdapter(auth_token="s3cret")
        agent = VoiceAgent(
            self.brain, logger=CallLogger(),
            config=VoiceForgeConfig(transfer_number="+18885551212"))
        agent.inbound_call("CA1", "+1000", twilio=tw)
        out = agent.caller_said("CA1", "get me a human agent", twilio=tw)
        root = ET.fromstring(out)
        self.assertIsNotNone(root.find("Dial"))
        self.assertIn("+18885551212", out)

    def test_transfer_fallback_when_no_target(self):
        agent = VoiceAgent(self.brain, logger=CallLogger(),
                           config=VoiceForgeConfig(transfer_number=""))
        agent.inbound_call("CA1", "+1000")
        out = agent.caller_said("CA1", "get me a human agent")
        self.assertIn("call you back", out.lower())

    def test_transfer_check_unreachable(self):
        agent = VoiceAgent(self.brain, logger=CallLogger(),
                           config=VoiceForgeConfig(
                               transfer_number="+18885551212"),
                           transfer_check=lambda n: False)
        agent.inbound_call("CA1", "+1000")
        out = agent.caller_said("CA1", "human please")
        self.assertIn("call you back", out.lower())

    def test_handoff_context(self):
        agent = VoiceAgent(self.brain, logger=CallLogger())
        agent.inbound_call("CA1", "+1000")
        agent.caller_said("CA1", "book please")
        agent.caller_said("CA1", "Ravi")
        ctx = agent.build_handoff_context(agent.calls["CA1"])
        self.assertIn("Ravi", ctx)
        self.assertIn("Test Co.", ctx)

    def test_outbound_campaign_facade(self):
        csv_data = "phone,name\n+15550001111,Ann\n+15550002222,Ben\n"
        report = self.agent.outbound_campaign(
            csv_data, opener="Hi {name}!", throttle_per_min=60000,
            dnc=["+15550002222"], on_transcript=lambda o: "what are your hours?",
            sleep=lambda s: None)
        by_phone = {r["phone"]: r["status"] for r in report}
        self.assertEqual(by_phone["+15550001111"], "completed")
        self.assertEqual(by_phone["+15550002222"], "skipped_dnc")

    def test_campaign_rejects_bad_csv(self):
        from voiceforge import CampaignError
        with self.assertRaises(CampaignError):
            self.agent.outbound_campaign("name\nAnn\n", opener="Hi",
                                         sleep=lambda s: None)
        with self.assertRaises(CampaignError):
            self.agent.outbound_campaign("phone,name\nnot-a-phone,Ann\n",
                                         opener="Hi", sleep=lambda s: None)

    def test_summary_counts(self):
        self.agent.inbound_call("CA200", "+1")
        self.agent.outbound_campaign(
            "phone,name\n+15550005555,Eve\n", opener="Hi {name}",
            throttle_per_min=60000, on_transcript=lambda o: "bye",
            sleep=lambda s: None)
        summary = self.agent.summary()
        self.assertEqual(summary["inbound"], 1)
        self.assertEqual(summary["outbound"], 1)

    def test_ai_disclosure_in_greeting(self):
        g = self.agent.inbound_call("CA9", "+1000")
        self.assertIn("AI assistant", g)
        self.assertIn("recorded", g)


# ---------------------------------------------------------------------------
# Vision: mms_watch
# ---------------------------------------------------------------------------

def make_watch(**kw):
    tmp = tempfile.mkdtemp()
    tw = TwilioAdapter(auth_token="s3cret")
    kw.setdefault("twilio", tw)
    kw.setdefault("logger", CallLogger())
    kw.setdefault("media_dir", tmp)
    kw.setdefault("downloader", lambda url: (b"FAKEIMG", "image/jpeg"))
    w = MmsWatch(**kw)
    w._sleep = lambda s: None
    return w, tmp


def mms_form(sid="SM1", num_media="1", ctype="image/jpeg"):
    return {"MessageSid": sid, "From": "+15550001111", "NumMedia": num_media,
            "MediaUrl0": "http://example.com/m.jpg",
            "MediaContentType0": ctype}


class TestMmsWatch(unittest.TestCase):
    def signed(self, watch, form, url="https://example.com/mms"):
        return twilio_sig("s3cret", url, form)

    def test_bad_signature_rejected(self):
        watch, _ = make_watch()
        form = mms_form()
        self.assertIsNone(watch.handle_webhook(
            "https://example.com/mms", form, "BADSIG"))

    def test_body_only_text_ignored(self):
        watch, _ = make_watch()
        form = mms_form(num_media="0")
        event = watch.handle_webhook(
            "https://example.com/mms", form,
            self.signed(watch, form))
        self.assertIsNone(event)

    def test_happy_path_downloads_and_persists(self):
        watch, tmp = make_watch()
        form = mms_form()
        event = watch.handle_webhook(
            "https://example.com/mms", form, self.signed(watch, form))
        self.assertIsNotNone(event)
        self.assertEqual(len(event.media), 1)
        self.assertTrue(os.path.exists(event.media[0].local_path))

    def test_duplicate_sid_is_noop(self):
        watch, _ = make_watch()
        form = mms_form()
        sig = self.signed(watch, form)
        self.assertIsNotNone(watch.handle_webhook(
            "https://example.com/mms", form, sig))
        self.assertIsNone(watch.handle_webhook(
            "https://example.com/mms", form, sig))

    def test_non_image_rejected(self):
        watch, _ = make_watch()
        form = mms_form(ctype="video/mp4")
        event = watch.handle_webhook(
            "https://example.com/mms", form, self.signed(watch, form))
        self.assertIsNone(event)

    def test_oversize_rejected(self):
        def big(url):
            raise ValueError("media exceeds 5000000 bytes")
        watch, _ = make_watch(downloader=big)
        form = mms_form()
        event = watch.handle_webhook(
            "https://example.com/mms", form, self.signed(watch, form))
        self.assertIsNone(event)

    def test_attach_matches_normalized_phone(self):
        watch, _ = make_watch()
        agent = VoiceAgent(make_brain(), logger=CallLogger())
        agent.inbound_call("CA1", "+1-555-000-1111")  # formatted differently
        form = mms_form()
        event = watch.handle_webhook(
            "https://example.com/mms", form, self.signed(watch, form))
        sid = watch.attach_to_session(agent, event)
        self.assertEqual(sid, "CA1")
        self.assertEqual(len(agent.calls["CA1"].media), 1)

    def test_orphan_reaped_on_later_call(self):
        watch, _ = make_watch()
        agent = VoiceAgent(make_brain(), logger=CallLogger())
        form = mms_form()
        form["From"] = "+19998887777"
        event = watch.handle_webhook(
            "https://example.com/mms", form, self.signed(watch, form))
        self.assertIsNone(watch.attach_to_session(agent, event))
        agent.inbound_call("CA2", "+19998887777")
        self.assertEqual(watch.reap_orphans(agent), 1)
        self.assertEqual(len(agent.calls["CA2"].media), 1)

    def test_landline_detection(self):
        watch, _ = make_watch(line_lookup=lambda p: "landline")
        self.assertTrue(watch.is_landline("+15550001111"))
        watch2, _ = make_watch(line_lookup=lambda p: "mobile")
        self.assertFalse(watch2.is_landline("+15550001111"))

    def test_line_lookup_failure_fails_open(self):
        def boom(p):
            raise RuntimeError("lookup down")
        watch, _ = make_watch(line_lookup=boom)
        self.assertFalse(watch.is_landline("+15550001111"))


# ---------------------------------------------------------------------------
# Vision: analyze (0.6 gate)
# ---------------------------------------------------------------------------

def vision_ctx():
    return VisionContext(business="Test Co.", recent_turns=[],
                         photo_request="the broken screen")


class TestVisionAnalyze(unittest.TestCase):
    def analyzer(self, response=None, **kw):
        client = MockVisionClient(response=response)
        return VisionAnalyzer(client, CallLogger(), **kw), client

    def test_gate_passes_at_exactly_0_6(self):
        az, _ = self.analyzer({"description": "d", "findings": ["f"],
                               "confidence": 0.6})
        res = az.analyze(b"img", "image/jpeg", vision_ctx())
        self.assertTrue(res.ok and res.gate_passed)

    def test_gate_fails_at_0_59_with_question(self):
        az, _ = self.analyzer({"description": "A blurry panel.",
                               "findings": ["Panel looks scorched."],
                               "confidence": 0.59})
        res = az.analyze(b"img", "image/jpeg", vision_ctx())
        self.assertTrue(res.ok)
        self.assertFalse(res.gate_passed)
        self.assertTrue(res.clarification_question)
        self.assertNotIn("scorched", res.clarification_question)

    def test_timeout_never_raises(self):
        client = MockVisionClient(exc=TimeoutError("slow"))
        az = VisionAnalyzer(client, CallLogger())
        res = az.analyze(b"img", "image/jpeg", vision_ctx())
        self.assertFalse(res.ok)
        self.assertEqual(res.error, "vision_timeout")

    def test_malformed_model_json_fails(self):
        az, _ = self.analyzer({"nonsense": True})
        res = az.analyze(b"img", "image/jpeg", vision_ctx())
        self.assertFalse(res.ok)
        self.assertIn("schema", res.error)

    def test_unexpected_exception_never_raises(self):
        client = MockVisionClient(exc=RuntimeError("weird"))
        az = VisionAnalyzer(client, CallLogger())
        res = az.analyze(b"img", "image/jpeg", vision_ctx())
        self.assertFalse(res.ok)


# ---------------------------------------------------------------------------
# Vision: context_inject
# ---------------------------------------------------------------------------

def good_result():
    return VisionResult(ok=True, description="Cracked screen.",
                        findings=["Glass cracked."],
                        confidence=0.9, gate_passed=True)


def low_result():
    return VisionResult(ok=True, description="Blurry.",
                        findings=["Maybe a crack."],
                        confidence=0.4, gate_passed=False,
                        clarification_question="Is that a crack or a smudge?")


class TestContextInject(unittest.TestCase):
    def setUp(self):
        self.inj = ContextInjector(CallLogger())

    def test_inject_appends_system_turn(self):
        s = new_session()
        out = self.inj.inject(s, good_result(), "k1")
        self.assertEqual(out.status, "injected")
        self.assertTrue(s.turns[-1].text.startswith("[vision:"))

    def test_duplicate_dedupe_key_dropped(self):
        s = new_session()
        self.inj.inject(s, good_result(), "k1")
        out = self.inj.inject(s, good_result(), "k1")
        self.assertEqual((out.status, out.reason), ("dropped", "duplicate"))
        vision_turns = [t for t in s.turns if t.text.startswith("[vision:")]
        self.assertEqual(len(vision_turns), 1)

    def test_dropped_after_hangup(self):
        s = new_session()
        s.hangup()
        out = self.inj.inject(s, good_result(), "k1")
        self.assertEqual((out.status, out.reason), ("dropped", "call_ended"))

    def test_deferred_mid_speech_then_flushed(self):
        s = new_session()
        out = self.inj.inject(s, good_result(), "k1", agent_speaking=True)
        self.assertEqual(out.status, "deferred")
        # nothing spoken yet
        self.assertFalse(any(t.text.startswith("[vision:")
                             for t in s.turns))
        spoken = self.inj.flush_pending(s)
        self.assertIsNone(spoken)  # gated result goes to LLM silently
        self.assertTrue(any(t.text.startswith("[vision:")
                            for t in s.turns))

    def test_low_confidence_flush_asks_question(self):
        s = new_session()
        self.inj.inject(s, low_result(), "k1", agent_speaking=True)
        spoken = self.inj.flush_pending(s)
        self.assertEqual(spoken, "Is that a crack or a smudge?")
        # findings never asserted into turns
        self.assertFalse(any("Maybe a crack" in t.text for t in s.turns))

    def test_failed_result_flush_speaks_fallback(self):
        s = new_session()
        self.inj.inject(s, VisionResult(ok=False, error="vision_timeout"),
                        "k1", agent_speaking=True)
        spoken = self.inj.flush_pending(s)
        self.assertEqual(spoken, VOICE_ONLY_FALLBACK)

    def test_flush_with_nothing_pending(self):
        self.assertIsNone(self.inj.flush_pending(new_session()))


# ---------------------------------------------------------------------------
# Vision: end-to-end flow
# ---------------------------------------------------------------------------

class TestVisionFlow(unittest.TestCase):
    def make_flow(self, clock_val=None, line_lookup=None):
        watch, tmp = make_watch(
            line_lookup=line_lookup or (lambda p: "mobile"))
        logger = CallLogger()
        analyzer = VisionAnalyzer(MockVisionClient(), logger)
        injector = ContextInjector(logger)
        now = [1000.0]
        clock = lambda: now[0]  # noqa: E731
        flow = VisionFlow(watch, analyzer, injector, logger, clock=clock)
        return flow, watch, now, tmp

    def make_agent(self, flow=None):
        return VoiceAgent(make_brain(), logger=CallLogger(),
                          vision_flow=flow)

    def test_landline_skips_visual_flow(self):
        flow, _, _, _ = self.make_flow(
            line_lookup=lambda p: "landline")
        agent = self.make_agent(flow)
        agent.inbound_call("CA1", "+15550001111")
        ask = agent.request_photo("CA1", "the damaged panel")
        self.assertEqual(ask, "")
        self.assertEqual(agent.calls["CA1"].vision_state, "unavailable")

    def test_request_photo_sets_deadline(self):
        flow, _, now, _ = self.make_flow()
        agent = self.make_agent(flow)
        agent.inbound_call("CA1", "+15550001111")
        ask = agent.request_photo("CA1", "the damaged panel")
        self.assertIn("photo", ask)
        self.assertEqual(agent.calls["CA1"].vision_state, "waiting_media")
        self.assertEqual(agent.calls["CA1"].vision_deadline_ts, 1090.0)

    def test_deadline_fires_fallback_once(self):
        flow, _, now, _ = self.make_flow()
        agent = self.make_agent(flow)
        agent.inbound_call("CA1", "+15550001111")
        agent.request_photo("CA1", "the panel")
        now[0] += 100.0  # past the 90 s budget
        fell = flow.check_deadlines(agent)
        self.assertEqual(fell, ["CA1"])
        # next turn speaks the single fallback line
        out = agent.caller_said("CA1", "hello?")
        self.assertEqual(out, VOICE_ONLY_FALLBACK)
        # and only once
        out2 = agent.caller_said("CA1", "still there?")
        self.assertNotEqual(out2, VOICE_ONLY_FALLBACK)

    def test_mms_webhook_bad_signature(self):
        flow, _, _, _ = self.make_flow()
        agent = self.make_agent(flow)
        self.assertFalse(flow.on_mms_webhook(
            agent, "https://example.com/mms", mms_form(), "BADSIG"))

    def test_mms_to_vision_end_to_end(self):
        flow, watch, _, _ = self.make_flow()
        agent = self.make_agent(flow)
        agent.inbound_call("CA1", "+15550001111")
        agent.request_photo("CA1", "the cracked screen")
        form = mms_form()
        url = "https://example.com/mms"
        sig = twilio_sig("s3cret", url, form)
        self.assertTrue(flow.on_mms_webhook(agent, url, form, sig))
        session = agent.calls["CA1"]
        self.assertEqual(session.vision_state, "ready")
        # next caller turn: gated result injected for the LLM (silent system
        # turn), agent reply continues the call normally
        out = agent.caller_said("CA1", "did you get it?")
        self.assertIsInstance(out, str)
        self.assertTrue(any(t.text.startswith("[vision:")
                            for t in session.turns))


# ---------------------------------------------------------------------------
# Config / latency budgets
# ---------------------------------------------------------------------------

class TestConfig(unittest.TestCase):
    def test_p0_budgets(self):
        c = VoiceForgeConfig()
        self.assertEqual(c.llm_timeout_s, 8.0)
        self.assertEqual(c.turn_deadline_s, 12.0)
        self.assertEqual(c.vision_timeout_s, 45.0)
        self.assertEqual(c.vision_budget_s, 90.0)
        self.assertEqual(c.min_confidence, 0.6)

    def test_normalize_phone(self):
        self.assertEqual(normalize_phone("+1-555-000-0001"), "5550000001")
        self.assertEqual(normalize_phone("+15550000001"), "5550000001")
        self.assertEqual(normalize_phone("1 (555) 000-0002"), "5550000002")
        self.assertEqual(normalize_phone("abc"), "")


# ---------------------------------------------------------------------------
# Review-gate fixes: one adversarial probe per blocker (B1-B7 + nits)
# ---------------------------------------------------------------------------

from voiceforge.session import VISION_WAITING_MEDIA


def _agent_with_vision(**kw):
    """VoiceAgent with a fully mocked vision stack (no network)."""
    brain = make_brain()
    agent = VoiceAgent(brain, logger=CallLogger(), **kw)
    tw = TwilioAdapter(auth_token="t")
    watch = MmsWatch(tw, CallLogger(), media_dir=tempfile.mkdtemp(),
                     downloader=lambda url: (b"IMG", "image/jpeg"))
    watch._sleep = lambda s: None
    flow = VisionFlow(watch,
                      VisionAnalyzer(MockVisionClient(), CallLogger()),
                      ContextInjector(CallLogger()), CallLogger())
    agent.vision_flow = flow
    return agent


def _signed_mms_form(secret, sid="SMW", from_phone="+1555"):
    form = {"MessageSid": sid, "From": from_phone, "NumMedia": "1",
            "MediaUrl0": "http://x/1.jpg",
            "MediaContentType0": "image/jpeg"}
    url = "http://x/hook"
    payload = url + "".join(k + form[k] for k in sorted(form))
    sig = base64.b64encode(
        hmac.new(secret.encode(), payload.encode(),
                 hashlib.sha1).digest()).decode()
    return url, form, sig


class _AuthRecorder:
    """Tiny localhost HTTP server that records the Authorization header."""

    def __init__(self):
        import http.server

        seen = {}
        outer = self

        class H(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                seen["auth"] = self.headers.get("Authorization")
                body = b"IMGDATA"
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, *a):
                pass

        self.seen = seen
        self.srv = http.server.HTTPServer(("127.0.0.1", 0), H)
        self.thread = threading.Thread(target=self.srv.serve_forever,
                                       daemon=True)
        self.thread.start()

    @property
    def url(self):
        return f"http://127.0.0.1:{self.srv.server_port}/x.jpg"

    def close(self):
        self.srv.shutdown()
        self.srv.server_close()


class TestB1MediaBasicAuth(unittest.TestCase):
    """B1: Twilio media download must carry HTTP Basic Auth."""

    def test_default_downloader_sends_basic_auth(self):
        from voiceforge.vision.mms_watch import default_downloader
        rec = _AuthRecorder()
        try:
            data, ctype = default_downloader(rec.url, auth=("AC123", "tok456"))
            self.assertEqual(data, b"IMGDATA")
            self.assertEqual(ctype, "image/jpeg")
            expect = "Basic " + base64.b64encode(b"AC123:tok456").decode()
            self.assertEqual(rec.seen.get("auth"), expect)
        finally:
            rec.close()

    def test_mms_watch_default_downloader_uses_adapter_creds(self):
        from voiceforge.vision.mms_watch import MmsWatch
        rec = _AuthRecorder()
        try:
            tw = TwilioAdapter(auth_token="tok456", account_sid="AC123")
            watch = MmsWatch(tw, CallLogger(),
                             media_dir=tempfile.mkdtemp())
            data, _ = watch.downloader(rec.url)
            self.assertEqual(data, b"IMGDATA")
            expect = "Basic " + base64.b64encode(b"AC123:tok456").decode()
            self.assertEqual(rec.seen.get("auth"), expect)
        finally:
            rec.close()

    def test_no_retry_on_401(self):
        # N9: a 401 will never succeed on retry — fail fast.
        calls = []

        def boom(url):
            calls.append(url)
            raise urllib.error.HTTPError(url, 401, "Unauthorized", {}, None)

        tw = TwilioAdapter(auth_token="x")
        watch = MmsWatch(tw, CallLogger(), media_dir=tempfile.mkdtemp(),
                         downloader=boom)
        watch._sleep = lambda s: None
        with self.assertRaises(urllib.error.HTTPError):
            watch._download("http://x/y.jpg")
        self.assertEqual(len(calls), 1)


class TestB7PhotoWiring(unittest.TestCase):
    """B7: vision must be reachable from a real conversation."""

    def test_photo_offer_triggers_vision_step(self):
        agent = _agent_with_vision()
        agent.inbound_call("CA1", "+15550001111")
        reply = agent.caller_said(
            "CA1", "Can I send you a photo of the cracked windshield?")
        s = agent.calls["CA1"]
        self.assertEqual(s.vision_state, VISION_WAITING_MEDIA)
        self.assertIn("photo", reply.lower())

    def test_photo_offer_without_vision_flow_degrades(self):
        brain = make_brain()
        s = new_session()
        reply = brain.handle(s, "I'll text you a photo of the leak")
        self.assertIn("describe", reply.lower())

    def test_photo_subject_extraction(self):
        self.assertEqual(
            ConversationBrain._photo_subject(
                "let me send you a photo of the cracked windshield"),
            "cracked windshield")
        self.assertEqual(ConversationBrain._photo_subject("sending pic"),
                         "the issue")


class TestB2UtteranceKeptOnFlush(unittest.TestCase):
    """B2: the caller utterance must survive vision-flush turns."""

    def test_caller_turn_recorded_when_flush_speaks(self):
        agent = _agent_with_vision()
        agent.inbound_call("CA1", "+15550001111")
        agent.request_photo("CA1", "the screen")
        s = agent.calls["CA1"]
        low = VisionResult(ok=True, description="Blurry.",
                           findings=["Maybe a crack."], confidence=0.4,
                           gate_passed=False,
                           clarification_question="Is that a crack?")
        agent._injector.inject(s, low, "k-low", agent_speaking=True)
        reply = agent.caller_said("CA1", "Did you get the photo?")
        self.assertIn("crack", reply.lower())
        caller_texts = [t.text for t in s.turns if t.role == "caller"]
        self.assertIn("Did you get the photo?", caller_texts)


class TestB3NoDictRaces(unittest.TestCase):
    """B3: concurrent webhook threads must never 500 on dict mutation."""

    def test_concurrent_mutate_and_iterate(self):
        agent = _agent_with_vision()
        watch = agent.vision_flow.mms_watch
        errors = []
        stop = threading.Event()

        def churn():
            i = 0
            while not stop.is_set():
                try:
                    agent.inbound_call(f"CAX{i}", f"+1555000{i:04d}")
                except RuntimeError as exc:
                    errors.append(exc)
                i += 1

        def iterate():
            while not stop.is_set():
                try:
                    agent.vision_flow.check_deadlines(agent)
                    agent.summary()
                    watch.reap_orphans(agent)
                except RuntimeError as exc:
                    errors.append(exc)

        threads = [threading.Thread(target=churn),
                   threading.Thread(target=iterate)]
        for t in threads:
            t.start()
        time.sleep(0.5)
        stop.set()
        for t in threads:
            t.join()
        self.assertEqual(errors, [])


class TestB4DayCorrection(unittest.TestCase):
    """B4: 'yes, friday instead' must re-resolve, never mis-commit."""

    def test_yes_friday_instead_reconfirms(self):
        # TODAY (test fixture) is Sunday 2026-10-04; tomorrow = Mon 2026-10-05
        brain = make_brain()
        s = new_session()
        for u in ("book please", "Ravi", "oil change", "tomorrow"):
            brain.handle(s, u)
        self.assertEqual(s.slots["_booking_step"], "confirm")
        r = brain.handle(s, "yes, friday instead")
        self.assertIn("friday", r.lower())
        self.assertIn("correct", r.lower())
        self.assertEqual(s.slots["_booking_step"], "confirm")
        self.assertEqual(s.slots["day_iso"], "2026-10-09")
        r2 = brain.handle(s, "yes")
        self.assertIn("2026-10-09", r2)

    def test_plain_yes_still_commits(self):
        brain = make_brain()
        s = new_session()
        for u in ("book please", "Ravi", "oil change", "tomorrow"):
            brain.handle(s, u)
        r = brain.handle(s, "yes")
        self.assertIn("booked", r.lower())
        self.assertEqual(s.slots["day"], "2026-10-05")


class TestB5TransferCheckGuarded(unittest.TestCase):
    """B5: a flaky carrier check must fall back, never 500."""

    def test_raising_transfer_check_takes_message(self):
        def boom(number):
            raise RuntimeError("carrier exploded")

        brain = make_brain()
        config = VoiceForgeConfig(transfer_number="+1999")
        agent = VoiceAgent(brain, logger=CallLogger(), config=config,
                           transfer_check=boom)
        agent.inbound_call("CA1", "+1555")
        out = agent.caller_said("CA1", "get me a human agent")
        self.assertIn("call you back", out.lower())


class TestB6WriteFailureGraceful(unittest.TestCase):
    """B6: disk failure on media write must degrade, never 500."""

    def test_oserror_on_write_returns_none(self):
        tw = TwilioAdapter(auth_token="t")
        watch = MmsWatch(tw, CallLogger(), media_dir=tempfile.mkdtemp(),
                         downloader=lambda url: (b"IMG", "image/jpeg"))
        url, form, sig = _signed_mms_form("t")
        with mock.patch("builtins.open", side_effect=OSError("disk full")):
            event = watch.handle_webhook(url, form, sig)
        self.assertIsNone(event)  # graceful degradation, no raise


class TestC1NoneTranscript(unittest.TestCase):
    """C1: a missing/empty transcript must never 500 the voice webhook."""

    def test_caller_said_none_returns_reply(self):
        brain, agent = make_brain(), VoiceAgent(make_brain(),
                                                logger=CallLogger())
        agent.inbound_call("CA900", "+1999")
        reply = agent.caller_said("CA900", None)  # must not raise
        self.assertIsInstance(reply, str)
        self.assertTrue(len(reply) > 0)

    def test_brain_handle_none_returns_reply(self):
        brain = make_brain()
        s = CallSession(call_sid="CA901", phone="+1000",
                        direction="inbound", state="active")
        reply = brain.handle(s, None)  # must not raise
        self.assertIsInstance(reply, str)


class TestN2True45sCap(unittest.TestCase):
    """N2: the 45 s vision budget is end-to-end, retry included."""

    def test_attempt_budget_sums_to_cap(self):
        seen_timeouts = []

        class Flaky(MockVisionClient):
            def describe(self, image_bytes, content_type, prompt,
                         timeout_s):
                seen_timeouts.append(timeout_s)
                if len(seen_timeouts) == 1:
                    raise TimeoutError("slow")
                return {"description": "OK.", "findings": ["Fine."],
                        "confidence": 0.9}

        az = VisionAnalyzer(Flaky(), CallLogger(), timeout_s=45.0)
        ctx = VisionContext(business="T", recent_turns=[],
                            photo_request="x")
        res = az.analyze(b"img", "image/jpeg", ctx)
        self.assertTrue(res.ok)
        self.assertEqual(len(seen_timeouts), 2)
        self.assertLessEqual(2 * seen_timeouts[0] + 1.0, 45.0)


class TestN3OrphanPurge(unittest.TestCase):
    """N3: RETENTION.md's deletion promises are actually implemented."""

    def test_expired_orphan_files_deleted(self):
        from voiceforge.session import MediaItem
        from voiceforge.vision.mms_watch import MmsEvent

        d = tempfile.mkdtemp()
        tw = TwilioAdapter(auth_token="t")
        watch = MmsWatch(tw, CallLogger(), media_dir=d)
        watch._sleep = lambda s: None
        p = os.path.join(d, "SMX_0.jpg")
        with open(p, "wb") as fh:
            fh.write(b"IMG")
        item = MediaItem(media_sid="SMX_0", message_sid="SMX",
                         content_type="image/jpeg", size=3, local_path=p)
        event = MmsEvent(message_sid="SMX", from_phone="+1999", media=[item])
        watch.orphans["SMX"] = (event, time.time() - 1)  # already expired
        agent = VoiceAgent(make_brain(), logger=CallLogger())
        watch.reap_orphans(agent)
        self.assertNotIn("SMX", watch.orphans)
        self.assertFalse(os.path.exists(p))

    def test_purge_media_files_by_age(self):
        d = tempfile.mkdtemp()
        tw = TwilioAdapter(auth_token="t")
        watch = MmsWatch(tw, CallLogger(), media_dir=d)
        old = os.path.join(d, "old.jpg")
        new = os.path.join(d, "new.jpg")
        for p in (old, new):
            with open(p, "wb") as fh:
                fh.write(b"x")
        ancient = time.time() - 10_000
        os.utime(old, (ancient, ancient))
        removed = watch.purge_media_files(max_age_s=3600)
        self.assertEqual(removed, 1)
        self.assertFalse(os.path.exists(old))
        self.assertTrue(os.path.exists(new))


class TestN4GateEnforcedInInject(unittest.TestCase):
    """N4: inject() itself enforces the 0.6 gate — never by convention."""

    def test_direct_inject_below_gate_never_asserts(self):
        inj = ContextInjector(CallLogger())
        s = new_session()
        low = VisionResult(ok=True, description="Blurry.",
                           findings=["Maybe a crack."], confidence=0.4,
                           gate_passed=False,
                           clarification_question="Is that a crack?")
        out = inj.inject(s, low, "k-direct")
        self.assertEqual(out.status, "deferred")
        self.assertFalse(any(t.text.startswith("[vision:")
                             for t in s.turns))
        spoken = inj.flush_pending(s)
        self.assertEqual(spoken, "Is that a crack?")


class TestDemoBookingRegression(unittest.TestCase):
    """SELL lane: the demo's booking flow visibly failed at runtime."""

    def test_tow_booked_starts_booking(self):
        brain = make_brain()
        s = new_session()
        r = brain.handle(s, "I need a tow booked please")
        self.assertIn("full name", r.lower())
        self.assertEqual(s.slots["_booking_step"], "name")

    def test_scheduled_triggers_booking(self):
        brain = make_brain()
        s = new_session()
        brain.handle(s, "I want to get my tires scheduled")
        self.assertEqual(s.slots["_booking_step"], "name")


class TestPilotSummary(unittest.TestCase):
    """Buyer-visible money-back metric: every call answered and logged."""

    def test_metric_counts(self):
        agent = VoiceAgent(make_brain(), logger=CallLogger())
        agent.inbound_call("CA1", "+1555")
        agent.caller_said("CA1", "what are your hours?")
        s = agent.pilot_summary()
        self.assertEqual(s["calls_received"], 1)
        self.assertEqual(s["calls_answered"], 1)
        self.assertTrue(s["guarantee_met"])
        text = agent.format_pilot_report()
        self.assertIn("100%", text)
        self.assertIn("Money-back metric", text)


if __name__ == "__main__":
    unittest.main(verbosity=1)
