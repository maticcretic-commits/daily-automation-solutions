"""Tests for VoiceForge (voice_agent.py). Stdlib only: run with `python3 test_voice_agent.py`."""

import base64
import hashlib
import hmac
import unittest

from voice_agent import (
    CallSession, CallLogger, ConversationBrain, MockLLMClient, MockTTSClient,
    TwilioAdapter, VapiAdapter, VoiceAgent,
)


class TestConversationBrain(unittest.TestCase):
    def setUp(self):
        self.brain = ConversationBrain(business="Test Co.")

    def new_session(self):
        return CallSession(call_sid="CA1", phone="+1000", direction="inbound",
                           state="active")

    def test_greeting_mentions_business(self):
        s = self.new_session()
        text = self.brain.greet(s)
        self.assertIn("Test Co.", text)
        self.assertEqual(s.turns[-1].role, "agent")

    def test_faq_hours(self):
        s = self.new_session()
        reply = self.brain.handle(s, "What are your opening hours?")
        self.assertIn("9 AM", reply)

    def test_faq_pricing(self):
        s = self.new_session()
        reply = self.brain.handle(s, "How much does it cost?")
        self.assertIn("$49", reply)

    def test_escalation_flags_session(self):
        s = self.new_session()
        reply = self.brain.handle(s, "Let me talk to a human agent please")
        self.assertTrue(s.escalated)
        self.assertIn("human", reply.lower())

    def test_goodbye_hangs_up(self):
        s = self.new_session()
        reply = self.brain.handle(s, "Okay thanks, goodbye!")
        self.assertEqual(s.state, "ended")
        self.assertIsNotNone(s.ended_at)
        self.assertIn("great day", reply.lower())

    def test_booking_three_step_flow(self):
        s = self.new_session()
        r1 = self.brain.handle(s, "I want to book an appointment")
        self.assertIn("name", r1.lower())
        r2 = self.brain.handle(s, "Ravi Sharma")
        self.assertIn("service", r2.lower())
        self.assertEqual(s.slots["name"], "Ravi Sharma")
        r3 = self.brain.handle(s, "oil change")
        self.assertIn("day", r3.lower())
        r4 = self.brain.handle(s, "Monday")
        self.assertIn("booked", r4.lower())
        self.assertIn("Ravi Sharma", r4)
        self.assertEqual(s.slots["service"], "oil change")
        self.assertEqual(s.slots["day"], "Monday")

    def test_fallback_uses_llm(self):
        brain = ConversationBrain(
            business="Test Co.",
            llm=MockLLMClient({"fallback": "CUSTOM FALLBACK REPLY"}))
        s = self.new_session()
        reply = brain.handle(s, "tell me about quantum tunneling")
        self.assertEqual(reply, "CUSTOM FALLBACK REPLY")


class TestTwilioAdapter(unittest.TestCase):
    def setUp(self):
        self.twilio = TwilioAdapter(auth_token="secret123",
                                    gather_action_url="/voice")

    def test_signature_verifies(self):
        url = "https://example.com/voice"
        params = {"From": "+1555", "CallSid": "CA1"}
        payload = url + "".join(k + params[k] for k in sorted(params))
        sig = base64.b64encode(
            hmac.new(b"secret123", payload.encode(),
                     hashlib.sha1).digest()).decode()
        self.assertTrue(self.twilio.verify_signature(url, params, sig))

    def test_signature_rejects_tampered(self):
        url = "https://example.com/voice"
        params = {"From": "+1555"}
        self.assertFalse(
            self.twilio.verify_signature(url, params, "AAAA"))

    def test_inbound_twiml_gathers_speech(self):
        xml = self.twilio.inbound_twiml("Hello there")
        self.assertIn("<Gather", xml)
        self.assertIn('input="speech"', xml)
        self.assertIn("Hello there", xml)

    def test_continue_twiml_end_call_hangs_up(self):
        xml = self.twilio.continue_twiml("Bye!", end_call=True)
        self.assertIn("<Hangup/>", xml)
        xml2 = self.twilio.continue_twiml("Hi", end_call=False)
        self.assertIn("<Gather", xml2)


class TestVapiAdapter(unittest.TestCase):
    def test_verify_ok(self):
        v = VapiAdapter(webhook_secret="s3cret")
        self.assertTrue(v.verify({"x-vapi-secret": "s3cret"}))

    def test_verify_rejects_wrong(self):
        v = VapiAdapter(webhook_secret="s3cret")
        self.assertFalse(v.verify({"x-vapi-secret": "wrong"}))

    def test_parse_assistant_request(self):
        v = VapiAdapter(webhook_secret="x")
        out = v.parse_message({"message": {"type": "assistant-request"},
                               "call": {"id": "c1"}})
        self.assertEqual(out["action"], "start_call")
        self.assertEqual(out["call_id"], "c1")

    def test_parse_user_transcript(self):
        v = VapiAdapter(webhook_secret="x")
        out = v.parse_message(
            {"message": {"type": "transcript", "role": "user",
                         "transcript": "hello"},
             "call": {"id": "c2"}})
        self.assertEqual(out["action"], "caller_said")
        self.assertEqual(out["text"], "hello")

    def test_parse_unknown_ignored(self):
        v = VapiAdapter(webhook_secret="x")
        out = v.parse_message({"message": {"type": "weird-event"}})
        self.assertEqual(out["action"], "ignore")


class TestVoiceAgentFlows(unittest.TestCase):
    def setUp(self):
        brain = ConversationBrain(business="Demo Co.")
        self.agent = VoiceAgent(brain, logger=CallLogger())

    def test_inbound_full_flow(self):
        greeting = self.agent.inbound_call("CA100", "+1999")
        self.assertIn("Demo Co.", greeting)
        self.agent.caller_said("CA100", "what are your hours?")
        self.agent.caller_said("CA100", "book please")
        self.agent.caller_said("CA100", "Neha")
        self.agent.caller_said("CA100", "consult")
        self.agent.caller_said("CA100", "Friday")
        session = self.agent.calls["CA100"]
        self.assertEqual(session.slots["day"], "Friday")
        self.agent.caller_said("CA100", "bye!")
        self.assertEqual(session.state, "ended")

    def test_inactive_call_raises(self):
        self.agent.inbound_call("CA101", "+1998")
        self.agent.caller_said("CA101", "bye")
        with self.assertRaises(ValueError):
            self.agent.caller_said("CA101", "hello again?")

    def test_outbound_campaign_respects_dnc_and_no_answer(self):
        csv_data = "phone,name\n+111,Ann\n+222,Ben\n+333,Cid\n"
        report = self.agent.outbound_campaign(
            csv_data, opener="Hi {name}!",
            throttle_per_min=60000, dnc=["+222"],
            on_transcript=lambda opener: None if "Cid" in opener
            else "what are your hours?")
        by_phone = {r["phone"]: r["status"] for r in report}
        self.assertEqual(by_phone["+111"], "completed")
        self.assertEqual(by_phone["+222"], "skipped_dnc")
        # opener is "Hi Cid!" for the third row -> lambda returns None -> no answer
        self.assertEqual(by_phone["+333"], "no_answer")

    def test_outbound_no_answer_path(self):
        csv_data = "phone,name\n+444,Dan\n"
        report = self.agent.outbound_campaign(
            csv_data, opener="Hi {name}!", throttle_per_min=60000,
            on_transcript=lambda opener: None)
        self.assertEqual(report[0]["status"], "no_answer")
        self.assertEqual(self.agent.calls[report[0]["sid"]].state, "ended")

    def test_summary_counts(self):
        self.agent.inbound_call("CA200", "+1")
        self.agent.outbound_campaign("phone,name\n+555,Eve\n",
                                     opener="Hi {name}",
                                     throttle_per_min=60000,
                                     on_transcript=lambda o: "bye")
        summary = self.agent.summary()
        self.assertEqual(summary["inbound"], 1)
        self.assertEqual(summary["outbound"], 1)
        self.assertEqual(summary["total_calls"], 2)

    def test_tts_returns_audio_bytes(self):
        audio = MockTTSClient().speak("hello", voice="alloy")
        self.assertTrue(len(audio) > len("hello"))


if __name__ == "__main__":
    unittest.main(verbosity=1)
