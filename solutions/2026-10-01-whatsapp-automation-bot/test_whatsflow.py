"""Tests for WhatsFlow. Run: python3 test_whatsflow.py"""

import hashlib
import hmac
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import whatsflow
from whatsflow import (
    BotEngine, BotReply, Broadcaster, ConsoleSender, InboundMessage,
    button_payload, check_signature, get_db, normalise_phone, parse_incoming,
    template_payload, text_payload, verify_webhook,
)


def fresh_db():
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    conn = get_db(Path(tmp.name))
    return conn


class TestWebhookVerification(unittest.TestCase):
    def test_success_returns_challenge(self):
        status, body = verify_webhook(
            {"hub.mode": ["subscribe"], "hub.verify_token": ["secret123"],
             "hub.challenge": ["CHALLENGE_TOKEN"]},
            "secret123",
        )
        self.assertEqual(status, 200)
        self.assertEqual(body, "CHALLENGE_TOKEN")

    def test_wrong_token_rejected(self):
        status, _ = verify_webhook(
            {"hub.mode": ["subscribe"], "hub.verify_token": ["wrong"],
             "hub.challenge": ["X"]},
            "secret123",
        )
        self.assertEqual(status, 403)

    def test_missing_challenge_rejected(self):
        status, _ = verify_webhook(
            {"hub.mode": ["subscribe"], "hub.verify_token": ["secret123"]},
            "secret123",
        )
        self.assertEqual(status, 403)


class TestSignature(unittest.TestCase):
    def test_valid_signature(self):
        body = b'{"hello":"world"}'
        sig = "sha256=" + hmac.new(b"appsecret", body, hashlib.sha256).hexdigest()
        self.assertTrue(check_signature(body, "appsecret", sig))

    def test_tampered_body_rejected(self):
        body = b'{"hello":"world"}'
        sig = "sha256=" + hmac.new(b"appsecret", body, hashlib.sha256).hexdigest()
        self.assertFalse(check_signature(b'{"hello":"tampered"}', "appsecret", sig))

    def test_missing_header_rejected(self):
        self.assertFalse(check_signature(b"{}", "appsecret", None))
        self.assertFalse(check_signature(b"{}", "appsecret", "garbage"))


class TestParsing(unittest.TestCase):
    def test_extracts_text_messages(self):
        payload = {
            "object": "whatsapp_business_account",
            "entry": [{
                "changes": [{
                    "value": {
                        "contacts": [{"wa_id": "919999999999",
                                      "profile": {"name": "Asha"}}],
                        "messages": [{
                            "from": "919999999999", "id": "wamid.1",
                            "type": "text", "text": {"body": "Hi, pricing please"},
                        }],
                    },
                }],
            }],
        }
        msgs = parse_incoming(payload)
        self.assertEqual(len(msgs), 1)
        self.assertEqual(msgs[0].phone, "919999999999")
        self.assertEqual(msgs[0].name, "Asha")
        self.assertEqual(msgs[0].body, "Hi, pricing please")

    def test_ignores_non_text_and_statuses(self):
        payload = {"entry": [{"changes": [{"value": {
            "messages": [{"from": "9199", "id": "w1", "type": "image",
                          "image": {"id": "img"}}],
            "statuses": [{"id": "w1", "status": "delivered"}],
        }}]}]}
        self.assertEqual(parse_incoming(payload), [])


class TestBotEngine(unittest.TestCase):
    def setUp(self):
        self.conn = fresh_db()
        self.engine = BotEngine(self.conn)

    def tearDown(self):
        self.conn.close()

    def ask(self, phone, text, name=""):
        return self.engine.handle(InboundMessage(phone, name, text, "m", "text"))

    def test_greeting_then_menu(self):
        reply = self.ask("+91111", "hello", "Ravi")
        self.assertIn("Hello Ravi", reply.text)
        self.assertIn("Pricing", reply.text)

    def test_pricing_faq_with_buttons(self):
        reply = self.ask("+91112", "What is your pricing?")
        self.assertIn("$49/month", reply.text)
        self.assertTrue(reply.buttons)

    def test_numbered_menu_shortcuts(self):
        self.assertIn("$49/month", self.ask("+91113", "1").text)
        self.assertIn("name", self.ask("+91114", "2").text.lower())
        self.assertIn("Mon", self.ask("+91115", "3").text)
        self.assertIn("human agent", self.ask("+91116", "4").text.lower())

    def test_demo_lead_capture_flow(self):
        r1 = self.ask("+91117", "I want a demo")
        self.assertIn("name", r1.text.lower())
        r2 = self.ask("+91117", "Priya Sharma")
        self.assertIn("Priya", r2.text)
        r3 = self.ask("+91117", "Pricing for my team")
        self.assertIn("reach out", r3.text)
        row = self.conn.execute(
            "SELECT name, interest, state FROM contacts WHERE phone='+91117'"
        ).fetchone()
        self.assertEqual(row["name"], "Priya Sharma")
        self.assertEqual(row["state"], "captured")
        self.assertIn("Pricing", row["interest"])

    def test_fallback_offers_menu(self):
        reply = self.ask("+91118", "asdfghjkl")
        self.assertIn("didn't quite get that", reply.text)
        self.assertIn("Pricing", reply.text)

    def test_opt_out_and_resubscribe(self):
        self.assertIn("unsubscribed", self.ask("+91119", "STOP").text)
        reply = self.ask("+91119", "hello again")
        self.assertIn("opted out", reply.text)
        self.assertIn("START", reply.text)
        reply = self.ask("+91119", "START")
        self.assertIn("subscribed again", reply.text)

    def test_human_handoff(self):
        self.assertIn("human agent", self.ask("+91120", "talk to a human").text.lower())
        row = self.conn.execute(
            "SELECT state FROM contacts WHERE phone='+91120'").fetchone()
        self.assertEqual(row["state"], "handoff")

    def test_messages_logged(self):
        self.ask("+91121", "hi")
        n = self.conn.execute(
            "SELECT COUNT(*) c FROM messages WHERE phone='+91121'").fetchone()["c"]
        self.assertEqual(n, 2)  # one in, one out


class TestPayloads(unittest.TestCase):
    def test_text_payload_shape(self):
        p = text_payload("+91123", "Hello")
        self.assertEqual(p["messaging_product"], "whatsapp")
        self.assertEqual(p["to"], "+91123")
        self.assertEqual(p["type"], "text")
        self.assertEqual(p["text"]["body"], "Hello")

    def test_template_payload_with_variables(self):
        p = template_payload("+91123", "hello_world", variables=["Ravi"])
        comps = p["template"]["components"]
        self.assertEqual(comps[0]["parameters"][0]["text"], "Ravi")

    def test_button_payload_caps_at_three(self):
        p = button_payload("+91123", "Pick one", ["A", "B", "C", "D"])
        self.assertEqual(len(p["interactive"]["action"]["buttons"]), 3)


class TestBroadcaster(unittest.TestCase):
    def setUp(self):
        self.conn = fresh_db()
        self.sender = ConsoleSender()
        # fast rate for tests: 1000/sec
        self.bc = Broadcaster(self.conn, self.sender, rate_per_second=1000)

    def tearDown(self):
        self.conn.close()

    def test_dedupe_optout_and_invalid(self):
        self.conn.execute(
            "INSERT INTO contacts (phone, name, state, opted_out, created_at, updated_at)"
            " VALUES ('+919999900001', 'X', 'new', 1, 0, 0)"
        )
        self.conn.commit()
        summary = self.bc.run(
            ["+919999900002", "+919999900002", "+919999900001", "not-a-phone"],
            "hello_world",
        )
        self.assertEqual(summary["sent"], 1)
        self.assertEqual(summary["skipped_dupe"], 1)
        self.assertEqual(summary["skipped_optout"], 1)
        self.assertEqual(summary["skipped_invalid"], 1)
        self.assertEqual(len(self.sender.sent), 1)

    def test_broadcast_log_written(self):
        self.bc.run(["+919999900003"], "hello_world")
        row = self.conn.execute(
            "SELECT status FROM broadcast_log WHERE phone='+919999900003'"
        ).fetchone()
        self.assertEqual(row["status"], "sent")


class TestPhoneNormalisation(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(normalise_phone("919999999999"), "+919999999999")
        self.assertEqual(normalise_phone("+1 415 555 2671"), "+14155552671")
        self.assertEqual(normalise_phone("+44 20 7946 0018"), "+442079460018")

    def test_invalid(self):
        self.assertIsNone(normalise_phone("abc"))
        self.assertIsNone(normalise_phone("123"))  # too short
        self.assertIsNone(normalise_phone(""))


class TestSampleWebhookEndToEnd(unittest.TestCase):
    """Replays a captured webhook payload through parse -> engine -> sender."""

    def test_replay(self):
        conn = fresh_db()
        engine = BotEngine(conn)
        sender = ConsoleSender()
        raw = json.dumps({
            "entry": [{"changes": [{"value": {
                "contacts": [{"wa_id": "919876543210",
                              "profile": {"name": "Karan"}}],
                "messages": [{"from": "919876543210", "id": "wamid.9",
                              "type": "text", "text": {"body": "demo"}}],
            }}]}]
        }).encode()
        sig = "sha256=" + hmac.new(b"s3cr3t", raw, hashlib.sha256).hexdigest()
        assert check_signature(raw, "s3cr3t", sig)
        for msg in parse_incoming(json.loads(raw)):
            reply = engine.handle(msg)
            sender.send(text_payload(msg.phone, reply.text))
        self.assertEqual(len(sender.sent), 1)
        self.assertIn("name", sender.sent[0]["text"]["body"].lower())
        conn.close()


if __name__ == "__main__":
    unittest.main(verbosity=1)
