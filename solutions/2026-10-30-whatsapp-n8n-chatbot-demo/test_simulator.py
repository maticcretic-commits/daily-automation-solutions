#!/usr/bin/env python3
"""Tests for the WhatsApp + n8n chatbot demo simulator.

Run:  python3 test_simulator.py -v

Covers:
  1. All 5 sample payloads produce a reply.
  2. Unknown intent falls back to a human-handoff message.
  3. Malformed payloads are rejected safely (no exception escapes).
  4. No credential-looking value appears in any demo file.
  5. The workflow JSON parses and the send node never targets a real
     WhatsApp/Meta host.
"""

import json
import re
import unittest
from pathlib import Path

from simulator import classify_intent, mock_responder, process_payload

HERE = Path(__file__).resolve().parent
PAYLOADS_DIR = HERE / "payloads"

# Patterns for credential-looking VALUES (not mere mentions of names).
# These must not match documentation that only names an environment
# variable (e.g. the literal name of a variable the live deployment
# would need) or a mock placeholder expression.
SECRET_VALUE_PATTERNS = [
    re.compile(r"sk-[A-Za-z0-9_-]{8,}"),                    # live key prefix
    re.compile(r"ghp_[A-Za-z0-9]{8,}"),                     # token prefix
    re.compile(r"xox[baprs]-[A-Za-z0-9-]{8,}"),             # token prefix
    re.compile(r"Bearer\s+[A-Za-z0-9._-]{10,}"),            # bearer value
    re.compile(r"\b\d{10,}\b"),                             # long digit run
    re.compile(r"\+\d{7,}"),                                # phone-style run
    re.compile(r"[\"'](?:api[_-]?key|access[_-]?token|client[_-]?secret)"
               r"[\"']\s*:\s*[\"'][^\"']{8,}[\"']"),         # assigned value
    re.compile(r"\b[0-9a-f]{32,}\b"),                       # long hex string
]


class TestSimulator(unittest.TestCase):
    def test_all_five_payloads_produce_a_reply(self):
        files = sorted(PAYLOADS_DIR.glob("*.json"))
        self.assertEqual(len(files), 5, "expected exactly 5 sample payloads")
        for path in files:
            data = json.loads(path.read_text(encoding="utf-8"))
            result = process_payload(data)
            self.assertTrue(result["ok"], "payload rejected: %s" % path.name)
            self.assertTrue(result["outbound"]["body"])
            self.assertTrue(result["outbound"]["simulated"])
            self.assertEqual(result["outbound"]["responder"], "MOCK")

    def test_unknown_intent_falls_back_to_human_handoff(self):
        data = json.loads((PAYLOADS_DIR / "05-unknown.json").read_text())
        result = process_payload(data)
        self.assertTrue(result["ok"])
        self.assertEqual(result["intent"], "human")
        self.assertFalse(result["understood"])
        self.assertIn("human", result["outbound"]["body"].lower())
        self.assertIn("did not understand", result["outbound"]["body"].lower())

    def test_known_intents_classify_correctly(self):
        self.assertEqual(classify_intent("What is the price?")[0], "price")
        self.assertEqual(
            classify_intent("What are your opening hours?")[0], "hours")
        self.assertEqual(
            classify_intent("Let me speak to a human")[0], "human")

    def test_mock_responder_is_labelled_and_template_only(self):
        reply = mock_responder("price", True)
        self.assertIn("Demo Traders", reply)
        self.assertNotIn("http", reply)

    def test_malformed_payloads_rejected_safely(self):
        malformed = [
            {},
            {"entry": []},
            {"entry": [{"changes": []}]},
            {"entry": [{"changes": [{"value": {}}]}]},
            {"entry": [{"changes": [{"value": {"messages": [{"from": ""}]}}]}]},
            {"entry": [{"changes": [{"value": {"messages": [
                {"from": "demo-customer-001", "text": {"body": "   "}}]}}]}]},
            "not a dict",
            None,
        ]
        for bad in malformed:
            result = process_payload(bad)
            self.assertFalse(result["ok"], "accepted bad input: %r" % (bad,))
            self.assertIn("error", result)

    def test_workflow_json_parses_and_has_required_nodes(self):
        workflow = json.loads((HERE / "workflow.json").read_text())
        names = [node["name"] for node in workflow["nodes"]]
        for required in ("WhatsApp Webhook (Demo)",
                         "Verify and Parse Payload",
                         "Intent Router",
                         "Reply Composer (Mock Responder)",
                         "Send via Mock Endpoint"):
            self.assertIn(required, names)
        self.assertIn("connections", workflow)

    def test_send_node_never_targets_real_whatsapp_host(self):
        workflow = json.loads((HERE / "workflow.json").read_text())
        raw = json.dumps(workflow)
        self.assertNotIn("graph.facebook.com", raw)
        self.assertNotIn("graph.whatsapp.com", raw)
        send = [n for n in workflow["nodes"]
                if n["name"] == "Send via Mock Endpoint"][0]
        url = send["parameters"]["url"]
        self.assertIn(".invalid", url)
        self.assertIn("mock", url)

    def test_no_credential_looking_values_in_demo_files(self):
        checked = []
        for path in sorted(HERE.rglob("*")):
            if path.is_dir() or "__pycache__" in path.parts:
                continue
            if path.name == "VERIFY.txt":
                continue  # generated from test output; holds no values
            if path.suffix.lower() not in (".py", ".json", ".md", ".svg"):
                continue
            text = path.read_text(encoding="utf-8", errors="ignore")
            checked.append(str(path.relative_to(HERE)))
            for pattern in SECRET_VALUE_PATTERNS:
                match = pattern.search(text)
                self.assertIsNone(
                    match,
                    "credential-looking value in %s: %r"
                    % (path.name, match.group(0) if match else ""))
        self.assertTrue(checked, "no demo files were scanned")


if __name__ == "__main__":
    unittest.main(verbosity=2)
