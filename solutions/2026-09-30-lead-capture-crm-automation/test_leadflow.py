"""Tests for the LeadFlow pipeline. Run:  python test_leadflow.py"""

import os
import tempfile
import unittest

from crm import ConsoleCRMClient, HubSpotPayloadBuilder
from leadflow import (
    Lead,
    LeadPipeline,
    LeadStore,
    normalize_email,
    score_lead,
    validate_lead,
)


def fresh_pipeline():
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    os.unlink(tmp.name)  # let sqlite create it fresh
    crm = ConsoleCRMClient()
    return LeadPipeline(store=LeadStore(tmp.name), crm_client=crm), crm, tmp.name


GOOD = {
    "name": "Priya Sharma",
    "email": "priya@acme-logistics.com",
    "phone": "+91 98765 43210",
    "company": "Acme Logistics",
    "source": "demo-request",
    "budget": 2500,
    "message": "We need lead automation for our 5-person sales team urgently.",
    "utm_campaign": "september-launch",
    "consent": True,
}


class TestValidation(unittest.TestCase):
    def test_valid_payload_passes(self):
        self.assertEqual(validate_lead(GOOD), [])

    def test_missing_name_and_bad_email(self):
        errs = validate_lead({"name": " ", "email": "not-an-email"})
        self.assertIn("name is required", errs)
        self.assertIn("email is not a valid email address", errs)

    def test_honeypot_catches_bots(self):
        bad = dict(GOOD, website="http://spam.example")
        errs = validate_lead(bad)
        self.assertTrue(any("bot" in e for e in errs))

    def test_consent_required(self):
        errs = validate_lead(dict(GOOD, consent=False))
        self.assertIn("consent is required", errs)


class TestNormalize(unittest.TestCase):
    def test_gmail_aliases_collapse(self):
        self.assertEqual(normalize_email("Priya.S+spam@gmail.com"),
                         normalize_email("priyas@gmail.com"))

    def test_case_insensitive(self):
        self.assertEqual(normalize_email("A@B.COM"), "a@b.com")


class TestScoring(unittest.TestCase):
    def test_hot_corporate_demo_request(self):
        lead = Lead(name="X", email="x@bigcorp.com", source="demo-request",
                    budget=2500, phone="1234567",
                    message="a" * 50, utm_campaign="q3")
        score, tier = score_lead(lead)
        self.assertEqual(tier, "hot")
        self.assertGreaterEqual(score, 60)

    def test_cold_free_email_no_signals(self):
        lead = Lead(name="Y", email="y@gmail.com")
        score, tier = score_lead(lead)
        self.assertEqual(tier, "cold")
        self.assertLess(score, 35)


class TestPipeline(unittest.TestCase):
    def test_end_to_end_hot_lead_reaches_crm(self):
        pipe, crm, _ = fresh_pipeline()
        result = pipe.handle(dict(GOOD))
        self.assertTrue(result["ok"])
        self.assertEqual(result["action"], "created")
        self.assertEqual(result["tier"], "hot")
        self.assertTrue(result["crm_routed"])
        self.assertEqual(len(crm.synced), 1)
        self.assertEqual(crm.synced[0]["email"], "priya@acme-logistics.com")

    def test_duplicate_email_updates_not_duplicates(self):
        pipe, crm, _ = fresh_pipeline()
        pipe.handle(dict(GOOD))
        second = dict(GOOD, budget=5000, message="Follow-up with bigger budget.")
        result = pipe.handle(second)
        self.assertEqual(result["action"], "updated")
        rows = pipe.store.all()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["submission_count"], 2)
        self.assertEqual(rows[0]["budget"], 5000)

    def test_cold_lead_not_routed_to_crm(self):
        pipe, crm, _ = fresh_pipeline()
        result = pipe.handle({"name": "Zed", "email": "zed@yahoo.com"})
        self.assertTrue(result["ok"])
        self.assertEqual(result["tier"], "cold")
        self.assertFalse(result["crm_routed"])
        self.assertEqual(crm.synced, [])

    def test_invalid_payload_rejected(self):
        pipe, _, _ = fresh_pipeline()
        result = pipe.handle({"name": "", "email": "bad"})
        self.assertFalse(result["ok"])
        self.assertTrue(result["errors"])


class TestHubSpotBuilder(unittest.TestCase):
    def test_builds_contact_properties(self):
        lead = Lead(name="Priya Sharma", email="priya@acme-logistics.com",
                    phone="123", company="Acme", source="demo-request",
                    budget=2500, message="hi", score=80, tier="hot")
        payload = HubSpotPayloadBuilder.build(lead)
        props = payload["properties"]
        self.assertEqual(props["email"], "priya@acme-logistics.com")
        self.assertEqual(props["firstname"], "Priya")
        self.assertEqual(props["hubspotscore"], 80)


if __name__ == "__main__":
    unittest.main(verbosity=2)
