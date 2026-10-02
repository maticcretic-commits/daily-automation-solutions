#!/usr/bin/env python3
"""Tests for blogflow.py — run with: python3 test_blogflow.py"""

import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from blogflow import (  # noqa: E402
    QueueStore, MockAIClient, MockWPClient, run_pipeline, process_item,
    daily_summary, build_seo, STATUS_PUBLISHED, STATUS_FAILED, STATUS_QUEUED,
)


class ExplodingAI(MockAIClient):
    def generate(self, kind, keyword):
        raise RuntimeError("AI provider down")


class ExplodingWP(MockWPClient):
    def publish(self, *a, **k):
        raise RuntimeError("WordPress REST 500")


def fresh_store(**settings):
    tmp = tempfile.mkdtemp(prefix="blogflow-test-")
    store = QueueStore(tmp)
    for k, v in settings.items():
        store.set_setting(k, v)
    return store


class TestQueueStore(unittest.TestCase):
    def test_enqueue_defaults(self):
        s = fresh_store()
        iid = s.enqueue("article", "compound interest")
        row = [r for r in s.queue() if r["id"] == iid][0]
        self.assertEqual(row["status"], STATUS_QUEUED)
        self.assertEqual(row["attempts"], "0")
        self.assertEqual(row["focus_keyword"], "compound interest")

    def test_fifo_order(self):
        s = fresh_store()
        a = s.enqueue("article", "aaa first")
        b = s.enqueue("article", "bbb second")
        due = s.due_items(10)
        self.assertEqual([due[0]["id"], due[1]["id"]], [a, b])

    def test_daily_limit_cap(self):
        s = fresh_store(daily_limit="2")
        for i in range(5):
            s.enqueue("article", f"kw {i}")
        self.assertEqual(len(s.due_items(99)), 5)
        run = run_pipeline(s, MockAIClient(), MockWPClient())
        self.assertEqual(run["published"], 2)
        self.assertEqual(s.published_today(), 2)

    def test_settings_roundtrip(self):
        s = fresh_store()
        s.set_setting("daily_limit", "7")
        self.assertEqual(s.settings()["daily_limit"], "7")


class TestPublishing(unittest.TestCase):
    def test_success_sets_seo_fields_and_sitemap(self):
        s = fresh_store(site_url="https://finance.example.com")
        iid = s.enqueue("article", "emergency fund", focus_keyword="emergency fund")
        wp = MockWPClient()
        outcome = process_item(s, s.due_items(5)[0], MockAIClient(), wp, s.settings())
        self.assertEqual(outcome, STATUS_PUBLISHED)
        row = [r for r in s.queue() if r["id"] == iid][0]
        self.assertEqual(row["status"], STATUS_PUBLISHED)
        self.assertTrue(row["wp_post_id"].isdigit())
        self.assertIn("Emergency Fund", row["seo_title"])
        self.assertTrue(row["meta_description"])
        self.assertEqual(row["alt_text"], "emergency fund")  # alt = focus keyword
        self.assertEqual(wp.sitemap_pings, 1)

    def test_publish_immediately_not_draft(self):
        s = fresh_store()
        s.enqueue("article", "roth ira")
        wp = MockWPClient()
        process_item(s, s.due_items(5)[0], MockAIClient(), wp, s.settings())
        self.assertEqual(wp.posts[0]["status"], "publish")

    def test_social_and_newsletter_queued(self):
        s = fresh_store(social_networks="pinterest,linkedin")
        s.enqueue("article", "index funds")
        wp = MockWPClient()
        process_item(s, s.due_items(5)[0], MockAIClient(), wp, s.settings())
        with open(os.path.join(s.root, "social_queue.csv")) as fh:
            social = list(__import__("csv").DictReader(fh))
        nets = sorted(r["network"] for r in social)
        self.assertEqual(nets, ["linkedin", "pinterest"])
        with open(os.path.join(s.root, "newsletter_queue.csv")) as fh:
            news = list(__import__("csv").DictReader(fh))
        self.assertEqual(len(news), 1)
        self.assertIn("index funds", news[0]["subject"])

    def test_seo_builder_uses_templates(self):
        seo_title, meta = build_seo("tax harvesting",
                                    {"seo_title_template": "{keyword} {year}",
                                     "meta_description_template": "About {keyword}."})
        self.assertIn("Tax Harvesting", seo_title)
        self.assertIn("tax harvesting", meta)


class TestResilience(unittest.TestCase):
    def test_retry_then_fail_after_three(self):
        s = fresh_store(max_retries="3")
        s.enqueue("article", "doomed topic")
        ai = ExplodingAI()
        wp = MockWPClient()
        for expected in ("retry", "retry", STATUS_FAILED):
            item = [r for r in s.queue() if r["status"] == STATUS_QUEUED][0]
            outcome = process_item(s, item, ai, wp, s.settings())
            self.assertEqual(outcome, expected)
        row = s.queue()[0]
        self.assertEqual(row["status"], STATUS_FAILED)
        self.assertEqual(row["attempts"], "3")
        self.assertTrue(row["error"])

    def test_one_failure_never_stops_run(self):
        s = fresh_store(max_retries="3")
        s.enqueue("article", "good one")
        s.enqueue("article", "bad one")
        s.enqueue("article", "good two")

        class FlakyAI(MockAIClient):
            def generate(self, kind, keyword):
                if keyword == "bad one":
                    raise RuntimeError("boom")
                return super().generate(kind, keyword)

        stats = run_pipeline(s, FlakyAI(), MockWPClient())
        self.assertEqual(stats["published"], 2)
        self.assertEqual(stats["retry"], 1)

    def test_wp_failure_retries_item(self):
        s = fresh_store(max_retries="2")
        s.enqueue("article", "wp down")
        outcome = process_item(s, s.due_items(5)[0], MockAIClient(),
                              ExplodingWP(), s.settings())
        self.assertEqual(outcome, "retry")
        row = s.queue()[0]
        self.assertEqual(row["status"], STATUS_QUEUED)
        self.assertEqual(row["attempts"], "1")

    def test_empty_queue_noop(self):
        s = fresh_store()
        stats = run_pipeline(s, MockAIClient(), MockWPClient())
        self.assertEqual(stats, {"published": 0, "retry": 0, "failed": 0, "skipped": 0})

    def test_zero_daily_limit_publishes_nothing(self):
        s = fresh_store(daily_limit="0")
        s.enqueue("article", "ignored")
        stats = run_pipeline(s, MockAIClient(), MockWPClient())
        self.assertEqual(stats["published"], 0)
        self.assertEqual(s.queue()[0]["status"], STATUS_QUEUED)

    def test_dry_run_reports_without_publishing(self):
        s = fresh_store(daily_limit="2")
        s.enqueue("article", "dry one")
        wp = MockWPClient()
        stats = run_pipeline(s, MockAIClient(), wp, dry_run=True)
        self.assertEqual(stats["skipped"], 1)
        self.assertEqual(len(wp.posts), 0)

    def test_limit_never_hardcoded(self):
        s = fresh_store(daily_limit="1")
        for i in range(3):
            s.enqueue("article", f"kw{i}")
        run_pipeline(s, MockAIClient(), MockWPClient())
        self.assertEqual(s.published_today(), 1)
        s.set_setting("daily_limit", "3")
        run_pipeline(s, MockAIClient(), MockWPClient())
        self.assertEqual(s.published_today(), 3)


class TestSummaryAndLog(unittest.TestCase):
    def test_daily_summary_counts(self):
        s = fresh_store(max_retries="1")
        s.enqueue("article", "ok")
        s.enqueue("article", "bad")
        s.enqueue("article", "waiting")

        class FlakyAI(MockAIClient):
            def generate(self, kind, keyword):
                if keyword == "bad":
                    raise RuntimeError("boom")
                return super().generate(kind, keyword)

        run_pipeline(s, FlakyAI(), MockWPClient())
        summary = daily_summary(s)
        self.assertEqual(summary["published"], 2)
        self.assertEqual(summary["failed"], 1)
        self.assertEqual(summary["queued"], 0)
        self.assertEqual(summary["published_today"], 2)

    def test_log_captures_run(self):
        s = fresh_store()
        s.enqueue("article", "logged")
        run_pipeline(s, MockAIClient(), MockWPClient())
        with open(os.path.join(s.root, "log.csv")) as fh:
            messages = [r["message"] for r in __import__("csv").DictReader(fh)]
        self.assertTrue(any("run start" in m for m in messages))
        self.assertTrue(any("run end" in m for m in messages))


class TestRealClientsSanity(unittest.TestCase):
    def test_claude_client_builds_valid_request_shape(self):
        from blogflow import ClaudeClient
        c = ClaudeClient(api_key="sk-test")
        self.assertEqual(c.API_URL, "https://api.anthropic.com/v1/messages")
        self.assertEqual(c.max_retries, 3)

    def test_wp_client_publish_payload_shape(self):
        from blogflow import WordPressClient
        import base64
        captured = {}

        class FakeResp:
            def __enter__(self): return self
            def __exit__(self, *a): return False
            def read(self, n=-1): return b'{"id": 42}'

        import urllib.request as urlreq
        orig = urlreq.urlopen
        def fake(req, timeout=None):
            captured["url"] = req.full_url
            captured["auth"] = req.headers.get("Authorization")
            body = json.loads(req.data.decode())
            assert body["status"] == "publish"
            assert "rank_math_title" in body["meta"]
            return FakeResp()
        urlreq.urlopen = fake
        try:
            c = WordPressClient("https://site.example", username="u", app_password="p")
            self.assertEqual(
                c.publish("T", "C", "SEO", "META", "ALT"), 42)
            self.assertTrue(captured["url"].endswith("/wp-json/wp/v2/posts"))
            self.assertTrue(captured["auth"].startswith("Basic "))
        finally:
            urlreq.urlopen = orig


if __name__ == "__main__":
    unittest.main(verbosity=2)
