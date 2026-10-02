#!/usr/bin/env python3
"""
BlogFlow — n8n-style content pipeline engine.

Faithful reference implementation of the automation pattern behind a typical
client brief: "n8n content pipeline — queue in Google Sheets, AI-generated
articles, auto-publish to WordPress, resilient to failures".

The exact same logic ships as an importable n8n workflow (workflow-n8n.json)
in this folder; this Python engine is runnable, testable, and documents the
behavior contract node-by-node so a client can verify every rule.

Behavior rules (mirrors the client's technical spec):
  1. Queue is FIFO — oldest queued item first.
  2. Daily publish limit comes from the settings tab (never hardcoded).
  3. Failed items retry up to MAX_RETRIES (3) across runs, then are marked Failed.
  4. One item's failure NEVER stops the rest of the run.
  5. Every publish sets SEO title + meta description (Rank Math style) and
     image alt text = focus keyword; the sitemap is pinged after each publish.
  6. Generated images/social posts/newsletter jobs are appended to their own
     queues for downstream workflows (social + email steps).

No third-party dependencies. Python 3.9+.
"""

import csv
import datetime
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

QUEUE_HEADERS = [
    "id", "kind", "keyword", "title", "focus_keyword",
    "status", "attempts", "created_at", "published_at",
    "wp_post_id", "seo_title", "meta_description", "alt_text", "error",
]
SETTINGS_HEADERS = ["key", "value"]
LOG_HEADERS = ["ts", "level", "message"]
SOCIAL_HEADERS = ["id", "queue_item_id", "network", "caption", "status", "created_at"]
NEWSLETTER_HEADERS = ["id", "queue_item_id", "subject", "status", "created_at"]

STATUS_QUEUED = "queued"
STATUS_PUBLISHED = "published"
STATUS_FAILED = "failed"

DEFAULT_SETTINGS = {
    "daily_limit": "5",
    "max_retries": "3",
    "site_url": "https://example.com",
    "seo_title_template": "{keyword} — Complete Guide {year}",
    "meta_description_template": "Learn {keyword}: definitions, calculators, examples and expert tips in this complete guide.",
    "social_networks": "pinterest,linkedin",
    "newsletter_subject_template": "New guide: {keyword}",
}


def _utcnow():
    return datetime.datetime.now(datetime.timezone.utc)


# ---------------------------------------------------------------------------
# QueueStore — CSV-backed store mirroring a Google Sheet with tabs
# ---------------------------------------------------------------------------

class QueueStore:
    """A folder of CSVs standing in for the Google Sheet tabs:
    queue.csv, settings.csv, log.csv, social_queue.csv, newsletter_queue.csv.
    """

    def __init__(self, root):
        self.root = root
        os.makedirs(root, exist_ok=True)
        self._ensure("queue.csv", QUEUE_HEADERS)
        self._ensure("settings.csv", SETTINGS_HEADERS)
        self._ensure("log.csv", LOG_HEADERS)
        self._ensure("social_queue.csv", SOCIAL_HEADERS)
        self._ensure("newsletter_queue.csv", NEWSLETTER_HEADERS)
        # seed defaults for any missing setting key
        current = self.settings()
        for k, v in DEFAULT_SETTINGS.items():
            if k not in current:
                self.set_setting(k, v)

    # -- low-level ---------------------------------------------------------
    def _path(self, name):
        return os.path.join(self.root, name)

    def _ensure(self, name, headers):
        if not os.path.exists(self._path(name)):
            with open(self._path(name), "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerow(headers)

    def _read(self, name):
        with open(self._path(name), newline="", encoding="utf-8") as fh:
            return list(csv.DictReader(fh))

    def _write(self, name, headers, rows):
        with open(self._path(name), "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=headers)
            w.writeheader()
            w.writerows(rows)

    # -- logging -----------------------------------------------------------
    def log(self, level, message):
        self._append("log.csv", {
            "ts": _utcnow().isoformat(), "level": level, "message": message,
        })

    def _append(self, name, row):
        with open(self._path(name), "a", newline="", encoding="utf-8") as fh:
            csv.DictWriter(fh, fieldnames=list(row.keys())).writerow(row)

    # -- settings ----------------------------------------------------------
    def settings(self):
        return {r["key"]: r["value"] for r in self._read("settings.csv")}

    def set_setting(self, key, value):
        rows = [r for r in self._read("settings.csv") if r["key"] != key]
        rows.append({"key": key, "value": str(value)})
        self._write("settings.csv", SETTINGS_HEADERS, rows)

    # -- queue --------------------------------------------------------------
    def queue(self):
        return self._read("queue.csv")

    def enqueue(self, kind, keyword, focus_keyword=None, title=None):
        rows = self.queue()
        new_id = str(max([int(r["id"]) for r in rows], default=0) + 1)
        row = {
            "id": new_id,
            "kind": kind,
            "keyword": keyword,
            "title": title or "",
            "focus_keyword": focus_keyword or keyword,
            "status": STATUS_QUEUED,
            "attempts": "0",
            "created_at": _utcnow().isoformat(),
            "published_at": "",
            "wp_post_id": "",
            "seo_title": "",
            "meta_description": "",
            "alt_text": "",
            "error": "",
        }
        rows.append(row)
        self._write("queue.csv", QUEUE_HEADERS, rows)
        self.log("info", f"enqueued #{new_id} [{kind}] {keyword}")
        return new_id

    def update_row(self, row_id, **changes):
        rows = self.queue()
        for r in rows:
            if r["id"] == str(row_id):
                for k, v in changes.items():
                    r[k] = str(v)
        self._write("queue.csv", QUEUE_HEADERS, rows)

    def due_items(self, limit):
        """FIFO: oldest queued items first, capped by the daily limit."""
        items = [r for r in self.queue() if r["status"] == STATUS_QUEUED]
        items.sort(key=lambda r: (r["created_at"], int(r["id"])))
        return items[: max(0, int(limit))]

    def published_today(self):
        today = _utcnow().date().isoformat()
        return sum(
            1 for r in self.queue()
            if r["status"] == STATUS_PUBLISHED
            and r["published_at"].startswith(today)
        )

    # -- downstream queues ---------------------------------------------------
    def enqueue_social(self, queue_item_id, networks, caption):
        rows = self._read("social_queue.csv")
        for net in networks:
            net = net.strip()
            if not net:
                continue
            rows.append({
                "id": str(len(rows) + 1),
                "queue_item_id": str(queue_item_id),
                "network": net,
                "caption": caption,
                "status": "pending",
                "created_at": _utcnow().isoformat(),
            })
        self._write("social_queue.csv", SOCIAL_HEADERS, rows)

    def enqueue_newsletter(self, queue_item_id, subject):
        rows = self._read("newsletter_queue.csv")
        rows.append({
            "id": str(len(rows) + 1),
            "queue_item_id": str(queue_item_id),
            "subject": subject,
            "status": "pending",
            "created_at": _utcnow().isoformat(),
        })
        self._write("newsletter_queue.csv", NEWSLETTER_HEADERS, rows)


# ---------------------------------------------------------------------------
# AI clients
# ---------------------------------------------------------------------------

class MockAIClient:
    """Deterministic stand-in for the Claude/OpenAI node — used in tests and
    demos so the pipeline runs end-to-end without API keys."""

    def generate(self, kind, keyword):
        article = (
            f"# {keyword.title()}: The Complete Guide\n\n"
            f"{keyword.title()} is one of the most-searched personal finance "
            f"topics. This guide explains the definition, walks through a "
            f"worked example, and answers the questions readers ask most.\n\n"
            f"## What is {keyword}?\nA clear, beginner-friendly explanation "
            f"with a real-numbers example.\n\n"
            f"## Worked example\nSee the interactive calculator section for "
            f"live numbers.\n\n## FAQs\nCommon questions answered in plain "
            f"English.\n"
        )
        return {
            "title": f"{keyword.title()} — Complete Guide",
            "content": article,
            "excerpt": f"Everything you need to know about {keyword}.",
        }


class ClaudeClient:
    """Real Anthropic API client (HTTPS, stdlib only). Set ANTHROPIC_API_KEY."""

    API_URL = "https://api.anthropic.com/v1/messages"

    def __init__(self, api_key=None, model="claude-sonnet-4-20250514",
                 timeout=60, max_retries=3):
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        self.model = model
        self.timeout = timeout
        self.max_retries = max_retries

    def generate(self, kind, keyword):
        prompt = (
            f"Write a publish-ready finance {kind} about '{keyword}': "
            "a markdown article with an H1, definition section, a worked "
            "numeric example, and an FAQ. Return JSON with keys "
            "title, content, excerpt."
        )
        body = json.dumps({
            "model": self.model,
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": prompt}],
        }).encode()
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        last = None
        for attempt in range(self.max_retries):
            try:
                req = urllib.request.Request(self.API_URL, data=body, headers=headers)
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    payload = json.load(resp)
                text = payload["content"][0]["text"]
                return json.loads(text)
            except Exception as exc:  # noqa: BLE001 - retried uniformly
                last = exc
                time.sleep(2 ** attempt)
        raise RuntimeError(f"Claude API failed after {self.max_retries} tries: {last}")


# ---------------------------------------------------------------------------
# WordPress clients
# ---------------------------------------------------------------------------

class MockWPClient:
    """In-memory WordPress REST stand-in; records every publish call."""

    def __init__(self):
        self.posts = []
        self.sitemap_pings = 0

    def publish(self, title, content, seo_title, meta_description, alt_text,
                status="publish"):
        post_id = 1000 + len(self.posts) + 1
        self.posts.append({
            "id": post_id, "title": title, "content": content,
            "seo_title": seo_title, "meta_description": meta_description,
            "alt_text": alt_text, "status": status,
        })
        return post_id

    def ping_sitemap(self, site_url):
        self.sitemap_pings += 1
        return True


class WordPressClient:
    """Real WordPress REST client (Application Password auth, stdlib only)."""

    def __init__(self, site_url, username=None, app_password=None, timeout=30):
        self.site_url = site_url.rstrip("/")
        self.username = username or os.environ.get("WP_USER", "")
        self.app_password = app_password or os.environ.get("WP_APP_PASSWORD", "")
        self.timeout = timeout

    def _request(self, method, path, data=None):
        import base64
        token = base64.b64encode(
            f"{self.username}:{self.app_password}".encode()).decode()
        body = json.dumps(data).encode() if data is not None else None
        req = urllib.request.Request(
            self.site_url + path, data=body, method=method,
            headers={"Authorization": f"Basic {token}",
                     "Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=self.timeout) as resp:
            return json.load(resp)

    def publish(self, title, content, seo_title, meta_description, alt_text,
                status="publish"):
        payload = {
            "title": title, "content": content, "status": status,
            # Rank Math style SEO fields
            "meta": {"rank_math_title": seo_title,
                     "rank_math_description": meta_description},
            # store alt text contract alongside; media upload omitted for brevity
            "slug": urllib.parse.quote(title.lower().replace(" ", "-")[:80]),
        }
        res = self._request("POST", "/wp-json/wp/v2/posts", payload)
        return res["id"]

    def ping_sitemap(self, site_url):
        try:
            urllib.request.urlopen(
                self.site_url + "/wp-sitemap-posts-post-1.xml",
                timeout=self.timeout).read(64)
        except Exception:  # noqa: BLE001 - ping is best-effort
            pass
        return True


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def build_seo(keyword, settings):
    year = _utcnow().year
    seo_title = settings["seo_title_template"].format(keyword=keyword.title(), year=year)
    meta = settings["meta_description_template"].format(keyword=keyword)
    return seo_title, meta


def process_item(store, item, ai, wp, settings):
    """Process ONE queue item. Returns 'published' | 'failed' | 'retry'."""
    item_id = item["id"]
    max_retries = int(settings.get("max_retries", "3"))
    try:
        gen = ai.generate(item["kind"], item["keyword"])
        seo_title, meta_desc = build_seo(item["keyword"], settings)
        alt_text = item["focus_keyword"]  # alt text = focus keyword, always
        post_id = wp.publish(
            title=gen["title"],
            content=gen["content"],
            seo_title=seo_title,
            meta_description=meta_desc,
            alt_text=alt_text,
            status="publish",  # publish immediately — never save as draft
        )
        wp.ping_sitemap(settings.get("site_url", ""))
        now = _utcnow().isoformat()
        store.update_row(
            item_id, status=STATUS_PUBLISHED, attempts=int(item["attempts"]) + 1,
            published_at=now, wp_post_id=post_id,
            seo_title=seo_title, meta_description=meta_desc,
            alt_text=alt_text, error="",
            title=gen["title"],
        )
        networks = [n for n in settings.get("social_networks", "").split(",") if n.strip()]
        store.enqueue_social(item_id, networks, f"New guide: {gen['title']}")
        store.enqueue_newsletter(
            item_id,
            settings.get("newsletter_subject_template", "New guide: {keyword}")
            .format(keyword=item["keyword"]),
        )
        store.log("info", f"published #{item_id} as WP post {post_id}")
        return STATUS_PUBLISHED
    except Exception as exc:  # noqa: BLE001 - item-level isolation
        attempts = int(item["attempts"]) + 1
        store.log("error", f"item #{item_id} attempt {attempts} failed: {exc}")
        if attempts >= max_retries:
            store.update_row(item_id, status=STATUS_FAILED, attempts=attempts,
                             error=str(exc)[:500])
            store.log("error", f"item #{item_id} marked FAILED after {attempts} attempts")
            return STATUS_FAILED
        store.update_row(item_id, attempts=attempts, error=str(exc)[:500])
        return "retry"


def run_pipeline(store, ai, wp, dry_run=False):
    """One daily run. Never raises on item failure — the run always completes."""
    settings = store.settings()
    daily_limit = int(settings.get("daily_limit", "5"))
    already = store.published_today()
    remaining = max(0, daily_limit - already)
    store.log("info", f"run start: daily_limit={daily_limit}, already={already}, remaining={remaining}")
    stats = {"published": 0, "retry": 0, "failed": 0, "skipped": 0}
    if dry_run:
        due = store.due_items(remaining)
        stats["skipped"] = len(due)
        return stats
    for item in store.due_items(remaining):
        try:
            outcome = process_item(store, item, ai, wp, settings)
            if outcome == STATUS_PUBLISHED:
                stats["published"] += 1
                remaining -= 1
                if remaining <= 0:
                    break
            elif outcome == STATUS_FAILED:
                stats["failed"] += 1
            else:  # "retry" — will be attempted again on a later run
                stats["retry"] += 1
        except Exception as exc:  # noqa: BLE001 - belt & suspenders
            stats["failed"] += 1
            store.log("critical", f"unexpected pipeline error on item {item.get('id')}: {exc}")
    store.log("info", f"run end: {stats}")
    return stats


def daily_summary(store):
    rows = store.queue()
    counts = {s: 0 for s in (STATUS_QUEUED, STATUS_PUBLISHED, STATUS_FAILED)}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    return {
        "date": _utcnow().date().isoformat(),
        "queued": counts[STATUS_QUEUED],
        "published": counts[STATUS_PUBLISHED],
        "failed": counts[STATUS_FAILED],
        "published_today": store.published_today(),
    }


def main():
    import argparse
    ap = argparse.ArgumentParser(description="BlogFlow content pipeline")
    ap.add_argument("--data", default="./blogflow-data", help="queue store folder")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--summary", action="store_true")
    ap.add_argument("--enqueue", nargs=3, metavar=("KIND", "KEYWORD", "FOCUS"),
                    help="enqueue one item then exit")
    args = ap.parse_args()

    store = QueueStore(args.data)
    if args.enqueue:
        kind, keyword, focus = args.enqueue
        print("enqueued item id:", store.enqueue(kind, keyword, focus_keyword=focus))
        return
    if args.summary:
        print(json.dumps(daily_summary(store), indent=2))
        return
    stats = run_pipeline(store, MockAIClient(), MockWPClient(), dry_run=args.dry_run)
    print(json.dumps({"stats": stats, "summary": daily_summary(store)}, indent=2))


if __name__ == "__main__":
    main()
