"""
LeadFlow — lead capture webhook pipeline.

Receives leads (e.g. from a website form / landing page), then:
  1. validates the payload (incl. a honeypot anti-bot field),
  2. normalises + de-duplicates by email (SQLite store),
  3. scores the lead 0-100 and tiers it hot/warm/cold,
  4. routes hot leads to the CRM immediately, everything else to nurture.

Standard library only — no pip installs needed.

Usage:
  python leadflow.py serve --port 8000        # POST JSON leads to /lead
  python leadflow.py export --tier hot        # export leads to CSV
  python leadflow.py report                   # print a summary report
"""

import argparse
import csv
import json
import re
import sqlite3
import sys
import time
from dataclasses import dataclass, asdict
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

from crm import ConsoleCRMClient, WebhookCRMClient

DB_PATH = "leads.db"

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
FREE_PROVIDERS = {
    "gmail.com", "yahoo.com", "yahoo.in", "hotmail.com", "outlook.com",
    "live.com", "aol.com", "icloud.com", "rediffmail.com",
}
HIGH_INTENT_SOURCES = {"referral", "demo-request", "pricing-page", "webinar"}


@dataclass
class Lead:
    name: str
    email: str
    phone: str = ""
    company: str = ""
    source: str = ""
    budget: float = 0.0
    message: str = ""
    utm_campaign: str = ""
    score: int = 0
    tier: str = "cold"
    submission_count: int = 1
    created_at: float = 0.0
    updated_at: float = 0.0


# ---------------------------------------------------------------- validation

def validate_lead(data: dict) -> list:
    """Return a list of validation error strings (empty = valid)."""
    errors = []
    if not data.get("name") or not str(data["name"]).strip():
        errors.append("name is required")
    email = str(data.get("email", "")).strip()
    if not email:
        errors.append("email is required")
    elif not EMAIL_RE.match(email):
        errors.append("email is not a valid email address")
    phone = str(data.get("phone", "")).strip()
    if phone:
        digits = re.sub(r"\D", "", phone)
        if len(digits) < 7 or len(digits) > 15:
            errors.append("phone number looks invalid")
    # Honeypot: real forms leave this hidden field empty; bots fill it.
    if str(data.get("website", "")).strip():
        errors.append("bot detected (honeypot filled)")
    if data.get("consent") is False:
        errors.append("consent is required")
    return errors


def normalize_email(email: str) -> str:
    """Lowercase/strip; drop gmail +tags and dots for reliable de-duping."""
    local, _, domain = email.strip().lower().partition("@")
    if domain == "gmail.com":
        local = local.split("+", 1)[0].replace(".", "")
    return f"{local}@{domain}"


# ---------------------------------------------------------------- storage

SCHEMA = """
CREATE TABLE IF NOT EXISTS leads (
    email TEXT PRIMARY KEY,
    name TEXT, phone TEXT, company TEXT, source TEXT,
    budget REAL, message TEXT, utm_campaign TEXT,
    score INTEGER, tier TEXT, submission_count INTEGER,
    created_at REAL, updated_at REAL
)
"""


class LeadStore:
    def __init__(self, path: str = DB_PATH):
        # check_same_thread=False: the HTTP server serves requests on its own
        # thread while the store is created at startup. Writes are still
        # serialised by the GIL + sqlite's own locking.
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)

    def upsert(self, lead: Lead) -> str:
        """Insert or update by normalised email. Returns 'created'|'updated'."""
        cur = self.conn.execute("SELECT * FROM leads WHERE email = ?", (lead.email,))
        row = cur.fetchone()
        now = time.time()
        if row is None:
            lead.created_at = lead.updated_at = now
            lead.submission_count = 1
            cols = ", ".join(asdict(lead).keys())
            self.conn.execute(
                f"INSERT INTO leads ({cols}) VALUES ({','.join('?' * len(asdict(lead)))})",
                tuple(asdict(lead).values()),
            )
            action = "created"
        else:
            lead.submission_count = row["submission_count"] + 1
            lead.updated_at = now
            # Keep the newest non-empty values, but never lose history.
            for field in ("name", "phone", "company", "source", "message", "utm_campaign"):
                if not getattr(lead, field):
                    setattr(lead, field, row[field])
            lead.budget = lead.budget or row["budget"]
            self.conn.execute(
                """UPDATE leads SET name=?, phone=?, company=?, source=?, budget=?,
                   message=?, utm_campaign=?, score=?, tier=?,
                   submission_count=?, updated_at=? WHERE email=?""",
                (lead.name, lead.phone, lead.company, lead.source, lead.budget,
                 lead.message, lead.utm_campaign, lead.score, lead.tier,
                 lead.submission_count, lead.updated_at, lead.email),
            )
            action = "updated"
        self.conn.commit()
        return action

    def all(self, tier: str | None = None) -> list:
        q = "SELECT * FROM leads" + (" WHERE tier = ?" if tier else "") + " ORDER BY score DESC"
        return [dict(r) for r in self.conn.execute(q, (tier,) if tier else ())]


# ---------------------------------------------------------------- scoring

def score_lead(lead: Lead) -> tuple[int, str]:
    """Rule-based fit score 0-100 and hot/warm/cold tier."""
    s = 0
    domain = lead.email.split("@", 1)[-1]
    if domain not in FREE_PROVIDERS:
        s += 30  # company email = business intent
    if lead.budget >= 1000:
        s += 20
    elif lead.budget >= 100:
        s += 10
    if lead.source in HIGH_INTENT_SOURCES:
        s += 15
    if lead.phone.strip():
        s += 10
    if len(lead.message.strip()) > 30:
        s += 10  # took time to write = engaged
    if lead.utm_campaign.strip():
        s += 5
    s = min(s, 100)
    tier = "hot" if s >= 60 else ("warm" if s >= 35 else "cold")
    return s, tier


# ---------------------------------------------------------------- pipeline

class LeadPipeline:
    """validate -> dedupe/store -> score -> route to CRM."""

    def __init__(self, store: LeadStore | None = None, crm_client=None):
        self.store = store or LeadStore()
        self.crm = crm_client or ConsoleCRMClient()

    def handle(self, payload: dict) -> dict:
        errors = validate_lead(payload)
        if errors:
            return {"ok": False, "errors": errors}

        email = normalize_email(payload["email"])
        lead = Lead(
            name=str(payload["name"]).strip(),
            email=email,
            phone=str(payload.get("phone", "")).strip(),
            company=str(payload.get("company", "")).strip(),
            source=str(payload.get("source", "")).strip().lower(),
            budget=float(payload.get("budget") or 0),
            message=str(payload.get("message", "")).strip(),
            utm_campaign=str(payload.get("utm_campaign", "")).strip(),
        )
        lead.score, lead.tier = score_lead(lead)
        action = self.store.upsert(lead)

        routed = None
        if lead.tier == "hot":
            routed = self.crm.create_or_update_contact(lead)
        return {
            "ok": True,
            "action": action,
            "email": email,
            "score": lead.score,
            "tier": lead.tier,
            "crm_routed": routed is not None,
        }


# ---------------------------------------------------------------- CLI / server

class Handler(BaseHTTPRequestHandler):
    pipeline: LeadPipeline = None  # set in serve()

    def _send(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if urlparse(self.path).path != "/lead":
            return self._send(404, {"ok": False, "error": "use POST /lead"})
        length = int(self.headers.get("Content-Length", 0))
        try:
            payload = json.loads(self.rfile.read(length) or b"{}")
        except json.JSONDecodeError:
            return self._send(400, {"ok": False, "error": "invalid JSON"})
        result = self.pipeline.handle(payload)
        self._send(200 if result["ok"] else 422, result)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/leads":
            return self._send(200, {"leads": self.pipeline.store.all()})
        if path == "/report":
            leads = self.pipeline.store.all()
            tiers = {"hot": 0, "warm": 0, "cold": 0}
            for l in leads:
                tiers[l["tier"]] = tiers.get(l["tier"], 0) + 1
            return self._send(200, {"total": len(leads), "by_tier": tiers})
        return self._send(404, {"ok": False, "error": "try /lead, /leads, /report"})

    def log_message(self, *args):
        pass  # keep console clean


def serve(port: int, webhook_url: str | None):
    crm = WebhookCRMClient(webhook_url) if webhook_url else ConsoleCRMClient()
    Handler.pipeline = LeadPipeline(crm_client=crm)
    srv = HTTPServer(("127.0.0.1", port), Handler)
    print(f"LeadFlow listening on http://127.0.0.1:{port}  (POST /lead)")
    print("Hot leads are pushed to the CRM; warm/cold go to nurture.")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")


def export_csv(tier: str | None, out: str):
    leads = LeadStore().all(tier)
    if not leads:
        print("No leads match.")
        return
    with open(out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(leads[0].keys()))
        w.writeheader()
        w.writerows(leads)
    print(f"Exported {len(leads)} lead(s) to {out}")


def report():
    leads = LeadStore().all()
    tiers = {"hot": 0, "warm": 0, "cold": 0}
    for l in leads:
        tiers[l["tier"]] = tiers.get(l["tier"], 0) + 1
    print(f"Total leads: {len(leads)}")
    for t, n in tiers.items():
        print(f"  {t}: {n}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="LeadFlow lead-capture pipeline")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("serve")
    s.add_argument("--port", type=int, default=8000)
    s.add_argument("--crm-webhook", default=None,
                   help="Optional HTTPS endpoint to POST hot leads to (a real CRM webhook)")
    e = sub.add_parser("export")
    e.add_argument("--tier", choices=["hot", "warm", "cold"], default=None)
    e.add_argument("--out", default="leads-export.csv")
    sub.add_parser("report")
    args = ap.parse_args(argv)

    if args.cmd == "serve":
        serve(args.port, args.crm_webhook)
    elif args.cmd == "export":
        export_csv(args.tier, args.out)
    elif args.cmd == "report":
        report()


if __name__ == "__main__":
    main()
