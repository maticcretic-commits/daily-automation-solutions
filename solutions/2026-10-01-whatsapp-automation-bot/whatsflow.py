"""
WhatsFlow — a working WhatsApp Business Cloud API automation bot.

A representative client problem (from real WhatsApp automation gigs):
a small business wants inbound WhatsApp messages answered automatically
(FAQ, lead capture, booking intent) and wants broadcast campaigns sent
with dedupe, opt-out handling and rate limiting — without writing any
Meta boilerplate.

What this does (stdlib only, no dependencies):
  * Webhook verification endpoint (Meta's GET handshake)
  * Inbound message handling with X-Hub-Signature-256 verification
  * Rule-based conversation engine: greeting -> menu -> FAQ -> lead
    capture (name + interest) -> human handoff; persisted per phone in SQLite
  * Outbound payloads for text, templates and interactive buttons
    (Cloud API format), sendable via the real API or a console mock
  * Campaign broadcaster: CSV import, phone dedupe, opt-out respect,
    per-second rate limit, send log in SQLite
  * CLI: `server` (webhook), `simulate` (try the bot offline), `broadcast`,
    `export` (leads CSV), `report`

Usage:
    python3 whatsflow.py simulate
    python3 whatsflow.py server --port 8000
    python3 whatsflow.py broadcast --csv campaign.csv --template hello_world
"""

from __future__ import annotations

import csv
import hashlib
import hmac
import json
import re
import sqlite3
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

DB_PATH = Path(__file__).with_name("whatsflow.db")

# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------


def get_db(path: Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS contacts (
            phone TEXT PRIMARY KEY,
            name TEXT,
            interest TEXT,
            state TEXT DEFAULT 'new',
            opted_out INTEGER DEFAULT 0,
            created_at REAL,
            updated_at REAL
        );
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT,
            direction TEXT,          -- 'in' | 'out'
            body TEXT,
            created_at REAL
        );
        CREATE TABLE IF NOT EXISTS broadcast_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT,
            template TEXT,
            status TEXT,             -- 'sent' | 'skipped_optout' | 'skipped_dupe'
            created_at REAL
        );
        """
    )
    return conn


# ---------------------------------------------------------------------------
# Webhook verification + signature check (Meta Cloud API contract)
# ---------------------------------------------------------------------------


def verify_webhook(query: dict, verify_token: str) -> tuple[int, str]:
    """Handle Meta's GET verification handshake.

    Returns (http_status, body). Meta expects the raw challenge back on 200.
    """
    mode = query.get("hub.mode", [""])[0]
    token = query.get("hub.verify_token", [""])[0]
    challenge = query.get("hub.challenge", [""])[0]
    if mode == "subscribe" and token == verify_token and challenge:
        return 200, challenge
    return 403, "verification failed"


def check_signature(raw_body: bytes, app_secret: str, signature_header: str | None) -> bool:
    """Validate the X-Hub-Signature-256 header on inbound webhooks."""
    if not signature_header or not signature_header.startswith("sha256="):
        return False
    expected = hmac.new(app_secret.encode(), raw_body, hashlib.sha256).hexdigest()
    return hmac.compare_digest("sha256=" + expected, signature_header)


# ---------------------------------------------------------------------------
# Inbound parsing
# ---------------------------------------------------------------------------


@dataclass
class InboundMessage:
    phone: str
    name: str
    body: str
    message_id: str
    msg_type: str


def parse_incoming(payload: dict) -> list[InboundMessage]:
    """Extract user messages from a WhatsApp Cloud API webhook payload."""
    out: list[InboundMessage] = []
    for entry in payload.get("entry", []):
        for change in entry.get("changes", []):
            value = change.get("value", {})
            contacts = {c.get("wa_id"): c.get("profile", {}).get("name", "")
                        for c in value.get("contacts", [])}
            for msg in value.get("messages", []):
                if msg.get("type") != "text":
                    continue
                phone = msg.get("from", "")
                out.append(
                    InboundMessage(
                        phone=phone,
                        name=contacts.get(phone, ""),
                        body=(msg.get("text") or {}).get("body", ""),
                        message_id=msg.get("id", ""),
                        msg_type="text",
                    )
                )
    return out


# ---------------------------------------------------------------------------
# Conversation engine
# ---------------------------------------------------------------------------

MENU = (
    "How can I help you today?\n"
    "1. Pricing\n"
    "2. Book a demo\n"
    "3. Business hours\n"
    "4. Talk to a human"
)

FAQ = {
    "pricing": (
        "Our plans start at $49/month (Starter), $149/month (Growth) and "
        "$399/month (Scale). All plans include the WhatsApp Business API "
        "connection and unlimited inbound messages. Reply DEMO to book a walkthrough."
    ),
    "hours": (
        "We're available Mon–Fri, 9 AM to 7 PM IST. Messages received outside "
        "these hours are answered first thing next business day."
    ),
    "demo": (
        "Great! I can book a 20-minute demo for you. What's your name?"
    ),
}

KEYWORD_MAP = [
    (re.compile(r"\b(price|pricing|cost|charge|fee|plan)\b", re.I), "pricing"),
    (re.compile(r"\b(hour|timing|open|close|when.*available)\b", re.I), "hours"),
    (re.compile(r"\b(demo|trial|book|schedule|appointment)\b", re.I), "demo"),
    (re.compile(r"\b(hi|hello|hey|namaste|good\s*(morning|afternoon|evening))\b", re.I), "greeting"),
    (re.compile(r"\b(menu|option|help)\b", re.I), "menu"),
    (re.compile(r"\b(human|agent|person|call me|support)\b", re.I), "human"),
    (re.compile(r"\b(stop|unsubscribe|opt.?out|don't.*message)\b", re.I), "stop"),
]


@dataclass
class BotReply:
    text: str
    buttons: list[str] = field(default_factory=list)


class BotEngine:
    """Stateful per-phone conversation engine backed by SQLite."""

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn

    # -- contact helpers --------------------------------------------------
    def _contact(self, phone: str, name: str = "") -> sqlite3.Row:
        row = self.conn.execute(
            "SELECT * FROM contacts WHERE phone = ?", (phone,)
        ).fetchone()
        if row is None:
            now = time.time()
            self.conn.execute(
                "INSERT INTO contacts (phone, name, state, created_at, updated_at)"
                " VALUES (?, ?, 'new', ?, ?)",
                (phone, name, now, now),
            )
            self.conn.commit()
            row = self.conn.execute(
                "SELECT * FROM contacts WHERE phone = ?", (phone,)
            ).fetchone()
        elif name and not row["name"]:
            self.conn.execute(
                "UPDATE contacts SET name = ?, updated_at = ? WHERE phone = ?",
                (name, time.time(), phone),
            )
            self.conn.commit()
        return row

    def _set_state(self, phone: str, state: str, **fields) -> None:
        sets = ", ".join([f"{k} = ?" for k in fields] + ["state = ?", "updated_at = ?"])
        self.conn.execute(
            f"UPDATE contacts SET {sets} WHERE phone = ?",
            (*fields.values(), state, time.time(), phone),
        )
        self.conn.commit()

    def _log(self, phone: str, direction: str, body: str) -> None:
        self.conn.execute(
            "INSERT INTO messages (phone, direction, body, created_at)"
            " VALUES (?, ?, ?, ?)",
            (phone, direction, body, time.time()),
        )
        self.conn.commit()

    # -- main entry -------------------------------------------------------
    def handle(self, msg: InboundMessage) -> BotReply:
        contact = self._contact(msg.phone, msg.name)
        self._log(msg.phone, "in", msg.body)
        text = msg.body.strip()

        if contact["opted_out"]:
            if re.search(r"\bstart\b", text, re.I):
                self._set_state(msg.phone, "new", opted_out=0)
                reply = BotReply("You're subscribed again. " + MENU)
            else:
                reply = BotReply("You've opted out of messages. Reply START to resubscribe.")
            self._log(msg.phone, "out", reply.text)
            return reply

        state = contact["state"]

        # Lead-capture states take priority over keyword matching.
        if state == "awaiting_name":
            self._set_state(msg.phone, "awaiting_interest", name=text.title())
            reply = BotReply(
                f"Thanks {text.title()}! What are you interested in — "
                "Pricing, a Demo, or something else?"
            )
        elif state == "awaiting_interest":
            self._set_state(msg.phone, "captured", interest=text)
            reply = BotReply(
                "Noted! Our team will reach out within one business day. "
                "Anything else? Reply MENU for options."
            )
        else:
            reply = self._route(text, contact)

        self._log(msg.phone, "out", reply.text)
        return reply

    def _route(self, text: str, contact: sqlite3.Row) -> BotReply:
        for pattern, intent in KEYWORD_MAP:
            if pattern.search(text):
                break
        else:
            intent = "fallback"

        if intent == "greeting":
            name = f" {contact['name']}" if contact["name"] else ""
            return BotReply(f"Hello{name}! Welcome to Acme Automation. " + MENU)
        if intent == "menu":
            return BotReply(MENU)
        if intent == "pricing":
            return BotReply(FAQ["pricing"], buttons=["Book a demo", "Talk to a human"])
        if intent == "hours":
            return BotReply(FAQ["hours"])
        if intent == "demo":
            self._set_state(contact["phone"], "awaiting_name")
            return BotReply(FAQ["demo"])
        if intent == "human":
            self._set_state(contact["phone"], "handoff")
            return BotReply(
                "Connecting you with a human agent — they'll reply here within "
                "one business day. Your reference is kept in this chat."
            )
        if intent == "stop":
            self._set_state(contact["phone"], "new", opted_out=1)
            return BotReply(
                "You've been unsubscribed and won't receive further messages. "
                "Reply START anytime to resubscribe."
            )
        if text.strip() == "1":
            return BotReply(FAQ["pricing"])
        if text.strip() == "2":
            self._set_state(contact["phone"], "awaiting_name")
            return BotReply(FAQ["demo"])
        if text.strip() == "3":
            return BotReply(FAQ["hours"])
        if text.strip() == "4":
            self._set_state(contact["phone"], "handoff")
            return BotReply("Connecting you with a human agent — they'll reply here within one business day.")
        return BotReply(
            "I didn't quite get that. " + MENU
        )


# ---------------------------------------------------------------------------
# Outbound payloads (WhatsApp Cloud API format)
# ---------------------------------------------------------------------------


def text_payload(to: str, body: str) -> dict:
    return {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "text",
        "text": {"preview_url": False, "body": body},
    }


def template_payload(to: str, template: str, language: str = "en_US",
                     variables: list[str] | None = None) -> dict:
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "template",
        "template": {"name": template, "language": {"code": language}},
    }
    if variables:
        payload["template"]["components"] = [
            {"type": "body",
             "parameters": [{"type": "text", "text": v} for v in variables]}
        ]
    return payload


def button_payload(to: str, body: str, buttons: list[str]) -> dict:
    return {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "interactive",
        "interactive": {
            "type": "button",
            "body": {"text": body},
            "action": {
                "buttons": [
                    {"type": "reply", "reply": {"id": f"btn_{i}", "title": b[:20]}}
                    for i, b in enumerate(buttons[:3])
                ]
            },
        },
    }


class ConsoleSender:
    """Mock sender: prints payloads instead of calling Meta. Used for demos."""

    def __init__(self):
        self.sent: list[dict] = []

    def send(self, payload: dict) -> dict:
        self.sent.append(payload)
        to = payload.get("to")
        kind = payload.get("type")
        print(f"[SEND -> {to}] ({kind})")
        return {"mock": True, "to": to}


class CloudApiSender:
    """Real sender: POSTs to the WhatsApp Cloud API."""

    def __init__(self, phone_number_id: str, access_token: str):
        self.url = (
            f"https://graph.facebook.com/v21.0/{phone_number_id}/messages"
        )
        self.access_token = access_token

    def send(self, payload: dict) -> dict:
        req = urllib.request.Request(
            self.url,
            data=json.dumps(payload).encode(),
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())


# ---------------------------------------------------------------------------
# Broadcaster
# ---------------------------------------------------------------------------

PHONE_RE = re.compile(r"^\+?[1-9]\d{7,14}$")


def normalise_phone(raw: str) -> str | None:
    digits = re.sub(r"\D", "", raw.strip())
    if not digits:
        return None
    phone = "+" + digits
    return phone if PHONE_RE.match(phone) else None


class Broadcaster:
    def __init__(self, conn: sqlite3.Connection, sender,
                 rate_per_second: float = 1.0):
        self.conn = conn
        self.sender = sender
        self.min_interval = 1.0 / rate_per_second if rate_per_second > 0 else 0
        self._last_send = 0.0

    def _throttle(self) -> None:
        wait = self._last_send + self.min_interval - time.monotonic()
        if wait > 0:
            time.sleep(wait)
        self._last_send = time.monotonic()

    def run(self, phones: list[str], template: str,
            variables: list[str] | None = None) -> dict:
        """Send a template campaign with dedupe + opt-out respect.

        Returns a summary dict; every attempt is recorded in broadcast_log.
        """
        seen: set[str] = set()
        summary = {"sent": 0, "skipped_dupe": 0, "skipped_optout": 0,
                   "skipped_invalid": 0}
        for raw in phones:
            phone = normalise_phone(raw)
            if not phone or phone in seen:
                summary["skipped_dupe" if phone else "skipped_invalid"] += 1
                seen.add(phone or raw)
                continue
            seen.add(phone)
            row = self.conn.execute(
                "SELECT opted_out FROM contacts WHERE phone = ?", (phone,)
            ).fetchone()
            if row and row["opted_out"]:
                self._log_attempt(phone, template, "skipped_optout")
                summary["skipped_optout"] += 1
                continue
            self._throttle()
            self.sender.send(template_payload(phone, template,
                                              variables=variables))
            self._log_attempt(phone, template, "sent")
            summary["sent"] += 1
        self.conn.commit()
        return summary

    def _log_attempt(self, phone: str, template: str, status: str) -> None:
        self.conn.execute(
            "INSERT INTO broadcast_log (phone, template, status, created_at)"
            " VALUES (?, ?, ?, ?)",
            (phone, template, status, time.time()),
        )


# ---------------------------------------------------------------------------
# Webhook HTTP server
# ---------------------------------------------------------------------------


class WebhookHandler(BaseHTTPRequestHandler):
    verify_token = "changeme"
    app_secret = "changeme"
    sender = ConsoleSender()
    engine: BotEngine | None = None

    def log_message(self, *args):  # quieter logs
        sys.stderr.write("[webhook] " + args[0] % args[1:] + "\n")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)
        status, body = verify_webhook(query, self.verify_token)
        self.send_response(status)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(body.encode())

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        sig = self.headers.get("X-Hub-Signature-256")
        if not check_signature(raw, self.app_secret, sig):
            self.send_response(403)
            self.end_headers()
            self.wfile.write(b"bad signature")
            return
        try:
            payload = json.loads(raw.decode())
        except json.JSONDecodeError:
            self.send_response(400)
            self.end_headers()
            return
        for msg in parse_incoming(payload):
            reply = self.engine.handle(msg)
            if reply.buttons:
                out = button_payload(msg.phone, reply.text, reply.buttons)
            else:
                out = text_payload(msg.phone, reply.text)
            self.sender.send(out)
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def cmd_simulate() -> None:
    conn = get_db()
    engine = BotEngine(conn)
    sender = ConsoleSender()
    print("WhatsFlow simulator — type messages as the customer (quit to exit).")
    phone = "+911234567890"
    while True:
        try:
            text = input("you> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if text.lower() in {"quit", "exit"}:
            break
        reply = engine.handle(InboundMessage(phone, "Demo User", text, "m1", "text"))
        sender.send(text_payload(phone, reply.text))
        print("bot>", reply.text.replace("\n", "\n     "))
    conn.close()


def cmd_server(port: int, verify_token: str, app_secret: str) -> None:
    conn = get_db()
    WebhookHandler.engine = BotEngine(conn)
    WebhookHandler.verify_token = verify_token
    WebhookHandler.app_secret = app_secret
    server = HTTPServer(("0.0.0.0", port), WebhookHandler)
    print(f"Webhook listening on :{port} (verify token set, signature check on).")
    print("Expose via ngrok/cloudflared and register the public URL with Meta.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    conn.close()


def cmd_broadcast(csv_path: str, template: str) -> None:
    phones = []
    with open(csv_path, newline="") as f:
        for row in csv.DictReader(f):
            phones.append(row.get("phone", ""))
    conn = get_db()
    summary = Broadcaster(conn, ConsoleSender()).run(phones, template)
    conn.close()
    print(json.dumps(summary, indent=2))


def cmd_export(csv_path: str) -> None:
    conn = get_db()
    rows = conn.execute(
        "SELECT phone, name, interest, state, opted_out FROM contacts ORDER BY phone"
    ).fetchall()
    with open(csv_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["phone", "name", "interest", "state", "opted_out"])
        w.writerows([tuple(r) for r in rows])
    conn.close()
    print(f"Exported {len(rows)} contacts to {csv_path}")


def cmd_report() -> None:
    conn = get_db()
    n_in = conn.execute(
        "SELECT COUNT(*) c FROM messages WHERE direction='in'").fetchone()["c"]
    n_out = conn.execute(
        "SELECT COUNT(*) c FROM messages WHERE direction='out'").fetchone()["c"]
    leads = conn.execute(
        "SELECT COUNT(*) c FROM contacts WHERE state='captured'").fetchone()["c"]
    handoffs = conn.execute(
        "SELECT COUNT(*) c FROM contacts WHERE state='handoff'").fetchone()["c"]
    print(f"inbound: {n_in}  outbound: {n_out}  leads captured: {leads}  "
          f"human handoffs: {handoffs}")
    conn.close()


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1
    cmd = argv[1]
    if cmd == "simulate":
        cmd_simulate()
    elif cmd == "server":
        port = int(argv[argv.index("--port") + 1]) if "--port" in argv else 8000
        vt = argv[argv.index("--verify-token") + 1] if "--verify-token" in argv else "changeme"
        secret = argv[argv.index("--app-secret") + 1] if "--app-secret" in argv else "changeme"
        cmd_server(port, vt, secret)
    elif cmd == "broadcast":
        csv_path = argv[argv.index("--csv") + 1]
        template = argv[argv.index("--template") + 1] if "--template" in argv else "hello_world"
        cmd_broadcast(csv_path, template)
    elif cmd == "export":
        cmd_export(argv[argv.index("--out") + 1] if "--out" in argv else "leads.csv")
    elif cmd == "report":
        cmd_report()
    else:
        print(f"unknown command: {cmd}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
