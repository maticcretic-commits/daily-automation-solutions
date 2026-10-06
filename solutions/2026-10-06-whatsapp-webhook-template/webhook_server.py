"""WhatsApp Business Cloud API webhook template.

Reusable foundation for WhatsApp automation builds (RFQ bots, lead capture,
support flows). Demo/template code, not client work.

Endpoints:
    GET  /webhook  -- Meta verification handshake
    POST /webhook  -- incoming message/status events (signature-verified)

Configuration is env-vars ONLY -- never hard-code tokens:
    VERIFY_TOKEN  token you invent when configuring the webhook in Meta's dashboard
    APP_SECRET    your Meta app's App Secret (used for X-Hub-Signature-256 validation)
    PORT          listen port (default 5000)

Run:
    VERIFY_TOKEN=... APP_SECRET=... python webhook_server.py
"""

import hashlib
import hmac
import json
import logging
import os
import sys

from flask import Flask, Response, request

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",  # we emit structured JSON lines ourselves
)
log = logging.getLogger("whatsapp-webhook")


def _required_env(name):
    value = os.environ.get(name, "")
    if not value:
        raise RuntimeError(
            f"{name} is not set. Configure it via environment variable -- "
            "never hard-code tokens."
        )
    return value


def create_app():
    verify_token = _required_env("VERIFY_TOKEN")
    app_secret = _required_env("APP_SECRET").encode("utf-8")

    app = Flask(__name__)

    def _log_event(payload):
        """Emit one structured JSON log line per message/status. Never logs
        secrets or raw payloads -- only the fields needed for debugging."""
        for entry in payload.get("entry", []):
            for change in entry.get("changes", []):
                value = change.get("value", {})
                for msg in value.get("messages", []):
                    log.info(
                        json.dumps(
                            {
                                "event": "message",
                                "from": msg.get("from"),
                                "timestamp": msg.get("timestamp"),
                                "type": msg.get("type"),
                                "message_id": msg.get("id"),
                            }
                        )
                    )
                for status in value.get("statuses", []):
                    log.info(
                        json.dumps(
                            {
                                "event": "status",
                                "recipient_id": status.get("recipient_id"),
                                "status": status.get("status"),
                                "timestamp": status.get("timestamp"),
                                "message_id": status.get("id"),
                            }
                        )
                    )

    @app.get("/webhook")
    def verify():
        """Meta's verification handshake: proves we own this callback URL."""
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge", "")
        if mode == "subscribe" and token == verify_token:
            return Response(challenge, status=200, mimetype="text/plain")
        return Response("Forbidden", status=403)

    @app.post("/webhook")
    def receive():
        """Receive events. Rejects anything without a valid HMAC-SHA256
        signature (X-Hub-Signature-256) -- tampered or forged posts get 403."""
        body = request.get_data()
        signature = request.headers.get("X-Hub-Signature-256", "")
        expected = "sha256=" + hmac.new(app_secret, body, hashlib.sha256).hexdigest()
        if not signature or not hmac.compare_digest(signature, expected):
            return Response("Forbidden", status=403)
        try:
            payload = json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return Response("Bad Request", status=400)
        _log_event(payload)
        return Response("EVENT_RECEIVED", status=200)

    return app


app = create_app()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
