"""Tests for the WhatsApp Cloud API webhook template.

Env vars are set BEFORE importing the app module, because the app fails
fast at import time when VERIFY_TOKEN / APP_SECRET are missing -- which is
exactly the behaviour we want in production.
"""

import hashlib
import hmac
import json
import os

os.environ.setdefault("VERIFY_TOKEN", "test-verify-token")
os.environ.setdefault("APP_SECRET", "test-app-secret")

import pytest  # noqa: E402

import webhook_server  # noqa: E402
from webhook_server import create_app  # noqa: E402


@pytest.fixture()
def client():
    app = create_app()
    app.config["TESTING"] = True
    return app.test_client()


def _sign(secret: str, body: bytes) -> str:
    return "sha256=" + hmac.new(secret.encode(), body, hashlib.sha256).hexdigest()


# --- GET /webhook: verification handshake -----------------------------------

def test_handshake_accept_correct_token(client):
    resp = client.get(
        "/webhook",
        query_string={
            "hub.mode": "subscribe",
            "hub.verify_token": "test-verify-token",
            "hub.challenge": "CHALLENGE_123",
        },
    )
    assert resp.status_code == 200
    assert resp.get_data(as_text=True) == "CHALLENGE_123"


def test_handshake_reject_wrong_token(client):
    resp = client.get(
        "/webhook",
        query_string={
            "hub.mode": "subscribe",
            "hub.verify_token": "wrong-token",
            "hub.challenge": "CHALLENGE_123",
        },
    )
    assert resp.status_code == 403


def test_handshake_reject_wrong_mode(client):
    resp = client.get(
        "/webhook",
        query_string={
            "hub.mode": "unsubscribe",
            "hub.verify_token": "test-verify-token",
            "hub.challenge": "CHALLENGE_123",
        },
    )
    assert resp.status_code == 403


def test_handshake_reject_missing_params(client):
    resp = client.get("/webhook")
    assert resp.status_code == 403


# --- POST /webhook: signature validation -------------------------------------

SAMPLE_EVENT = {
    "object": "whatsapp_business_account",
    "entry": [
        {
            "id": "12345",
            "changes": [
                {
                    "field": "messages",
                    "value": {
                        "messages": [
                            {
                                "from": "919876543210",
                                "id": "wamid.abc123",
                                "timestamp": "1791300000",
                                "type": "text",
                                "text": {"body": "Need 3 forklifts in Sohar"},
                            }
                        ]
                    },
                }
            ],
        }
    ],
}


def test_post_valid_signature_accepted(client):
    body = json.dumps(SAMPLE_EVENT).encode()
    resp = client.post(
        "/webhook",
        data=body,
        content_type="application/json",
        headers={"X-Hub-Signature-256": _sign("test-app-secret", body)},
    )
    assert resp.status_code == 200
    assert resp.get_data(as_text=True) == "EVENT_RECEIVED"


def test_post_invalid_signature_rejected(client):
    body = json.dumps(SAMPLE_EVENT).encode()
    resp = client.post(
        "/webhook",
        data=body,
        content_type="application/json",
        headers={"X-Hub-Signature-256": _sign("wrong-secret", body)},
    )
    assert resp.status_code == 403


def test_post_tampered_body_rejected(client):
    body = json.dumps(SAMPLE_EVENT).encode()
    signature = _sign("test-app-secret", body)
    tampered = body.replace(b"forklifts", b"cranes")  # signed, then altered
    resp = client.post(
        "/webhook",
        data=tampered,
        content_type="application/json",
        headers={"X-Hub-Signature-256": signature},
    )
    assert resp.status_code == 403


def test_post_missing_signature_rejected(client):
    body = json.dumps(SAMPLE_EVENT).encode()
    resp = client.post("/webhook", data=body, content_type="application/json")
    assert resp.status_code == 403


def test_post_malformed_json_rejected(client):
    body = b'{"entry": [not valid json'
    resp = client.post(
        "/webhook",
        data=body,
        content_type="application/json",
        headers={"X-Hub-Signature-256": _sign("test-app-secret", body)},
    )
    assert resp.status_code == 400


# --- structured logging -------------------------------------------------------

def test_message_event_logged_without_secrets(client, caplog):
    body = json.dumps(SAMPLE_EVENT).encode()
    with caplog.at_level("INFO", logger="whatsapp-webhook"):
        resp = client.post(
            "/webhook",
            data=body,
            content_type="application/json",
            headers={"X-Hub-Signature-256": _sign("test-app-secret", body)},
        )
    assert resp.status_code == 200
    logged = "\n".join(record.getMessage() for record in caplog.records)
    assert '"from": "919876543210"' in logged
    assert '"type": "text"' in logged
    # secrets must never appear in logs
    assert "test-app-secret" not in logged
    assert "test-verify-token" not in logged


# --- fail-fast config ----------------------------------------------------------

def test_missing_env_vars_fail_fast(monkeypatch):
    monkeypatch.delenv("VERIFY_TOKEN", raising=False)
    monkeypatch.delenv("APP_SECRET", raising=False)
    with pytest.raises(RuntimeError):
        webhook_server.create_app()
