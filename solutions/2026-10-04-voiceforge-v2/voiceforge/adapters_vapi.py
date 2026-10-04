"""Vapi server-message adapter with fail-closed secret handling.

P0-1(1): an empty/misconfigured webhook secret must NEVER silently accept
requests — construction raises, so a misconfigured deploy fails loudly at
startup instead of running an open webhook.
"""

from __future__ import annotations

import hmac
import os
from typing import Any, Dict


class VapiAdapter:
    """
    Handles Vapi server messages (webhook events). Verifies the
    X-Vapi-Secret header and turns assistant-request messages into
    conversation actions understood by the brain.
    """

    def __init__(self, webhook_secret: str = ""):
        if not webhook_secret:
            webhook_secret = os.environ.get("VAPI_WEBHOOK_SECRET", "")
        if not webhook_secret:
            raise ValueError(
                "VapiAdapter requires a webhook secret (arg or "
                "VAPI_WEBHOOK_SECRET); refusing to run unverified.")
        self.webhook_secret = webhook_secret

    def verify(self, headers: Dict[str, str]) -> bool:
        got = headers.get("x-vapi-secret", "")
        if not got:
            return False
        return hmac.compare_digest(got, self.webhook_secret)

    @staticmethod
    def parse_message(body: Dict[str, Any]) -> Dict[str, Any]:
        """Normalise a Vapi server message into a simple action dict.

        Defensive throughout: any malformed shape yields {"action":
        "ignore"} instead of raising.
        """
        if not isinstance(body, dict):
            return {"action": "ignore"}
        message = body.get("message")
        if not isinstance(message, dict):
            return {"action": "ignore"}
        call = body.get("call")
        call_id = call.get("id", "unknown") if isinstance(call, dict) \
            else "unknown"
        msg_type = message.get("type", "")
        if msg_type == "assistant-request":
            return {"action": "start_call", "call_id": call_id}
        if msg_type == "transcript" and message.get("role") == "user":
            text = message.get("transcript")
            if not isinstance(text, str):
                return {"action": "ignore", "type": msg_type}
            return {"action": "caller_said", "call_id": call_id,
                    "text": text}
        if msg_type == "end-of-call-report":
            return {"action": "call_ended", "call_id": call_id,
                    "summary": body.get("summary", "") or ""}
        return {"action": "ignore", "type": msg_type}
