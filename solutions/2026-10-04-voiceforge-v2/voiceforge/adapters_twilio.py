"""Twilio telephony adapter: signature verification + TwiML builders.

The XML escape invariant: replies are interpolated into ELEMENT TEXT only
(<Say>), never into attributes — so escaping & < > is sufficient.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
from typing import Dict

from .config import VoiceForgeConfig


class TwilioAdapter:
    """
    Handles Twilio voice webhooks:
    - verifies X-Twilio-Signature (HMAC-SHA1 over URL + sorted params)
    - returns TwiML that says the reply and gathers the next utterance
    - builds warm-transfer <Dial> with a whisper leg for context handoff

    NOTE on URL reconstruction: Twilio signs the exact public URL it called,
    including non-standard ports. Behind a proxy, build `url` from
    X-Forwarded-Proto/Host exactly as Twilio sees it, or verification will
    silently fail — log the reconstructed payload when debugging.
    """

    def __init__(self, auth_token: str = "",
                 gather_action_url: str = "/gather",
                 config: "VoiceForgeConfig | None" = None,
                 account_sid: str = ""):
        if not auth_token:
            auth_token = os.environ.get("TWILIO_AUTH_TOKEN", "")
        if not auth_token:
            raise ValueError(
                "TwilioAdapter requires an auth token (arg or "
                "TWILIO_AUTH_TOKEN); refusing to run unverified.")
        self.auth_token = auth_token
        # B1: Twilio media URLs require HTTP Basic Auth (Account SID as the
        # username). Optional here — webhook verification needs only the
        # token — but media download fails closed without it.
        if not account_sid:
            account_sid = os.environ.get("TWILIO_ACCOUNT_SID", "")
        self.account_sid = account_sid
        self.gather_action_url = gather_action_url
        self.config = config

    def verify_signature(self, url: str, params: Dict[str, str],
                         signature: str) -> bool:
        payload = url + "".join(k + params[k] for k in sorted(params))
        digest = hmac.new(self.auth_token.encode(), payload.encode(),
                          hashlib.sha1).digest()
        expected = base64.b64encode(digest).decode()
        return hmac.compare_digest(expected, signature)

    @staticmethod
    def _xml_escape(text: str) -> str:
        return (text.replace("&", "&amp;").replace("<", "&lt;")
                    .replace(">", "&gt;"))

    def inbound_twiml(self, greeting: str) -> str:
        return (f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Response><Say voice="alice">'
                f'{self._xml_escape(greeting)}'
                f'</Say><Gather input="speech" action="{self.gather_action_url}" '
                f'method="POST" speechTimeout="auto"/></Response>')

    def continue_twiml(self, reply: str, end_call: bool = False) -> str:
        say = (f'<Say voice="alice">{self._xml_escape(reply)}</Say>')
        if end_call:
            return (f'<?xml version="1.0" encoding="UTF-8"?>'
                    f'<Response>{say}<Hangup/></Response>')
        return (f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Response>{say}<Gather input="speech" '
                f'action="{self.gather_action_url}" method="POST" '
                f'speechTimeout="auto"/></Response>')

    def transfer_twiml(self, announcement: str, target_number: str,
                       whisper_url: str, fallback_action_url: str,
                       timeout_s: int = 20) -> str:
        """
        Warm transfer: say the announcement, then <Dial> the human with a
        whisper leg (whisper_url serves context to the agent before bridging).
        If the human does not answer, Twilio POSTs to fallback_action_url.
        """
        esc = self._xml_escape
        return (
            f'<?xml version="1.0" encoding="UTF-8"?>'
            f'<Response>'
            f'<Say voice="alice">{esc(announcement)}</Say>'
            f'<Dial timeout="{timeout_s}" action="{fallback_action_url}" '
            f'method="POST">'
            f'<Number url="{whisper_url}">{esc(target_number)}</Number>'
            f'</Dial>'
            f'</Response>')

    def whisper_twiml(self, context_summary: str) -> str:
        """TwiML served at whisper_url: briefs the human agent pre-bridge."""
        return (
            f'<?xml version="1.0" encoding="UTF-8"?>'
            f'<Response><Say voice="alice">'
            f'{self._xml_escape(context_summary)}'
            f'</Say></Response>')

    def transfer_fallback_twiml(self) -> str:
        """When the transfer target does not answer: take a message instead."""
        return self.continue_twiml(
            "I'm sorry, all of our teammates are busy right now. "
            "Please tell me your name and number after the tone, and "
            "we'll call you back within the hour.")
