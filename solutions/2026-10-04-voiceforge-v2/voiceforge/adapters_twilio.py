"""Twilio telephony adapter: signature verification + TwiML builders.

The XML escape invariant: replies are interpolated into ELEMENT TEXT only
(<Say>), never into attributes — so escaping & < > is sufficient.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
from typing import Dict, Optional, Tuple

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
                 account_sid: str = "",
                 sms_from: str = ""):
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
        # SMS sender number (a Twilio number on the account). SMS is only
        # promised/offered when both the SID and a sender number exist —
        # otherwise the booking flow gives a reference number instead.
        if not sms_from:
            sms_from = os.environ.get("TWILIO_PHONE_NUMBER", "")
        self.sms_from = sms_from
        self.gather_action_url = gather_action_url
        self.config = config

    @property
    def sms_configured(self) -> bool:
        """True only when real SMS credentials are present."""
        return bool(self.account_sid and self.sms_from)

    def send_sms(self, to: str, body: str):
        """
        Send a real SMS via the Twilio Messages API (stdlib only).
        Returns (True, message_sid) or (False, error_text). Never raises.
        """
        if not self.sms_configured:
            return False, "sms_not_configured"
        import urllib.parse
        import urllib.request
        url = (f"https://api.twilio.com/2010-04-01/Accounts/"
               f"{self.account_sid}/Messages.json")
        data = urllib.parse.urlencode(
            {"To": to, "From": self.sms_from, "Body": body}).encode()
        req = urllib.request.Request(url, data=data, method="POST")
        creds = base64.b64encode(
            f"{self.account_sid}:{self.auth_token}".encode()).decode()
        req.add_header("Authorization", f"Basic {creds}")
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                payload = json.loads(resp.read().decode())
        except Exception as exc:  # noqa: BLE001 - degrade, never crash
            return False, str(exc)
        sid = payload.get("sid", "")
        if not sid:
            return False, payload.get("message", "unknown_error")
        return True, sid

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

    def inbound_twiml(self, greeting: str, record: Optional[bool] = None) -> str:
        # The greeting discloses that the call may be recorded — so record
        # it (default ON for pilot, VOICEFORGE_RECORD_CALLS=0 to disable).
        # `record` overrides the config flag when given (tests).
        cfg_record = self.config.record_calls if self.config else True
        do_record = cfg_record if record is None else record
        rec = "<Record/>" if do_record else ""
        return (f'<?xml version="1.0" encoding="UTF-8"?>'
                f'<Response><Say voice="alice">'
                f'{self._xml_escape(greeting)}'
                f'</Say>{rec}<Gather input="speech" action="{self.gather_action_url}" '
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
