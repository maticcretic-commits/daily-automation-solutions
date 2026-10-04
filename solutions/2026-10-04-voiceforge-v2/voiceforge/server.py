"""VoiceForge v2 HTTP server — stdlib-only webhook bindings.

This is the missing last mile: the engine produces TwiML and parses
webhooks, but nothing listens on a port. This server binds every webhook
route the adapters define, and it is what Twilio points its webhooks at.

Run (pilot):
    export PUBLIC_BASE_URL="https://<your-host>"   # required
    export TWILIO_AUTH_TOKEN="<token>"             # required
    export TWILIO_ACCOUNT_SID="<sid>"              # media download + SMS
    export TWILIO_PHONE_NUMBER="+15551234567"      # enables real SMS
    export VOICEFORGE_BUSINESS="Acme Services"
    export VOICEFORGE_TRANSFER_NUMBER="+15557654321"
    python3 server.py            # or: python3 -m voiceforge.server
    # Twilio console -> Phone Numbers -> Voice webhook:  {PUBLIC_BASE_URL}/voice
    # Twilio console -> Phone Numbers -> Messaging webhook: {PUBLIC_BASE_URL}/messaging

Optional:
    LLM_API_URL / LLM_API_KEY / LLM_MODEL   live LLM (OpenAI-compatible)
    VOICEFORGE_REQUIRE_LIVE_LLM=1           fail fast if the LLM is a mock
    VAPI_WEBHOOK_SECRET                     enables POST /vapi (Twilio-only pilot otherwise)
    VOICEFORGE_AI_DISCLOSURE=0              disable AI identity disclosure (not recommended)
    VOICEFORGE_RECORD_CALLS=0               disable <Record> (pilot default: ON)
    VOICEFORGE_MEDIA_DIR                    MMS photo storage (default ./media)
    PORT                                    listen port (default 8080)

Stdlib only — no dependencies.
"""

from __future__ import annotations

import json
import os
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Dict, Optional

from .adapters_twilio import TwilioAdapter
from .adapters_vapi import VapiAdapter
from .agent import VoiceAgent
from .brain import ConversationBrain
from .config import VoiceForgeConfig
from .logging import CallLogger
from .providers import (HttpLLMClient, HttpVisionClient, MockLLMClient,
                        MockVisionClient)
from .vision import ContextInjector, MmsWatch, VisionAnalyzer, VisionFlow


# ---------------------------------------------------------------------------
# wiring: env -> live objects, with fail-fast validation
# ---------------------------------------------------------------------------

def build_llm(config: VoiceForgeConfig):
    """Live LLM when configured; mock only when explicitly allowed."""
    api_url = os.environ.get("LLM_API_URL", "").strip()
    if api_url:
        api_key = os.environ.get("LLM_API_KEY", "").strip()
        if not api_key:
            sys.exit("FATAL: LLM_API_URL is set but LLM_API_KEY is missing.")
        model = os.environ.get("LLM_MODEL", "gpt-4o-mini").strip()
        return HttpLLMClient(api_url=api_url, api_key=api_key,
                             model=model, config=config)
    if os.environ.get("VOICEFORGE_REQUIRE_LIVE_LLM", "0") == "1":
        sys.exit("FATAL: VOICEFORGE_REQUIRE_LIVE_LLM=1 but no LLM_API_URL — "
                 "refusing to serve a parroting mock as a live agent.")
    print("WARNING: no LLM_API_URL — serving with MockLLMClient. "
          "Set LLM_API_URL/LLM_API_KEY (or VOICEFORGE_REQUIRE_LIVE_LLM=1 "
          "to fail fast) for a real pilot.", file=sys.stderr, flush=True)
    return MockLLMClient()


class VoiceForgeApp:
    """All live objects, built once at startup."""

    def __init__(self):
        business = os.environ.get("VOICEFORGE_BUSINESS", "Acme Services")
        self.config = VoiceForgeConfig.from_env(business=business)
        self.base_url = self.config.public_base_url
        if not self.base_url:
            sys.exit("FATAL: PUBLIC_BASE_URL is not set — Twilio rejects "
                     "relative callback URLs. Export PUBLIC_BASE_URL, e.g. "
                     "https://<your-host> (ngrok works for the pilot).")
        self.logger = CallLogger()
        self.twilio = TwilioAdapter(
            # auth_token/account_sid fall back to env; missing token raises
            gather_action_url=f"{self.base_url}/gather",
            config=self.config)
        brain = ConversationBrain(business=business, config=self.config,
                                  llm=build_llm(self.config))
        # SMS promise honesty: the brain promises SMS only when the adapter
        # can actually send one.
        brain.sms_enabled = self.twilio.sms_configured
        if not self.twilio.sms_configured:
            print("NOTE: SMS not configured (need TWILIO_ACCOUNT_SID + "
                  "TWILIO_PHONE_NUMBER) — booking confirmations will give "
                  "a reference number instead.", file=sys.stderr, flush=True)
        self.vision_flow = VisionFlow(
            mms_watch=MmsWatch(
                self.twilio, self.logger, config=self.config,
                media_dir=os.environ.get("VOICEFORGE_MEDIA_DIR", "media")),
            analyzer=VisionAnalyzer(
                self._vision_client(), self.logger, config=self.config),
            injector=ContextInjector(self.logger),
            logger=self.logger, config=self.config)
        self.agent = VoiceAgent(brain, logger=self.logger,
                                config=self.config,
                                vision_flow=self.vision_flow)
        vapi_secret = os.environ.get("VAPI_WEBHOOK_SECRET", "")
        self.vapi = VapiAdapter(webhook_secret=vapi_secret) \
            if vapi_secret else None
        print(f"VoiceForge v2 serving {business} at {self.base_url} "
              f"(sms={'on' if self.twilio.sms_configured else 'off'}, "
              f"record={'on' if self.config.record_calls else 'off'})",
              flush=True)

    def _vision_client(self):
        api_url = os.environ.get("LLM_API_URL", "").strip()
        api_key = os.environ.get("LLM_API_KEY", "").strip()
        if api_url and api_key:
            return HttpVisionClient(api_url=api_url, api_key=api_key,
                                    model=os.environ.get("LLM_MODEL",
                                                         "gpt-4o-mini"))
        return MockVisionClient()


APP: Optional[VoiceForgeApp] = None


# ---------------------------------------------------------------------------
# request handling
# ---------------------------------------------------------------------------

def _read_form(handler: BaseHTTPRequestHandler) -> Dict[str, str]:
    length = int(handler.headers.get("Content-Length", "0") or "0")
    raw = handler.rfile.read(length).decode("utf-8", "replace")
    return {k: v[0] for k, v in
            urllib.parse.parse_qs(raw, keep_blank_values=True).items()}


def _twilio_url(handler: BaseHTTPRequestHandler, base: str) -> str:
    """Reconstruct the exact public URL Twilio called (for signature)."""
    path = handler.path.split("?", 1)
    url = base + path[0]
    if len(path) == 2 and path[1]:
        url += "?" + path[1]
    return url


class Handler(BaseHTTPRequestHandler):
    server_version = "VoiceForge/2.0"

    # -- helpers ------------------------------------------------------------
    def _send(self, code: int, body: str,
              ctype: str = "text/xml; charset=utf-8"):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _twiml(self, body: str):
        self._send(200, body)

    def _forbidden(self, why: str):
        APP.logger.log("webhook_rejected", reason=why, path=self.path)
        self._send(403, "Forbidden", "text/plain; charset=utf-8")

    # -- GET ------------------------------------------------------------------
    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path == "/healthz":
            self._send(200, json.dumps({"ok": True,
                                        "service": "voiceforge-v2"}),
                       "application/json")
        else:
            self._send(404, "Not found", "text/plain; charset=utf-8")

    # -- POST -----------------------------------------------------------------
    def do_POST(self):
        path = self.path.split("?", 1)[0]
        query = urllib.parse.parse_qs(
            self.path.split("?", 1)[1] if "?" in self.path else "")
        routes = {
            "/voice": self._voice,
            "/gather": self._voice,
            "/whisper": lambda: self._whisper(query),
            "/transfer_fallback": self._transfer_fallback,
            "/messaging": self._messaging,
            "/vapi": self._vapi,
        }
        handler = routes.get(path)
        if handler is None:
            self._send(404, "Not found", "text/plain; charset=utf-8")
            return
        try:
            handler()
        except (BrokenPipeError, ConnectionResetError):
            pass  # caller hung up mid-write; nothing to do

    # -- Twilio voice -----------------------------------------------------------
    def _voice(self):
        form = _read_form(self)
        url = _twilio_url(self, APP.base_url)
        if not APP.twilio.verify_signature(
                url, form, self.headers.get("X-Twilio-Signature", "")):
            self._forbidden("bad_twilio_signature")
            return
        call_sid = form.get("CallSid", "")
        from_phone = form.get("From", "")
        if not call_sid:
            self._forbidden("missing_call_sid")
            return
        if APP.agent.has_call(call_sid):
            reply = APP.agent.caller_said(
                call_sid, form.get("SpeechResult", ""), twilio=APP.twilio)
        else:
            reply = APP.agent.inbound_call(call_sid, from_phone,
                                           twilio=APP.twilio)
        self._twiml(reply)

    def _whisper(self, query):
        sid = (query.get("sid") or [""])[0]
        session = dict(APP.agent.calls_snapshot()).get(sid)
        context = APP.agent.build_handoff_context(session) if session \
            else "Incoming transfer."
        self._twiml(APP.twilio.whisper_twiml(context))

    def _transfer_fallback(self):
        self._twiml(APP.twilio.transfer_fallback_twiml())

    # -- Twilio messaging (MMS side channel) --------------------------------------
    def _messaging(self):
        form = _read_form(self)
        url = _twilio_url(self, APP.base_url)
        sig = self.headers.get("X-Twilio-Signature", "")
        if APP.vision_flow.on_mms_webhook(APP.agent, url, form, sig):
            self._twiml('<?xml version="1.0" encoding="UTF-8"?>'
                        '<Response/>')
        else:
            self._forbidden("bad_twilio_signature")

    # -- Vapi (Twilio-only pilot: verified + logged, not live-wired) ---------------
    def _vapi(self):
        if APP.vapi is None:
            self._send(503, json.dumps({"error": "vapi_not_configured"}),
                       "application/json")
            return
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8", "replace")
        try:
            body = json.loads(raw or "{}")
        except json.JSONDecodeError:
            body = {}
        if not APP.vapi.verify(
                {"x-vapi-secret": self.headers.get("X-Vapi-Secret", "")}):
            self._forbidden("bad_vapi_secret")
            return
        action = VapiAdapter.parse_message(body)
        APP.logger.log("vapi_message", action=action.get("action"))
        # First pilot is Twilio-only; Vapi events are verified and logged.
        self._send(200, json.dumps({}), "application/json")

    def log_message(self, fmt, *args):  # quieter stdlib logging
        sys.stderr.write(f"[http] {self.address_string()} {fmt % args}\n")


def main() -> None:
    global APP
    APP = VoiceForgeApp()
    port = int(os.environ.get("PORT", "8080"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Listening on 0.0.0.0:{port} — point Twilio webhooks at "
          f"{APP.base_url}/voice and {APP.base_url}/messaging", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
