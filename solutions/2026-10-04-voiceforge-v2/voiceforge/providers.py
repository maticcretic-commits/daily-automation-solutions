"""Provider interfaces (LLM / TTS / vision) with mocks and stdlib HTTP clients.

P0-1(2): every real network call goes through retry-once with backoff and
then a safe canned fallback. Response schemas are validated before any
indexing. No provider method used on the call loop raises for transient
failures.
"""

from __future__ import annotations

import json
import time
import urllib.error
import urllib.request
from typing import Any, Callable, Dict, List, Optional, Tuple

from .config import VoiceForgeConfig
from .session import CallSession


# ---------------------------------------------------------------------------
# retry helper
# ---------------------------------------------------------------------------

def _transient(exc: Exception) -> bool:
    """True for failures worth one retry: timeouts, 429, 5xx, DNS."""
    if isinstance(exc, TimeoutError):
        return True
    if isinstance(exc, urllib.error.HTTPError):
        return exc.code == 429 or 500 <= exc.code < 600
    if isinstance(exc, urllib.error.URLError):
        return True
    return False


def call_with_retry(fn: Callable[[], Any], retries: int,
                    backoff_s: float) -> Tuple[bool, Any]:
    """
    Run fn(); retry up to `retries` times on transient failures with
    linear backoff. Returns (ok, result-or-exception).
    """
    attempt = 0
    while True:
        try:
            return True, fn()
        except Exception as exc:  # noqa: BLE001 - intentional boundary
            if _transient(exc) and attempt < retries:
                attempt += 1
                time.sleep(backoff_s * attempt)
                continue
            return False, exc


# ---------------------------------------------------------------------------
# LLM
# ---------------------------------------------------------------------------

class LLMClient:
    def reply(self, session: CallSession, transcript_text: str) -> str:
        raise NotImplementedError


class MockLLMClient(LLMClient):
    """Deterministic offline stand-in for demos and tests."""

    def __init__(self, script: Optional[Dict[str, str]] = None):
        self.script = script or {}

    def reply(self, session: CallSession, transcript_text: str) -> str:
        return self.script.get(
            "fallback", "Thanks, I've noted that. Is there anything else?")


class HttpLLMClient(LLMClient):
    """OpenAI-compatible chat-completions endpoint, stdlib-only."""

    SAFE_FALLBACK = ("I'm having a little trouble on my end right now — "
                     "could you say that once more?")

    def __init__(self, api_url: str, api_key: str, model: str = "gpt-4o-mini",
                 system_prompt: str = "You are a concise phone assistant.",
                 config: Optional[VoiceForgeConfig] = None):
        self.api_url = api_url
        self.api_key = api_key
        self.model = model
        self.system_prompt = system_prompt
        self.config = config or VoiceForgeConfig()

    def _request(self, session: CallSession) -> str:
        messages = [{"role": "system", "content": self.system_prompt}]
        for t in session.turns[-self.config.turn_window:]:
            if t.role in ("caller", "agent"):
                messages.append({
                    "role": "user" if t.role == "caller" else "assistant",
                    "content": t.text})
            elif t.role == "system" and t.text.startswith("[vision:"):
                # pin the latest vision context regardless of window
                messages.append({"role": "system", "content": t.text})
        body = json.dumps({
            "model": self.model,
            "messages": messages,
            "max_tokens": self.config.llm_max_tokens}).encode()
        req = urllib.request.Request(
            self.api_url, data=body,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {self.api_key}"})
        with urllib.request.urlopen(req,
                                    timeout=self.config.llm_timeout_s) as resp:
            data = json.loads(resp.read().decode())
        # validate schema before indexing (P0-1(2))
        choices = data.get("choices")
        if not isinstance(choices, list) or not choices:
            raise ValueError("LLM response has no choices")
        content = choices[0].get("message", {}).get("content")
        if not isinstance(content, str) or not content.strip():
            raise ValueError("LLM response has no message content")
        return content.strip()

    def reply(self, session: CallSession, transcript_text: str) -> str:
        ok, result = call_with_retry(
            lambda: self._request(session),
            retries=self.config.llm_retries,
            backoff_s=self.config.llm_retry_backoff_s)
        if ok:
            return result
        return self.SAFE_FALLBACK


# ---------------------------------------------------------------------------
# TTS
# ---------------------------------------------------------------------------

class TTSClient:
    def speak(self, text: str, voice: str = "default") -> bytes:
        raise NotImplementedError


class MockTTSClient(TTSClient):
    def speak(self, text: str, voice: str = "default") -> bytes:
        return f"[AUDIO voice={voice} chars={len(text)}] {text}".encode()


# ---------------------------------------------------------------------------
# Vision model client
# ---------------------------------------------------------------------------

class VisionClient:
    """Return a raw parsed JSON dict for (image, prompt); may raise."""

    def describe(self, image_bytes: bytes, content_type: str,
                 prompt: str, timeout_s: float) -> Dict[str, Any]:
        raise NotImplementedError


class MockVisionClient(VisionClient):
    """Canned vision responses for tests and offline demos."""

    def __init__(self, response: Optional[Dict[str, Any]] = None,
                 latency_s: float = 0.0,
                 exc: Optional[Exception] = None):
        self.response = response or {
            "description": "A cracked phone screen with a spiderweb fracture.",
            "findings": ["Screen glass is cracked across the upper half.",
                         "No visible damage to the camera module."],
            "confidence": 0.85}
        self.latency_s = latency_s
        self.exc = exc
        self.calls = 0

    def describe(self, image_bytes: bytes, content_type: str,
                 prompt: str, timeout_s: float) -> Dict[str, Any]:
        self.calls += 1
        if self.latency_s:
            time.sleep(self.latency_s)
        if self.exc is not None:
            raise self.exc
        return dict(self.response)


class HttpVisionClient(VisionClient):
    """OpenAI-compatible vision endpoint with forced JSON output."""

    def __init__(self, api_url: str, api_key: str,
                 model: str = "gpt-4o-mini"):
        self.api_url = api_url
        self.api_key = api_key
        self.model = model

    def describe(self, image_bytes: bytes, content_type: str,
                 prompt: str, timeout_s: float) -> Dict[str, Any]:
        import base64
        b64 = base64.b64encode(image_bytes).decode()
        body = json.dumps({
            "model": self.model,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url",
                     "image_url": {
                         "url": f"data:{content_type};base64,{b64}"}}]}],
            "response_format": {"type": "json_object"},
            "max_tokens": 400}).encode()
        req = urllib.request.Request(
            self.api_url, data=body,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {self.api_key}"})
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            data = json.loads(resp.read().decode())
        choices = data.get("choices")
        if not isinstance(choices, list) or not choices:
            raise ValueError("vision response has no choices")
        content = choices[0].get("message", {}).get("content")
        if not isinstance(content, str):
            raise ValueError("vision response has no content")
        parsed = json.loads(content)
        if not isinstance(parsed, dict):
            raise ValueError("vision response is not a JSON object")
        return parsed
