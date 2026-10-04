"""vision_analyze — vision model call with the 0.6 confidence gate.

Safety invariant: confidence < 0.6 -> the agent may ASK, never ASSERT.
Low-confidence output is converted into a clarification question; findings
never reach the spoken path. `analyze` NEVER raises to the call loop — every
failure mode returns a VisionResult with ok=False.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import List, Optional

from ..config import VoiceForgeConfig
from ..logging import CallLogger
from ..providers import VisionClient, call_with_retry
from ..session import Turn

SYSTEM_PROMPT = (
    "You are a careful visual inspector helping a phone support agent. "
    "Describe only what you can plainly see in the photo. Do not guess "
    "beyond the visible evidence. Reply with JSON only, exactly matching "
    "this schema: {\"description\": string (1-2 sentences), "
    "\"findings\": [string] (atomic claims, each independently true), "
    "\"confidence\": number between 0 and 1 (your certainty)}. "
    "If the image is too blurry, dark, or irrelevant to answer, set "
    "confidence below 0.4 and say so in the description.")


@dataclass
class VisionContext:
    business: str
    recent_turns: List[Turn]           # last ~6, caller+agent
    photo_request: str                # what the agent asked the caller to photo
    locale_hints: Optional[str] = None


@dataclass
class VisionResult:
    ok: bool
    description: str = ""
    findings: List[str] = field(default_factory=list)
    confidence: float = 0.0           # 0..1, model's self-reported certainty
    gate_passed: bool = False         # ok AND schema valid AND conf >= 0.6
    clarification_question: str = ""  # when gate fails but image was usable
    error: str = ""                   # when ok is False
    latency_s: float = 0.0
    model: str = ""


class VisionAnalyzer:
    """Calls the vision model, validates the schema, enforces the gate."""

    MIN_CONFIDENCE = 0.6

    def __init__(self, client: VisionClient, logger: CallLogger,
                 config: Optional[VoiceForgeConfig] = None,
                 timeout_s: Optional[float] = None,
                 min_confidence: Optional[float] = None,
                 model: str = ""):
        self.client = client
        self.logger = logger
        self.config = config or VoiceForgeConfig()
        self.timeout_s = timeout_s if timeout_s is not None \
            else self.config.vision_timeout_s
        self.min_confidence = min_confidence if min_confidence is not None \
            else self.config.min_confidence
        self.model = model

    def _prompt(self, context: VisionContext) -> str:
        turns = "\n".join(f"{t.role}: {t.text}"
                          for t in context.recent_turns[-6:])
        prompt = (f"{SYSTEM_PROMPT}\n\nBusiness: {context.business}\n"
                  f"The caller was asked to photograph: "
                  f"{context.photo_request}\n"
                  f"Recent conversation:\n{turns}")
        if context.locale_hints:
            prompt += f"\nLocale hints: {context.locale_hints}"
        return prompt

    @staticmethod
    def _validate_schema(raw: object) -> Optional[str]:
        """Return an error string, or None when the payload is valid."""
        if not isinstance(raw, dict):
            return "vision_schema_violation: not a JSON object"
        if not isinstance(raw.get("description"), str):
            return "vision_schema_violation: bad description"
        findings = raw.get("findings")
        if (not isinstance(findings, list) or not findings
                or not all(isinstance(f, str) and f.strip()
                            for f in findings)):
            return "vision_schema_violation: bad findings"
        conf = raw.get("confidence")
        if not isinstance(conf, (int, float)) or isinstance(conf, bool):
            return "vision_schema_violation: bad confidence"
        if not 0.0 <= float(conf) <= 1.0:
            return "vision_schema_violation: confidence out of range"
        return None

    def _clarify(self, description: str) -> str:
        return ("The photo is a bit unclear — " +
                (description.strip() or "I couldn't make out the details.") +
                " Could you describe what you see, or send a clearer photo?")

    def analyze(self, image_bytes: bytes, content_type: str,
                context: VisionContext) -> VisionResult:
        """Analyze one image. Never raises: failures -> ok=False results."""
        started = time.time()
        try:
            prompt = self._prompt(context)
            # N2: the 45 s budget is END-TO-END including the one retry:
            # 45 + backoff(1) + 45 was really ~91 s. Split the budget so
            # attempt + backoff + attempt can never exceed timeout_s.
            attempt_timeout = max(5.0, (self.timeout_s - 1.0) / 2.0)
            ok, result = call_with_retry(
                lambda: self.client.describe(
                    image_bytes, content_type, prompt, attempt_timeout),
                retries=1, backoff_s=1.0)
            latency = time.time() - started
            if not ok:
                err = result
                code = "vision_timeout" if isinstance(err, TimeoutError) \
                    else "vision_provider_error"
                return self._fail(code, latency, str(err))
            schema_err = self._validate_schema(result)
            if schema_err:
                return self._fail(schema_err, latency)
            confidence = float(result["confidence"])
            gate_passed = confidence >= self.min_confidence
            res = VisionResult(
                ok=True,
                description=result["description"].strip(),
                findings=[f.strip() for f in result["findings"]],
                confidence=confidence,
                gate_passed=gate_passed,
                latency_s=latency,
                model=self.model)
            if not gate_passed:
                res.clarification_question = self._clarify(res.description)
            self.logger.log("vision_analyzed", confidence=confidence,
                            gate_passed=gate_passed, latency_s=round(latency, 2),
                            model=self.model)
            return res
        except Exception as exc:  # noqa: BLE001 - never raise to call loop
            return self._fail("vision_internal_error",
                              time.time() - started, str(exc))

    def _fail(self, error: str, latency: float,
              detail: str = "") -> VisionResult:
        self.logger.log("vision_failed", error=error, detail=detail,
                        latency_s=round(latency, 2), model=self.model)
        return VisionResult(ok=False, error=error, latency_s=latency,
                            model=self.model)
