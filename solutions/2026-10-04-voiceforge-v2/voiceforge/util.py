"""Small shared helpers: phone normalization and PII redaction."""

from __future__ import annotations

import re

_PHONE_RUN = re.compile(r"\+?\d[\d\-\s().]{5,}\d")


def normalize_phone(raw: str) -> str:
    """
    Normalize a phone number to a comparable digit string for DNC checks
    and session matching.

    Strips every non-digit; drops a single leading North-American country
    code '1' when the result is 11 digits long so that "+1-555-000-0001"
    and "+15550000001" compare equal. Returns "" for unusable input.
    """
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if not 7 <= len(digits) <= 15:  # E.164 range sanity check
        return ""
    return digits


def valid_phone(raw: str) -> bool:
    """True when the number normalizes to a plausible digit string."""
    return normalize_phone(raw) != ""


def redact_pii(text: str) -> str:
    """
    Redact phone-number-like runs from free text before it hits the audit
    log. Conservative on purpose: any run of 7+ digits (with common
    separators) becomes [REDACTED_PHONE].
    """
    if not text:
        return text
    return _PHONE_RUN.sub("[REDACTED_PHONE]", text)
