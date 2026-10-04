"""Outbound campaign dialer with DNC safety.

P0-1(7): DNC comparison happens on NORMALIZED E.164 digit strings, so
"+1-555-000-0001" and "+15550000001" are the same number. Campaign CSVs are
validated (header + phone format) and one bad row can never abort a run.
"""

from __future__ import annotations

import csv
import secrets
import time
from typing import Callable, Dict, List, Optional, Set

from .brain import ConversationBrain
from .logging import CallLogger
from .providers import TTSClient, MockTTSClient
from .session import CallSession
from .util import normalize_phone, valid_phone

REQUIRED_COLUMNS = {"phone"}


class CampaignError(ValueError):
    """Raised for malformed campaign input before any dial happens."""


def _validate_rows(numbers_csv: str) -> List[Dict[str, str]]:
    text = (numbers_csv or "").strip()
    if not text:
        raise CampaignError("campaign CSV is empty")
    reader = csv.DictReader(text.splitlines())
    headers = {h.strip().lower() for h in (reader.fieldnames or [])}
    if not REQUIRED_COLUMNS.issubset(headers):
        raise CampaignError(
            f"campaign CSV must have a 'phone' column; got {reader.fieldnames}")
    rows = []
    for i, row in enumerate(reader, start=2):  # 1-based incl. header
        phone = (row.get("phone") or row.get("Phone") or "").strip()
        name = (row.get("name") or row.get("Name") or "").strip()
        if not phone:
            raise CampaignError(f"row {i}: missing phone number")
        if not valid_phone(phone):
            raise CampaignError(f"row {i}: invalid phone number {phone!r}")
        rows.append({"phone": phone, "name": name})
    if not rows:
        raise CampaignError("campaign CSV has no data rows")
    return rows


class CampaignRunner:
    """Dial through a CSV of numbers with DNC, throttling and per-row safety."""

    def __init__(self, brain: ConversationBrain,
                 tts: Optional[TTSClient] = None,
                 logger: Optional[CallLogger] = None):
        self.brain = brain
        self.tts = tts or MockTTSClient()
        self.logger = logger or CallLogger()
        self.sessions: Dict[str, CallSession] = {}

    def run(self, numbers_csv: str, opener: str,
            throttle_per_min: int = 20,
            dnc: Optional[List[str]] = None,
            on_transcript: Optional[Callable[[str], Optional[str]]] = None,
            sleep: Callable[[float], None] = time.sleep
            ) -> List[Dict[str, str]]:
        """
        Returns a per-call report. `sleep` is injectable so tests never wait.
        """
        rows = _validate_rows(numbers_csv)
        dnc_set: Set[str] = {normalize_phone(n) for n in (dnc or [])}
        report: List[Dict[str, str]] = []
        min_interval = 60.0 / max(throttle_per_min, 1)
        last_dial = 0.0

        for row in rows:
            phone, name = row["phone"], row["name"]
            norm = normalize_phone(phone)
            try:
                if norm in dnc_set:
                    report.append({"phone": phone, "status": "skipped_dnc"})
                    self.logger.log("campaign_skip", phone=phone,
                                    reason="do_not_call")
                    continue
                wait = min_interval - (time.time() - last_dial)
                if wait > 0:
                    sleep(wait)
                last_dial = time.time()

                sid = f"CA{secrets.token_hex(6)}"
                session = CallSession(call_sid=sid, phone=phone,
                                      direction="outbound", state="active",
                                      slots={"name": name})
                self.sessions[sid] = session
                opener_text = opener.format(name=name or "there")
                session.say("agent", opener_text)
                callee = on_transcript(opener_text) if on_transcript else None
                if callee is None:
                    session.hangup(reason="no_answer")
                    status = "no_answer"
                else:
                    self.brain.handle(session, callee)
                    session.hangup(reason="outbound_complete")
                    status = "completed"
                report.append({"phone": phone, "name": name, "status": status,
                               "sid": sid, "turns": str(len(session.turns))})
                self.logger.log("campaign_call", phone=phone, status=status,
                                sid=sid)
            except Exception as exc:  # noqa: BLE001 - per-row containment
                report.append({"phone": phone, "name": name,
                               "status": "error", "error": str(exc)})
                self.logger.log("campaign_row_error", phone=phone,
                                error=str(exc))
        return report
