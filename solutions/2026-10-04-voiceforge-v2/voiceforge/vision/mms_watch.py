"""mms_watch — incoming media intake for the live-call vision step.

Owns the Twilio Messaging webhook: signature-verified, content-type and size
gated BEFORE download, idempotent on MessageSid, orphan-media TTL for
early-arriving photos, landline detection that skips the visual flow.
"""

from __future__ import annotations

import base64
import os
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

from ..adapters_twilio import TwilioAdapter
from ..config import VoiceForgeConfig
from ..logging import CallLogger
from ..session import MediaItem, VISION_UNAVAILABLE
from ..util import normalize_phone

EXT_BY_TYPE = {"image/jpeg": ".jpg", "image/png": ".png",
               "image/webp": ".webp"}


@dataclass
class MmsEvent:
    message_sid: str
    from_phone: str
    media: List[MediaItem]
    received_ts: float = field(default_factory=time.time)
    num_media: int = 0


# line_lookup(phone) -> "landline" | "mobile" | "voip" | "unknown"
LineLookup = Callable[[str], str]
# downloader(url) -> (bytes, content_type); injectable for tests
Downloader = Callable[[str], Tuple[bytes, str]]


def default_downloader(url: str, timeout_s: float = 10.0,
                       max_bytes: int = 5_000_000,
                       auth: Optional[Tuple[str, str]] = None
                       ) -> Tuple[bytes, str]:
    """
    Stream a download with a hard byte cap; aborts over the limit.

    B1: Twilio media URLs require HTTP Basic Auth (Account SID as username,
    Auth Token as password) — without it every fetch 401s and the vision
    step silently ignores every photo. `auth` is (username, password).
    """
    req = urllib.request.Request(url)
    if auth and auth[0] and auth[1]:
        token = base64.b64encode(
            f"{auth[0]}:{auth[1]}".encode()).decode()
        req.add_header("Authorization", f"Basic {token}")
    with urllib.request.urlopen(req, timeout=timeout_s) as resp:
        ctype = resp.headers.get("Content-Type", "").split(";")[0].strip()
        chunks: List[bytes] = []
        total = 0
        while True:
            chunk = resp.read(65536)
            if not chunk:
                break
            total += len(chunk)
            if total > max_bytes:
                raise ValueError(f"media exceeds {max_bytes} bytes")
            chunks.append(chunk)
        return b"".join(chunks), ctype


class MmsWatch:
    """Intake for MMS photos during (or just before) an active call."""

    def __init__(self, twilio: TwilioAdapter,
                 logger: CallLogger,
                 config: Optional[VoiceForgeConfig] = None,
                 media_dir: str = "media",
                 line_lookup: Optional[LineLookup] = None,
                 downloader: Optional[Downloader] = None):
        self.twilio = twilio
        self.logger = logger
        self.config = config or VoiceForgeConfig()
        self.media_dir = media_dir
        os.makedirs(media_dir, exist_ok=True)
        # lookup failure -> assume NOT landline (fail open; the 90 s timer
        # is the backstop)
        self.line_lookup: LineLookup = line_lookup or (lambda p: "unknown")
        # B1: the default downloader authenticates with the adapter's
        # credentials — Twilio media URLs 401 without HTTP Basic Auth.
        tw = twilio

        def _authed_download(url: str) -> Tuple[bytes, str]:
            sid = getattr(tw, "account_sid", "") or ""
            auth = (sid, tw.auth_token) if sid and tw.auth_token else None
            return default_downloader(
                url, max_bytes=self.config.max_media_bytes, auth=auth)

        self.downloader = downloader or _authed_download
        self.seen_sids: set = set()          # MessageSid idempotency
        self.orphans: Dict[str, Tuple[MmsEvent, float]] = {}
        self._line_cache: Dict[str, Tuple[str, float]] = {}
        self._sleep = time.sleep  # injectable; tests set to a no-op

    # -- webhook entry ----------------------------------------------------
    def handle_webhook(self, url: str, form: Dict[str, str],
                       signature: str) -> Optional[MmsEvent]:
        """
        Process a Twilio Messaging webhook form. Returns an MmsEvent, or
        None for: bad signature (caller must answer 403), body-only texts,
        duplicates, or rejected media.
        """
        if not self.twilio.verify_signature(url, form, signature):
            self.logger.log("mms_rejected", reason="bad_signature")
            return None
        sid = form.get("MessageSid", "")
        if not sid:
            return None
        if sid in self.seen_sids:
            self.logger.log("mms_duplicate", message_sid=sid)
            return None  # duplicate delivery: no-op
        try:
            num_media = int(form.get("NumMedia", "0") or "0")
        except ValueError:
            num_media = 0
        if num_media == 0:
            return None  # body-only text; not our concern

        from_phone = form.get("From", "")
        items: List[MediaItem] = []
        for i in range(num_media):
            media_url = form.get(f"MediaUrl{i}", "")
            content_type = form.get(f"MediaContentType{i}", "").split(";")[0]
            if not media_url or content_type not in \
                    self.config.allowed_media_types:
                self.logger.log("mms_media_rejected",
                                reason="content_type",
                                content_type=content_type)
                continue
            local_path: Optional[str] = None
            try:
                data, seen_type = self._download(media_url)
                ext = EXT_BY_TYPE.get(content_type, ".bin")
                local_path: Optional[str] = os.path.join(
                    self.media_dir, f"{sid}_{i}{ext}")
                # B6: the file write lives INSIDE the try — a disk-full or
                # permissions failure degrades to "media rejected", never a
                # 500 that Twilio retries forever.
                with open(local_path, "wb") as fh:
                    fh.write(data)
            except Exception as exc:  # noqa: BLE001 - per-item containment
                reason = ("write_failed" if isinstance(exc, OSError)
                          else "download_failed")
                self.logger.log("mms_media_rejected", reason=reason,
                                error=str(exc))
                if local_path and os.path.exists(local_path):
                    try:
                        os.remove(local_path)  # no partial files left behind
                    except OSError:
                        pass
                continue
            items.append(MediaItem(
                media_sid=f"{sid}_{i}", message_sid=sid,
                content_type=content_type, size=len(data),
                local_path=local_path))
            _ = seen_type
        if not items:
            return None
        self.seen_sids.add(sid)
        event = MmsEvent(message_sid=sid, from_phone=from_phone,
                         media=items, num_media=num_media)
        self.logger.log("mms_received", message_sid=sid,
                        from_phone=from_phone, items=len(items))
        return event

    def _download(self, url: str) -> Tuple[bytes, str]:
        """3 attempts, exponential backoff (1s, 2s, 4s), 10 s per attempt."""
        last: Exception = ValueError("download failed")
        for attempt in range(3):
            try:
                return self.downloader(url)
            except urllib.error.HTTPError as exc:
                # N9: non-transient 4xx (bad auth, gone, forbidden) will
                # never succeed on retry — fail fast instead of burning the
                # webhook's time budget.
                if 400 <= exc.code < 500 and exc.code not in (408, 429):
                    raise
                last = exc
                self._sleep(float(1 << attempt))
            except Exception as exc:  # noqa: BLE001
                last = exc
                self._sleep(float(1 << attempt))
        raise last

    # -- session attachment -------------------------------------------------
    def attach_to_session(self, agent: "VoiceAgent",  # noqa: F821
                          event: MmsEvent) -> Optional[str]:
        """
        Attach media to the most-recent ACTIVE session whose phone matches
        (normalized E.164). No match -> park in orphans with TTL; returns the
        call_sid or None.
        """
        want = normalize_phone(event.from_phone)
        best_sid: Optional[str] = None
        best_ts = -1.0
        # B3: iterate a snapshot — inbound_call() mutates agent.calls from
        # other webhook threads; iterating the live dict raised
        # "dictionary changed size during iteration" -> 500 -> retry storm.
        for sid, session in agent.calls_snapshot():
            if session.state != "active":
                continue
            if normalize_phone(session.phone) == want \
                    and session.started_at > best_ts:
                best_sid, best_ts = sid, session.started_at
        if best_sid is None:
            self.orphans[event.message_sid] = (
                event, time.time() + self.config.orphan_media_ttl_s)
            self.logger.log("mms_orphaned", message_sid=event.message_sid)
            return None
        sessions = dict(agent.calls_snapshot())
        session = sessions[best_sid]
        with session.lock:
            have = {m.media_sid for m in session.media}
            for item in event.media:
                if item.media_sid not in have:
                    session.media.append(item)
        self.logger.log("mms_attached", message_sid=event.message_sid,
                        call_sid=best_sid)
        return best_sid

    def reap_orphans(self, agent: "VoiceAgent") -> int:
        """Retry orphan attachment; drop expired ones. Returns attached count.

        N3: expired orphans have their media files deleted from disk —
        RETENTION.md promises deletion, so the engine actually does it.
        """
        now = time.time()
        attached = 0
        for sid in list(self.orphans):
            event, expiry = self.orphans[sid]
            if now > expiry:
                self._delete_event_files(event)
                self.logger.log("mms_orphan_purged", message_sid=sid)
                del self.orphans[sid]
                continue
            if self.attach_to_session(agent, event):
                attached += 1
                del self.orphans[sid]
        self.purge_media_files(self.config.media_retention_s)
        return attached

    def _delete_event_files(self, event: MmsEvent) -> int:
        """Delete an event's media files; returns files removed."""
        removed = 0
        for item in event.media:
            try:
                if item.local_path and os.path.exists(item.local_path):
                    os.remove(item.local_path)
                    removed += 1
            except OSError:
                pass
        return removed

    def purge_media_files(self, max_age_s: float) -> int:
        """
        Delete media files older than max_age_s (RETENTION.md: MMS photos
        are kept 30 days). Returns files removed. Never raises.
        """
        now = time.time()
        removed = 0
        try:
            entries = list(os.scandir(self.media_dir))
        except OSError:
            return 0
        for entry in entries:
            try:
                if entry.is_file() and now - entry.stat().st_mtime > max_age_s:
                    os.remove(entry.path)
                    removed += 1
            except OSError:
                continue
        if removed:
            self.logger.log("media_purged", files=removed)
        return removed

    # -- landline detection ---------------------------------------------------
    def is_landline(self, phone: str) -> bool:
        """True when the number is a landline -> skip the visual flow."""
        now = time.time()
        cached = self._line_cache.get(phone)
        if cached and now - cached[1] < 86400:
            return cached[0] == "landline"
        try:
            kind = self.line_lookup(phone)
        except Exception as exc:  # noqa: BLE001 - fail open, log it
            self.logger.log("line_lookup_failed", error=str(exc))
            return False
        self._line_cache[phone] = (kind, now)
        if kind == "landline":
            self.logger.log("landline_detected", phone=phone)
            return True
        return False

    def mark_unavailable(self, session) -> None:
        session.vision_state = VISION_UNAVAILABLE
