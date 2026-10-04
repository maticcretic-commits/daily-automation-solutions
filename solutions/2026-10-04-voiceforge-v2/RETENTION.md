# VoiceForge v2 — PII & Media Retention / Deletion Policy

Applies from the first pilot call. The engine enforces the technical parts;
the operator (you) owns the calendar parts.

## What is collected

| Data | Where | Why |
|---|---|---|
| Call transcripts (caller + agent turns) | JSONL audit log | QA, dispute handling, pilot metrics |
| Phone numbers | **never stored raw** — truncated SHA-256 handle only | call correlation without reversible PII |
| MMS photos (faces, IDs, license plates possible) | `media/` dir, keyed by MessageSid | the live-call vision step |
| Vision analysis results (description, findings, confidence) | JSONL audit log | provenance for what the agent said |

Raw image bytes are **never** written to the audit log — only the structured
result. Reply text is XML-escaped for element-text use only.

## Retention windows

| Data | Retain | Then |
|---|---|---|
| Transcripts + structured results | 90 days | auto-delete |
| Phone-number hashes | 30 days | auto-delete (transcripts keep working; correlation ends) |
| MMS photos | 30 days | secure delete (overwrite + unlink) |
| Orphaned MMS (no matching call) | 10 minutes | dropped without analysis |

## Deletion

- A caller saying "delete my data" → flag the call SID; purge its media
  immediately and its log lines within 72 hours.
- End of pilot → export the metrics summary, then purge all media and
  transcripts older than the windows above.
- Model providers (LLM/vision APIs) are data processors: prefer vendors
  with zero-retention API tiers for the pilot; record the choice per pilot.

## Call-time disclosures (wired into the greeting)

1. "I'm the AI assistant for {business}." — AI identity, call #1.
2. "This call may be recorded to help us serve you better." — recording
   notice, call #1.

Both are on by default (`VOICEFORGE_AI_DISCLOSURE=1`); turning them off is
a conscious operator decision, not an accident.

## What the engine guarantees

- `CallLogger` redacts phone-number-like digit runs from every logged string
  and persists only the number hash.
- `MmsWatch` gates content-type and size *before* download; rejects
  non-images; stores bytes on disk (never RAM-only) under
  `media/<MessageSid>_<i>.<ext>`.
- `VisionAnalyzer` logs model, latency, confidence and gate outcome — never
  raw image bytes.
- Webhook secrets are fail-closed: empty secret = startup error, never an
  open endpoint.
