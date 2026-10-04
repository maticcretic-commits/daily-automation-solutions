# VoiceForge v2 vs. major voice-AI platforms

> Spec/review-based comparison — researched 2026-10-04. No live API calls were made (no keys exist). Vendor cells come from `competitors.md` / `review-sentiment.md`; VoiceForge cells come from the local code + 20/20 scenario benchmark. 'Not covered in research' means the research found nothing — not proof of absence.

## Booking correctness

**VoiceForge v2** — Confirm-before-commit booking loop with real date normalization ('tomorrow' -> actual date), past-date rejection, per-day caps, no double-booking, day corrections re-resolved ('yes, friday instead'). Verified by 20/20 local scenario benchmark.

**Vapi** — Not covered in research — booking behavior is builder-defined via Vapi's function tools/workflows; no platform-level booking-correctness guarantee found.

**Retell AI** — Not covered in research — builder-defined; no platform-level booking-correctness guarantee found.

**Bland AI** — Not covered in research — Pathways builder is builder-defined; no platform booking guarantee found.

**ElevenLabs** — Not covered in research for ElevenAgents — builder-defined; no platform booking guarantee found.

## Human handoff

**VoiceForge v2** — Warm transfer: announces the handoff, dials the human with a whisper leg carrying conversation + vision context, falls back to message-taking when the target is unreachable (guarded, never 500s). Carrier behavior NOT yet tested on a live call.

**Vapi** — Call transfer supported; no review-derived handoff quality data found in research.

**Retell AI** — WEAKEST in set: failed transfers are the #1 forum topic ('telephony provider or carrier declined'); +1-2s added on warm transfers (community reports).

**Bland AI** — No review-derived handoff quality data found in research.

**ElevenLabs** — No review-derived handoff quality data found in research.

## Mid-call vision

**VoiceForge v2** — Native mid-call vision step: caller texts a photo (MMS/WhatsApp side channel), vision_analyze with strict JSON schema + 0.6 confidence gate (below gate: ask, never assert), context injected into the live call. Flagship differentiator; verified in local tests only.

**Vapi** — No native mid-call vision. Third-party hackathons wire external vision models via function tools (DIY); knowledge base warns image-based PDFs are unreadable.

**Retell AI** — No public documentation of mid-call vision. SMS exists as text-only flows; no MMS/image input documented.

**Bland AI** — No public documentation of mid-call vision. Feature lists cover Pathways, voice cloning, SMS (enterprise-only) — no image/vision on calls.

**ElevenLabs** — Closest rival on vision: agents handle images/PDFs on the WHATSAPP messaging channel (May 2026), but NO public documentation of vision on live telephony phone calls.

## Pricing + transparency

**VoiceForge v2** — Self-hosted code (no per-minute platform fee); recommended pilot pricing fixed Rs.25,000 India / $399 international, 14 days, 50% advance, written money-back on a named metric. Real per-call COGS unmeasured — flagged as estimate.

**Vapi** — $0.05/min hosting + model costs passed through (Deepgram STT, LLM, ElevenLabs TTS); worked example 1,000 min ~ $82-129/mo all-in. $5 free credits, 1 phone number, 4 concurrent calls. Review complaint: '$0.05 headline sitting on a stack that multiplies it by three to five.'

**Retell AI** — $0.07-$0.31/min orchestration + LLM + TTS + telephony billed separately; cheapest working call ~$0.087/min. $10 free credits. Review complaint: prompts over 4,000 tokens scale the billed duration of the ENTIRE call; token drift with no config change.

**Bland AI** — Start $0 + $0.14/connected min; Build $299/mo + $0.12/min; Scale $499/mo + $0.11/min (LLM+STT+TTS bundled). Review complaint: Dec 2025 price rise +56% is the #1 churn driver; billing docs contradict the pricing FAQ on telephony inclusion.

**ElevenLabs** — Credit-based subscriptions (Free $0 ... Business $990/mo) + pay-as-you-go; conversational AI ~$0.08/min. May 2026 reset cut TTS -55%, STT -45%, agents -20%. Review complaint: credit model is the #1 complaint cluster — credits burned on failed output, forfeited on cancellation, no pre-run cost visibility.

## Latency (claimed vs review-reported)

**VoiceForge v2** — Claimed: not published. Measured: unmeasured — all tests use mock TTS/STT/LLM. Review corpus: n/a (no users yet). Honest status: unknown until ~20 live calls are metered.

**Vapi** — Claimed: not published as a single number. Review-reported: 'unpredictable latency' across G2 reviews; users ask for better latency control and interruption handling.

**Retell AI** — Claimed: ~600ms round-trip. Review-reported: ~3,000ms configured-vs-actual gaps, ~5,000ms in Rigid mode vs ~1,300ms Flex (community threads); Trustpilot: 'latency didnt support fast paced conversations.'

**Bland AI** — Claimed: ~400ms (Speech v3). Review-reported: won an independent 19,363-call verification test (4.19% vs 3.29% pooled) — strongest measured latency-adjacent evidence in the set.

**ElevenLabs** — Claimed: not published as a single number. Review-reported: voice quality is the category benchmark; latency complaints not prominent in the review corpus (TTS heritage).

## Top review-derived weakness

**VoiceForge v2** — Zero production history: no live-carrier testing, no measured latency/cost, one maintainer at 14 hrs/week, no dashboard, no enterprise features (HIPAA/SSO), no review corpus at all.

**Vapi** — Support: Trustpilot 2.4/5 on lag + slow support; 'Support is non-existent' (Ringly). No no-code path; dashboard 'overwhelming'. $50M Series B scale is real (1B+ calls claimed) but the long tail complains.

**Retell AI** — Support after payment — the entire 1-star tail ('no Support no one here to help after making payment'; Discord-only). Outbound gated by 1-2 week KYC ('they wanted my f'ing passport'). Country whitelists (Turkey 403, 'unusable in the UK').

**Bland AI** — Cost unpredictability + the 56% hike destroyed trust; Start plan capped at 10 concurrent / 100 calls/day; review corpus tiny (G2 5.0 from 11 reviews) — almost no independent verification exists.

**ElevenLabs** — Billing/support: Trustpilot 3.1 with 'horrible customer service' cluster, ticket backlogs admitted by staff, automated abuse-flags locking paying accounts, cancellation friction. Feature sprawl overwhelms newcomers.

## Where VoiceForge v2 wins (honest)

- Mid-call vision as a first-class, developer-facing step (photo -> analysis -> continued spoken answer with a 0.6 ask-don't-assert gate) — no dev platform in the set productizes this.
- Booking correctness as a built-in: confirm-before-commit loop, real date normalization, no double-booking — the exact churn event reviewers never praise incumbents for.
- Fail-closed engineering: webhook secrets reject empty, idempotent turns (no retry storms), PII-redacting audit log — the failure modes reviewers complain about (fragility, opaque failures) are designed out.
- Fixed, transparent pilot pricing with a written money-back metric — the direct antidote to the #1 and #2 cross-incumbent complaints (unpredictable billing, support collapse).
- India-first pragmatics: WhatsApp side-channel design and Hinglish roadmap target the market the US-centric incumbents price out.

## Where VoiceForge v2 loses (honest — no cherry-picking)

- Production proof: Vapi claims 1B+ calls, Retell holds G2 4.8/5 from 2,638 reviews. VoiceForge v2 has zero live calls, zero reviews, zero measured latency or cost — every number on our side is local-test or estimate.
- Voice quality: ElevenLabs is the category benchmark for natural, expressive speech. VoiceForge v2's TTS/STT are pluggable with only mock clients tested; real voice quality is unproven.
- No-code/dashboard: every incumbent ships a dashboard (even if users call Vapi's 'overwhelming'); we have none — operator dashboard is an explicit 'before scale' gap.
- Enterprise trust: HIPAA ($2,000/mo on Vapi), SSO, data residency, concurrency guarantees — we have none of these, and one maintainer at ~14 hrs/week cannot staff an SLA the incumbents sell.
- Distribution and docs: incumbents have docs, SDKs, Discord communities, and template galleries. We have a README and a demo script.
- Vision moat is thin: Quiq already markets mid-call photo-to-AI (thin evidence), and ElevenLabs could extend WhatsApp vision to telephony. Our window is execution speed, not IP.

_One-line verdict: the only dev platform with a real mid-call vision step and fail-closed engineering — but with zero production calls it is a prototype challenging companies with billions of calls, and every trust number on our side is still unearned._
