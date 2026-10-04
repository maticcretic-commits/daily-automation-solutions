# Free testing guide — VoiceForge v2

How to test a voice AI agent end-to-end for $0. No paid spend, no API
keys beyond free signups. Facts below were verified against official
docs/pages opened 2026-10-04 unless marked otherwise.

## The cheapest genuinely-free path to a REAL voice test

**Twilio trial + ngrok (or Cloudflare) tunnel + your own phone. $0.**

1. Sign up at twilio.com/try-twilio — **no credit card required**
   (official docs: https://www.twilio.com/docs/usage/tutorials/how-to-use-your-free-trial-account).
   You get **75 voice minutes, 100 SMS, 100 WhatsApp messages** as free
   units. Trial **expires after 30 days**.
2. Verify your own mobile number in the Console (Phone Numbers →
   Verified Caller IDs). Trial accounts can **only call verified
   numbers, in your sign-up country** — for Nitesh that means his own
   Indian mobile, which is exactly what a first test needs.
3. Run the VoiceForge server locally: `python3 server_entry.py`
   (see README for env vars), then expose it:
   `ngrok http 8000` — free tier gives 3 endpoints, 1 GB/month, 20k
   requests/month (third-party research, 2026 — treat limits as
   approximate). Alternative with no account at all: Cloudflare Quick
   Tunnels (`cloudflared tunnel --url http://localhost:8000`).
   Caveat: ngrok free shows an interstitial warning page on browser
   visits; Twilio webhook POSTs are unaffected in practice, but if a
   webhook misbehaves, suspect the tunnel first.
4. In the Twilio Console, point a voice webhook at
   `https://<your-tunnel>/voice` (the route server.py binds).
5. Call the Twilio number from your verified mobile. You are now talking
   to VoiceForge through a real carrier: real STT latency, real
   turn-taking, real interruptions.
6. Watch the 75-minute budget in the Console's free-units tracker.

That is ~75 minutes of real voice testing for $0 and about an hour of
setup. Nothing else on this page exercises the true voice path that
cheaply.

## Other free options, honestly graded

| Option | Free terms (verified) | Exercises the voice path? |
|---|---|---|
| **Twilio trial** (above) | 75 voice min, 30-day expiry, verified numbers only, same-country only (official docs) | **Yes — full PSTN path** |
| **Vapi free tier** | $5 credits, 1 phone number, 4 concurrent calls (vapi.ai/pricing, via prior research) | Yes, but tests *Vapi's* stack, not VoiceForge |
| **ElevenLabs free tier** | $0 tier exists (prior research); good for TTS voice-quality comparison | Partial — TTS only |
| **SIPp** (open source) | Free forever; scripted SIP call flows with RTP audio playback (sipp GitHub) | Signaling only — never touches the AI brain |
| **Asterisk/FreeSWITCH + softphone** (open source) | Free forever; run a local PBX, call it from Linphone (free softphone) | Yes for SIP/RTP — but not Twilio webhooks; best for the self-hosted story |
| **Cloudflare Quick Tunnels** | Free, no account (third-party research) | Tunnel only — pairs with the Twilio trial |
| **Zero telephony** (this repo) | Free forever | Brain/booking/vision logic only: `python3 scenario_runner.py` (20/20), `python3 -m unittest` (103 tests), `python3 run_demo.py` |

## What each layer actually proves

- **Scenario/unit tests (free, now):** the conversation brain is correct —
  booking, escalation, vision gating, fail-closed webhooks. 20/20 + 103
  green is real evidence, but every provider is a mock.
- **Twilio trial (~75 min, $0):** adds the real carrier — STT accuracy on
  your accent, turn latency end-to-end, barge-in behavior, whisper-leg
  transfers, MMS delivery timing. This is the layer the partner gate
  demands before a paid pilot.
- **~20 live calls:** enough to meter real per-call cost (STT+LLM+TTS+
  telephony) and reprice the pilot at 3–4× measured COGS. Still $0 on
  trial credit if you stay inside 75 minutes.

## India-specific notes

- Trial voice is restricted to the sign-up country: an Indian signup
  tests India-to-India calls — fine for the pilot.
- India long-codes carry heavy DLT registration requirements for SMS;
  for the voice pilot this does not matter. WhatsApp testing: use the
  Twilio WhatsApp Sandbox (free, `join <code>` from your phone) —
  no Meta Business verification needed for sandbox.
- The vision step's MMS side-channel is weak in India (no MMS culture);
  the WhatsApp side-channel is the one to test — the sandbox covers it.

## What is NOT free (don't bother until revenue)

- A Twilio Indian virtual number with full SMS/voice (paid, DLT
  overhead), Vapi/Retell/Bland paid tiers, ElevenLabs paid minutes,
  any "free trial" that asks for a credit card first (e.g. Air AI's
  infamous $23 top-up — see review-sentiment.md; avoid).
