# VoiceForge v2 — hardened AI voice agent with live-call vision

> **For business owners (clinics, home services):** this is the engine
> behind an AI receptionist that answers every call your front desk
> misses, books the appointment, and can even look at a photo a caller
> texts mid-call. Start with [HANDOUT.md](HANDOUT.md) (one page),
> [CALL_SCRIPT.md](CALL_SCRIPT.md) (the 10-minute demo), and
> [PILOT_AGREEMENT.md](PILOT_AGREEMENT.md) (the 14-day pilot template).
> The money-back metric is buyer-visible via `VoiceAgent.pilot_summary()`.
> Everything below is the technical documentation.

A dependency-free Python engine for inbound/outbound voice calls (Twilio/Vapi
compatible) with a **live-call vision step**: mid-call, the agent can ask the
caller to text a photo, analyzes it, and continues the conversation with what
it saw — or falls back to voice-only without the caller ever hearing a
technical error.

v2 is a ground-up hardening of the v1 reference demo. Every one of the seven
pilot-killing bugs found in the v1 code review is fixed and covered by a
regression test.

## What v2 adds over v1

- **7 pilot-killer fixes** — fail-closed webhook secrets, LLM retry+fallback,
  idempotent turn handling (no Twilio retry storms), booking-trap fixes,
  word-boundary escalation ("reagent" no longer transfers you), thread-safe
  PII-redacting audit log, E.164-normalized DNC checks.
- **Booking correctness** — slot text normalized to real dates (no past
  dates), confirm-before-commit loop, per-day caps, no double bookings.
- **Warm human transfer** — announces the handoff, briefs the human with
  conversation/vision context via a whisper leg, clean fallback when nobody
  answers.
- **Live-call vision** (`voiceforge/vision/`) — `mms_watch` (verified MMS
  intake, idempotent, landline skip), `vision_analyze` (0.6 confidence gate:
  below it the agent may *ask*, never *assert*), `context_inject`
  (exactly-once delivery at turn boundaries, never mid-utterance).
- **Compliance** — AI identity disclosed in the greeting from call #1;
  written retention/deletion policy in `RETENTION.md`.
- **Latency budgets** — LLM ≤ 8 s, per-turn deadline ~12 s (Twilio's webhook
  timeout is ~15 s).

## Layout

```
voiceforge/
  config.py            # every timeout, budget and gate in one dataclass
  session.py           # CallSession (voice + vision state), MediaItem
  brain.py             # intents, hardened booking flow, escalation
  providers.py         # LLM/TTS/vision interfaces + stdlib HTTP clients
  adapters_twilio.py   # signature verify, TwiML, warm-transfer Dial
  adapters_vapi.py     # fail-closed secret verify, message parsing
  campaign.py          # outbound dialer with normalized DNC
  agent.py             # call lifecycle engine (idempotent turns)
  logging.py           # thread-safe, PII-redacting JSONL audit log
  util.py              # phone normalization, PII redaction
  vision/
    mms_watch.py       # MMS intake
    vision_analyze.py  # 0.6-gated vision analysis
    context_inject.py  # turn-boundary injection, exactly-once
    flow.py            # 90 s budget orchestrator + degradation ladder
test_voiceforge_v2.py  # 82 tests, stdlib only
run_demo.py            # end-to-end simulated day (voice + vision)
RETENTION.md           # PII/media retention & deletion policy
```

## Quick start

```bash
python3 test_voiceforge_v2.py   # 82 tests, no credentials, no network
python3 run_demo.py             # simulated inbound call + vision flow
```

## Wiring real providers

```python
from voiceforge import VoiceAgent, ConversationBrain, TwilioAdapter, \
    VapiAdapter, VoiceForgeConfig, CallLogger
from voiceforge.providers import HttpLLMClient, HttpVisionClient
from voiceforge.vision import MmsWatch, VisionAnalyzer, ContextInjector
from voiceforge.vision.flow import VisionFlow

config = VoiceForgeConfig.from_env(business="Swift Towing Co.")
brain = ConversationBrain(business="Swift Towing Co.", config=config)
agent = VoiceAgent(brain, logger=CallLogger("audit.jsonl"), config=config)
twilio = TwilioAdapter()   # reads TWILIO_AUTH_TOKEN, fails fast if missing
vapi = VapiAdapter()       # reads VAPI_WEBHOOK_SECRET, fails fast if missing

watch = MmsWatch(twilio, agent.logger, config)
flow = VisionFlow(watch,
                  VisionAnalyzer(HttpVisionClient(VISION_URL, VISION_KEY),
                                 agent.logger, config),
                  ContextInjector(agent.logger), agent.logger, config)
agent.vision_flow = flow
```

## Running the server (pilot)

The package is a library; `voiceforge/server.py` is the service Twilio
talks to. It binds every webhook route over stdlib HTTP (no dependencies):

| Method | Route | Purpose |
|---|---|---|
| `GET` | `/healthz` | liveness probe (`{"ok": true}`) |
| `POST` | `/voice`, `/gather` | Twilio voice webhooks (inbound + gather) |
| `POST` | `/whisper?sid=…` | warm-transfer whisper leg (absolute URL) |
| `POST` | `/transfer_fallback?sid=…` | no-answer fallback (absolute URL) |
| `POST` | `/messaging` | Twilio MMS side channel (vision) |
| `POST` | `/vapi` | Vapi webhook — verified + logged (Twilio-only pilot) |

```bash
export PUBLIC_BASE_URL="https://<your-host>"  # required — Twilio rejects relative URLs
export TWILIO_AUTH_TOKEN="<token>"            # required — signature verification
export TWILIO_ACCOUNT_SID="<sid>"             # media download + real SMS
export TWILIO_PHONE_NUMBER="+15551234567"     # enables real SMS confirmations
export VOICEFORGE_BUSINESS="Acme Services"
export VOICEFORGE_TRANSFER_NUMBER="+15557654321"
# optional: LLM_API_URL + LLM_API_KEY for a live brain
# (without them the server serves the mock and says so loudly at startup;
#  set VOICEFORGE_REQUIRE_LIVE_LLM=1 to fail fast instead)

python3 server_entry.py        # listens on $PORT (default 8080)
# or: docker build -t voiceforge-v2 . && docker run -p 8080:8080 --env-file .env voiceforge-v2
```

Then in the Twilio console, point the number's webhooks at:
- Voice: `{PUBLIC_BASE_URL}/voice` (HTTP POST)
- Messaging: `{PUBLIC_BASE_URL}/messaging` (HTTP POST)

For the pilot, expose the server with ngrok (`ngrok http 8080`) and use
the ngrok URL as `PUBLIC_BASE_URL`. The server fails fast at startup when
`PUBLIC_BASE_URL` is missing or when live mode is requested without
providers.

## Environment variables

| Variable | Required | Purpose |
|---|---|---|
| `PUBLIC_BASE_URL` | **yes (server)** | absolute callback URLs Twilio accepts |
| `TWILIO_AUTH_TOKEN` | for Twilio webhooks | HMAC-SHA1 signature verification |
| `TWILIO_ACCOUNT_SID` | for media/SMS | Basic Auth on media download, SMS sender identity |
| `TWILIO_PHONE_NUMBER` | for SMS | sender number; SMS promised only when set |
| `VAPI_WEBHOOK_SECRET` | for Vapi webhooks | `X-Vapi-Secret` verification |
| `VOICEFORGE_TRANSFER_NUMBER` | for warm transfer | human target, E.164 |
| `VOICEFORGE_AI_DISCLOSURE` | no (default `1`) | AI identity in greeting |
| `VOICEFORGE_RECORD_CALLS` | no (default `1`) | `<Record>` in inbound TwiML |
| `LLM_API_URL` / `LLM_API_KEY` / `LLM_MODEL` | for live brain | OpenAI-compatible chat endpoint |
| `VOICEFORGE_REQUIRE_LIVE_LLM` | no | `1` = fail fast instead of serving the mock |
| `VOICEFORGE_MEDIA_DIR` | no (default `./media`) | MMS photo storage |
| `PORT` | no (default `8080`) | listen port |

Secrets are read at startup and fail fast when missing. They are never
logged — the audit log stores only a truncated SHA-256 handle per number.

## Vision flow (pilot path)

1. `agent.request_photo(call_sid, "the damaged panel")` — spoken ask, 90 s
   budget starts. Landlines skip silently (voice-only, v1-identical path).
2. Twilio Messaging webhook → `flow.on_mms_webhook(...)` — verified,
   downloaded, attached to the active call (or orphan-parked ≤ 10 min).
3. `vision_analyze` — strict JSON schema, 0.6 gate, 45 s cap, never raises.
4. `context_inject` — gated results enter as a system turn at the next turn
   boundary; low confidence becomes a spoken clarification question; any
   failure becomes one spoken voice-only fallback line.
5. Budget expiry → voice-only fallback, exactly once.

## Test suite

82 tests, stdlib `unittest` only — no credentials, no network:

- ported v1 behaviors (greeting, FAQ, escalation, TwiML, Vapi, campaign)
- one regression test per pilot-killer (7)
- booking correctness (normalization, confirm loop, caps, double-book guard)
- warm transfer (Dial+whisper TwiML, unreachable-target fallback)
- vision modules with mocked providers (gate at 0.60/0.59, dedupe,
  mid-speech deferral, timeout, orphan media, landline skip)
