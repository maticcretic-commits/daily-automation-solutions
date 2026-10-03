# VoiceForge — AI Voice Agent Engine (Vapi + Twilio style)

A working, dependency-free Python engine that implements the full call
lifecycle of a production AI voice agent, built for the kind of gig that
asks for "Vapi + Twilio AI Voice Agent for Inbound & Outbound Calls":

```
phone call -> telephony webhook (Twilio / Vapi) -> STT transcript
         -> conversation brain (intents + LLM) -> TTS reply
         -> action (book / escalate / hang up) -> JSONL audit log
```

Everything external (telephony, speech-to-text, LLM, text-to-speech) sits
behind small provider interfaces with **mock** implementations, so the
entire engine runs end-to-end with zero credentials and zero network —
and swaps to real providers (Twilio, Vapi, OpenAI-compatible LLM APIs)
by passing a real client into the same slots.

## What's inside

| Piece | File | What it does |
|---|---|---|
| `VoiceAgent` | `voice_agent.py` | Call lifecycle engine: inbound greeting, per-turn handling, outbound campaigns, day summary |
| `ConversationBrain` | `voice_agent.py` | Intent routing: escalation → goodbye → 3-step booking slot-fill → FAQ → LLM fallback |
| `TwilioAdapter` | `voice_agent.py` | `X-Twilio-Signature` verification (HMAC-SHA1) + TwiML replies with `<Gather input="speech">` |
| `VapiAdapter` | `voice_agent.py` | `X-Vapi-Secret` verification + server-message parsing (assistant-request / transcript / end-of-call-report) |
| `HttpLLMClient` | `voice_agent.py` | OpenAI-compatible chat-completions client (stdlib `urllib` only) — drop in any API key |
| `Outbound campaign` | `voice_agent.py` | CSV dialing with do-not-call list, per-minute throttle, no-answer handling, per-call report |
| `CallLogger` | `voice_agent.py` | Append-only JSONL audit log of every call event |

## Run it

```bash
python3 voice_agent.py --demo     # end-to-end simulation: inbound call,
                                  # booking flow, escalation, Vapi message,
                                  # outbound campaign with DNC
python3 test_voice_agent.py       # 22 tests: intents, booking, Twilio
                                  # signatures, Vapi messages, campaigns
```

## Try a booking flow in 6 lines

```python
from voice_agent import VoiceAgent, ConversationBrain

agent = VoiceAgent(ConversationBrain(business="Swift Towing Co."))
print(agent.inbound_call("CA1", "+15551234567"))
print(agent.caller_said("CA1", "I want to book a tow"))
print(agent.caller_said("CA1", "John Carter"))   # name
print(agent.caller_said("CA1", "flatbed tow"))   # service
print(agent.caller_said("CA1", "tomorrow"))      # day -> confirmed
```

## Wiring to real providers

```python
twilio = TwilioAdapter(auth_token="<twilio auth token>",
                       gather_action_url="https://yourapp.com/voice")
# in your webhook handler:
if twilio.verify_signature(request.url, request.form, request.headers["X-Twilio-Signature"]):
    return twilio.inbound_twiml(agent.inbound_call(...))

llm = HttpLLMClient(api_url="https://api.openai.com/v1/chat/completions",
                    api_key="<key>", model="gpt-4o-mini",
                    system_prompt="You are a concise phone receptionist.")
brain = ConversationBrain(business="Your Business", llm=llm)
```

The mock clients keep every path testable offline; the real clients need
only an API key. No secrets are stored anywhere in this repo.

## Design notes

- **Intent order matters:** escalation is checked before anything else so a
  caller asking for a human is never trapped in a booking flow.
- **Booking never writes bad data:** the slot-fill asks again until all three
  slots (name, service, day) are captured; nothing is confirmed early.
- **Security:** Twilio signatures are verified with `hmac.compare_digest`
  (timing-safe); the demo secret is a placeholder, not a real credential.
- **Standard library only** — no pip install, runs on any Python 3.8+.
