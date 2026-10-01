# WhatsFlow — WhatsApp Business Cloud API automation bot

**Solved client problem (from a real gig):** a small business wants its WhatsApp
Business number to answer customer messages automatically (greeting, pricing
FAQ, business hours, demo booking with lead capture, human handoff) and to
send template broadcast campaigns with dedupe, opt-out handling and rate
limiting — without hand-writing Meta's webhook boilerplate.

This demo implements exactly that against the official
**WhatsApp Business Cloud API** contract (webhook verification, message
webhooks, `X-Hub-Signature-256`, message payloads), so it can be pointed at
a real Meta app by swapping the console sender for the included Cloud API
sender and registering the webhook URL.

## What it does

- **Webhook server** (`server`): Meta's GET verification handshake, then POST
  handling with HMAC-SHA256 signature validation. Every inbound text message
  gets a bot reply sent back through the Cloud API payload builder.
- **Conversation engine**: greeting → numbered menu → FAQ (pricing, hours,
  demo) → two-step lead capture (name, interest) → human handoff; per-phone
  state persisted in SQLite. `STOP` unsubscribes, `START` resubscribes.
- **Outbound payloads**: text, message templates (with variables) and
  interactive quick-reply buttons — all in Cloud API format.
- **Broadcaster** (`broadcast`): takes a CSV of phones, dedupes, skips
  invalid numbers and opted-out contacts, throttles sends per second, and
  logs every attempt to SQLite.
- **Ops CLI**: `simulate` (try the bot offline), `export` (leads CSV),
  `report` (message/lead/handoff counts).

## Run it

No dependencies — Python 3.10+ stdlib only.

```bash
# 1. Try the bot offline (you play the customer)
python3 whatsflow.py simulate
#   you> hello
#   bot>  Hello Demo User! Welcome to Acme Automation. How can I help...

# 2. Run the webhook (point Meta at the public URL with verify token)
python3 whatsflow.py server --port 8000 \
  --verify-token YOUR_VERIFY_TOKEN --app-secret YOUR_APP_SECRET

# 3. Send a template campaign from CSV (console mock prints sends)
python3 whatsflow.py broadcast --csv campaign_sample.csv --template hello_world

# 4. Export captured leads / print stats
python3 whatsflow.py export --out leads.csv
python3 whatsflow.py report
```

```bash
# 5. Run the test suite (24 tests, covers handshake, signatures,
#    conversation flows, opt-out, dedupe, payload shapes)
python3 test_whatsflow.py
```

## Going live with Meta (checklist)

1. Create a Meta app → add the WhatsApp product → get a phone number ID.
2. Set `VERIFY_TOKEN` / `APP_SECRET` to your own values and run `server`
   behind HTTPS (e.g. ngrok, Cloudflare tunnel, or any VPS).
3. Register the public `https://…/` URL as the webhook with Meta and
   subscribe to the `messages` field.
4. Replace `ConsoleSender` with `CloudApiSender(phone_number_id, access_token)`
   — one line in `cmd_server` / `WebhookHandler.sender`.
5. Create your message templates in WhatsApp Manager first (Meta requires
   pre-approved templates for business-initiated messages); pass template
   names and `{{1}}` variables to `broadcast`.

## Files

| File | Purpose |
|---|---|
| `whatsflow.py` | Webhook verify + signature check, parser, bot engine, payload builders, broadcaster, CLI |
| `test_whatsflow.py` | 24 unit/integration tests — all passing |
| `campaign_sample.csv` | Sample broadcast list (duplicates + one opted-out contact to demo filtering) |
| `README.md` | This file |

## Test result

`Ran 24 tests … OK` (Python 3.12, no third-party packages).

*Portfolio demo built from the gig brief — open for review, never connected
to a real client engagement.*
