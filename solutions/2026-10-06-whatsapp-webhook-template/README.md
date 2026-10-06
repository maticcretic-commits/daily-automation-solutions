# WhatsApp Cloud API Webhook Template

A minimal, tested webhook server for Meta's **official WhatsApp Business Cloud
API** — the reusable foundation for WhatsApp automation builds (RFQ bots, lead
capture, support flows). Reusable template / working demo, not client work.

## What it does

| Endpoint | Behaviour |
|---|---|
| `GET /webhook` | Meta's verification handshake: checks `hub.mode=subscribe` and that `hub.verify_token` matches your `VERIFY_TOKEN`, then echoes `hub.challenge`. Anything else → 403. |
| `POST /webhook` | Receives message/status events. Validates the `X-Hub-Signature-256` HMAC-SHA256 signature against your `APP_SECRET` — tampered or forged posts get 403. Logs one structured JSON line per message/status (`from`, `timestamp`, `type`, `message_id`) — never secrets or raw payloads. |

Configuration is **env vars only** — the server refuses to start without them:

```
VERIFY_TOKEN=your-invented-token   # must match what you enter in Meta's dashboard
APP_SECRET=your-meta-app-secret    # App Secret from the Meta app dashboard
PORT=5000                          # optional, default 5000
```

Run:

```bash
pip install -r requirements.txt
VERIFY_TOKEN=... APP_SECRET=... python webhook_server.py
```

Tests: `pytest` (11 tests — handshake accept/reject, signature valid/invalid/tampered/missing, malformed JSON, secret-free logging, fail-fast config).

## Test mode setup (free, no business verification)

1. Create a Meta Developer account at developers.facebook.com (your Facebook login).
2. Create an App → type **Business** → add the **WhatsApp** product.
3. Meta issues a **test phone number + temporary access token** immediately.
4. Add up to 5 of your own numbers as verified recipients — real send/receive works at once.
5. Configure the webhook: callback URL (this server, public HTTPS) + your `VERIFY_TOKEN`, subscribe to `messages`.

## Production notes (per client launch)

- **Temporary tokens die in 24 hours** — the classic silent breakage. Production must use a **permanent system-user token**, never the temp token.
- Business portfolio + **business verification** required (document review; days to weeks).
- Real phone number (Meta OTPs it for the WABA lifetime).
- **Display-name approval: 2–5 business days** (must include the registered business name, no marketing words).
- Webhook endpoint must be public HTTPS; signature validation (already in this template) is mandatory.
- **Pricing (Oct 1, 2026):** 1,000 free service messages/month per business number, then per-message at market rates — India ₹0.115/message. **Oman rate unverified — do not quote; check Meta's live rate card at go-live.**
- Without a payment method on file, replies **silently stop** after the free tier.
- 24-hour customer-service window: outside it, only pre-approved **template** messages. Design flows so the user's inbound message opens the window.
