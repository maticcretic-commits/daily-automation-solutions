# LeadFlow — Lead Capture → Scoring → CRM Sync

A common client problem: *"Leads come in from our website form, but they sit in
a spreadsheet. Hot prospects wait days for a reply, duplicates pile up, and
nobody knows which leads are worth calling first."*

LeadFlow fixes that with a small webhook pipeline:

1. **Capture** — `POST /lead` accepts JSON from any website form or landing page
2. **Validate** — required fields, email/phone checks, consent flag, plus a
   **honeypot field** that silently rejects bots
3. **De-duplicate** — normalises emails (gmail `+tags`/dots collapse) and
   upserts into SQLite instead of creating duplicate contacts
4. **Score** — rule-based 0–100 fit score → `hot` / `warm` / `cold` tier
   (company email, budget, high-intent source, phone, message depth)
5. **Route** — hot leads push to the CRM *immediately*; warm/cold stay in the
   store for nurture/export

## CRM integrations (`crm.py`)

The pipeline talks to a tiny `CRMClient` interface — swapping CRMs is one line:

- `ConsoleCRMClient` — logs the sync (demo/test default)
- `WebhookCRMClient` — POSTs the contact JSON to any HTTPS endpoint
  (n8n / Make / Zapier webhook, or a CRM's inbound webhook)
- `HubSpotPayloadBuilder` — builds a HubSpot v3 *create/update contact*
  payload; add your own private-app token to actually call the API
  (token stays with you — never committed)

## Run it

No dependencies — Python 3.10+ standard library only.

```bash
# 1. run the webhook server
python leadflow.py serve --port 8000

# 2. submit a lead (new terminal)
curl -X POST http://127.0.0.1:8000/lead -H 'Content-Type: application/json' -d '{
  "name": "Priya Sharma",
  "email": "priya@acme-logistics.com",
  "phone": "+91 98765 43210",
  "company": "Acme Logistics",
  "source": "demo-request",
  "budget": 2500,
  "message": "We need lead automation for our 5-person sales team.",
  "consent": true
}'

# 3. inspect + report
curl http://127.0.0.1:8000/report
python leadflow.py report
python leadflow.py export --tier hot --out hot-leads.csv
```

To push hot leads to a real endpoint instead of the console:

```bash
python leadflow.py serve --crm-webhook https://your-n8n/webhook/lead-in
```

## Tests

```bash
python test_leadflow.py
```

Covers validation, honeypot bot rejection, email normalisation/dedup,
scoring tiers, hot→CRM routing, cold staying in nurture, and the HubSpot
payload builder.

## What a client would get next

- Swap `ConsoleCRMClient` for their real CRM (HubSpot/Salesforce/GoHighLevel)
  via the webhook client or a native API client on the same interface
- Add Slack/email alerts for hot leads, and a scheduled nurture export
- Front it with their actual website form (the honeypot field stays hidden)
