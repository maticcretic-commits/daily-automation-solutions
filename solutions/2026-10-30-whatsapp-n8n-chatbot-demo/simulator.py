#!/usr/bin/env python3
"""WhatsApp + n8n chatbot demo — LOCAL SIMULATOR (no live APIs).

This simulator mirrors the reply logic in workflow.json using the same
keyword/intent rules. The "LLM" step here is a clearly-labelled MOCK
responder: fixed template replies chosen by keyword rules. No model is
called, no WhatsApp message is sent, and no network request is made.

All data is fictional: business "Demo Traders", fictional customers with
placeholder IDs (demo-customer-001, ...). No real phone numbers, tokens,
or credentials appear anywhere in this demo.

Usage:
    python3 simulator.py                # run all 5 sample payloads
    python3 simulator.py payloads/01-price.json   # run one payload file
"""

import json
import sys
from pathlib import Path

BUSINESS_NAME = "Demo Traders"

# Mock send endpoint. The ".invalid" domain can never resolve, so this
# URL can never reach a real service. It exists only to show where the
# live WhatsApp send step would sit in the n8n workflow.
MOCK_ENDPOINT = "https://mock-endpoint.invalid/demo-whatsapp/send"

PRICE_KEYWORDS = ("price", "cost", "how much", "rate", "pricing", "charge")
HOURS_KEYWORDS = ("hours", "open", "opening", "timing", "close", "closed", "when")
HUMAN_KEYWORDS = ("human", "agent", "person", "call me", "speak to",
                  "complaint", "representative")

REPLIES = {
    "price": (
        "Hello from Demo Traders (demo reply). Our standard pack is "
        "listed at Rs 100 in this demo. Reply 'PRICE LIST' for the full "
        "demo list, or ask for a human if you need help choosing."
    ),
    "hours": (
        "Hello from Demo Traders (demo reply). Our demo opening hours "
        "are Monday to Saturday, 9am to 7pm. We are closed on Sundays "
        "in this demo."
    ),
    "human": (
        "Thanks for your message. I have marked this for a human at "
        "Demo Traders. A person will reply during demo business hours. "
        "(Simulated handoff - no human was contacted.)"
    ),
}

FALLBACK_REPLY = (
    "Sorry, I did not understand that question in this demo. I have "
    "marked it for a human at Demo Traders - a person will reply "
    "during demo business hours. (Simulated handoff - no human was "
    "contacted.)"
)


def classify_intent(text):
    """Return (intent, understood). Unknown text falls back to 'human'."""
    lowered = text.lower()
    if any(word in lowered for word in HUMAN_KEYWORDS):
        return "human", True
    if any(word in lowered for word in PRICE_KEYWORDS):
        return "price", True
    if any(word in lowered for word in HOURS_KEYWORDS):
        return "hours", True
    return "human", False


def mock_responder(intent, understood):
    """MOCK responder. Template reply only — never a live model call."""
    if not understood:
        return FALLBACK_REPLY
    return REPLIES[intent]


def parse_payload(data):
    """Extract sender + message text from a WhatsApp-style webhook body.

    Raises ValueError (with a safe, generic message) for malformed input.
    The raw payload is never printed or echoed back.
    """
    if not isinstance(data, dict):
        raise ValueError("payload must be a JSON object")
    try:
        entry = data["entry"][0]
        change = entry["changes"][0]
        value = change["value"]
        message = value["messages"][0]
        sender = message["from"]
        body = message["text"]["body"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError("malformed payload: missing message fields") from exc
    if not isinstance(sender, str) or not sender:
        raise ValueError("malformed payload: missing sender")
    if not isinstance(body, str) or not body.strip():
        raise ValueError("malformed payload: missing message text")
    return {"from": sender, "body": body.strip()}


def process_payload(data):
    """Run one webhook payload through the demo reply logic.

    Returns a dict with ok=True and the outbound reply, or ok=False with
    a safe error message. Never raises on bad input.
    """
    try:
        parsed = parse_payload(data)
    except ValueError as exc:
        return {"ok": False, "error": str(exc)}
    intent, understood = classify_intent(parsed["body"])
    reply = mock_responder(intent, understood)
    return {
        "ok": True,
        "from": parsed["from"],
        "inbound": parsed["body"],
        "intent": intent,
        "understood": understood,
        "outbound": {
            "to": parsed["from"],
            "body": reply,
            "simulated": True,
            "responder": "MOCK",
        },
        "mock_endpoint": MOCK_ENDPOINT,
    }


def run_file(path):
    """Process one payload JSON file and print inbound + outbound."""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    result = process_payload(data)
    print("FILE: {}".format(Path(path).name))
    if not result["ok"]:
        print("  REJECTED: {}".format(result["error"]))
        return result
    print("  IN  [{}]: {}".format(result["from"], result["inbound"]))
    print("  INTENT: {} (understood: {})".format(
        result["intent"], result["understood"]))
    print("  OUT [{}] [MOCK responder, simulated]: {}".format(
        result["outbound"]["to"], result["outbound"]["body"]))
    print("  SEND: {} (mock endpoint, no request made)".format(
        result["mock_endpoint"]))
    return result


def main(argv):
    here = Path(__file__).resolve().parent
    if len(argv) > 1:
        paths = [Path(p) for p in argv[1:]]
    else:
        payloads_dir = here / "payloads"
        paths = sorted(payloads_dir.glob("*.json"))
    print("WhatsApp + n8n chatbot demo — SIMULATION ONLY")
    print("Business: {} (fictional). No live API is called.".format(
        BUSINESS_NAME))
    print()
    results = [run_file(p) for p in paths]
    print()
    ok_count = sum(1 for r in results if r["ok"])
    print("{}/{} payloads produced a reply ({} rejected safely).".format(
        ok_count, len(results), len(results) - ok_count))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
