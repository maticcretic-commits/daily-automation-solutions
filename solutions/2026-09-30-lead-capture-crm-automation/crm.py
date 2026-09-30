"""
CRM integrations for LeadFlow.

The pipeline talks to a small CRMClient interface, so swapping CRMs is a
one-line change. Ships with:

  * ConsoleCRMClient  — logs the payload (great for demos/tests).
  * WebhookCRMClient  — POSTs the contact JSON to any HTTPS endpoint
                        (n8n/Make/Zapier webhook, or a CRM's own webhook).
  * HubSpotPayloadBuilder — builds a HubSpot v3 "create/update contact"
                        payload from a lead (no network; plug your token in).
"""

import json
import urllib.request


class CRMClient:
    def create_or_update_contact(self, lead) -> dict:
        raise NotImplementedError


class ConsoleCRMClient(CRMClient):
    """Pretend CRM: prints what would be synced. Perfect for demos."""

    def __init__(self):
        self.synced = []

    def create_or_update_contact(self, lead) -> dict:
        contact = {
            "email": lead.email,
            "firstname": lead.name.split()[0] if lead.name else "",
            "lastname": " ".join(lead.name.split()[1:]) if lead.name else "",
            "phone": lead.phone,
            "company": lead.company,
            "lead_source": lead.source,
            "lead_score": lead.score,
            "lead_tier": lead.tier,
            "budget": lead.budget,
        }
        self.synced.append(contact)
        print(f"[CRM] synced hot lead -> {lead.email} (score {lead.score})")
        return contact


class WebhookCRMClient(CRMClient):
    """Push the contact JSON to any HTTPS webhook endpoint."""

    def __init__(self, url: str, timeout: int = 10):
        if not url.startswith("https://"):
            raise ValueError("webhook URL must be https://")
        self.url = url
        self.timeout = timeout

    def create_or_update_contact(self, lead) -> dict:
        contact = ConsoleCRMClient().create_or_update_contact(lead)  # reuse shape
        req = urllib.request.Request(
            self.url,
            data=json.dumps(contact).encode(),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=self.timeout) as resp:
            return {"status": resp.status, "contact": contact}


class HubSpotPayloadBuilder:
    """Build a HubSpot v3 CRM payload from a lead.

    POST it to https://api.hubapi.com/crm/v3/objects/contacts
    with header  Authorization: Bearer <your-private-app-token>
    (token is yours to keep — never commit it).
    """

    @staticmethod
    def build(lead) -> dict:
        name = (lead.name or "").split()
        return {
            "properties": {
                "email": lead.email,
                "firstname": name[0] if name else "",
                "lastname": " ".join(name[1:]) if name else "",
                "phone": lead.phone,
                "company": lead.company,
                "hs_lead_status": "NEW",
                "lead_source": lead.source or "website",
                "hubspotscore": lead.score,  # map to your score property
                "notes": (lead.message or "")[:500],
            }
        }
