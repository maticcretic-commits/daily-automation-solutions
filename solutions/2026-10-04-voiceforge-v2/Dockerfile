FROM python:3.12-slim

WORKDIR /app
COPY voiceforge/ ./voiceforge/
COPY server_entry.py ./

# stdlib only — no pip install needed
EXPOSE 8080

# Required at runtime: PUBLIC_BASE_URL, TWILIO_AUTH_TOKEN.
# Recommended: TWILIO_ACCOUNT_SID, TWILIO_PHONE_NUMBER (SMS),
# VOICEFORGE_BUSINESS, VOICEFORGE_TRANSFER_NUMBER,
# LLM_API_URL + LLM_API_KEY (live LLM).
CMD ["python3", "server_entry.py"]
