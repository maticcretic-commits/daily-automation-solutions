"""Entry point for `python3 server_entry.py` and the Dockerfile.

Keeps `python3 -m voiceforge.server` working too (see voiceforge/server.py).
"""

from voiceforge.server import main

if __name__ == "__main__":
    main()
