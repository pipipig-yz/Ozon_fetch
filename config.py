"""
Shared credential loading for Ozon API.

Reads OZON_CLIENT_ID and OZON_API_KEY from environment variables
or a local .env file next to the calling script.
"""

import os
import sys


def load_credentials() -> tuple[str, str]:
    """Read Client-Id and Api-Key from environment or a local .env file.

    Returns:
        (client_id, api_key) tuple.

    Exits with an error message if either credential is missing.
    """
    client_id = os.getenv("OZON_CLIENT_ID")
    api_key = os.getenv("OZON_API_KEY")

    if not client_id or not api_key:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(script_dir, ".env")
        if os.path.isfile(env_path):
            with open(env_path) as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k, v = k.strip(), v.strip().strip('"').strip("'")
                    if k == "OZON_CLIENT_ID" and not client_id:
                        client_id = v
                    elif k == "OZON_API_KEY" and not api_key:
                        api_key = v

    if not client_id or not api_key:
        sys.exit(
            "Missing credentials. Set OZON_CLIENT_ID and OZON_API_KEY env vars,\n"
            "or create a .env file next to this script with:\n"
            "    OZON_CLIENT_ID=your_client_id\n"
            "    OZON_API_KEY=your_api_key\n"
        )
    return client_id, api_key
