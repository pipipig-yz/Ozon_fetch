"""
Test Ozon API key permissions via POST /v1/roles.

Usage:
    python test_roles.py
"""

import json
import os
import sys

import requests

BASE_URL = "https://api-seller.ozon.ru"
TIMEOUT = 15


def load_credentials() -> tuple[str, str]:
    """Read Client-Id and Api-Key from .env file."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, ".env")

    client_id = os.getenv("OZON_CLIENT_ID")
    api_key = os.getenv("OZON_API_KEY")

    if not client_id or not api_key:
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
        sys.exit("Missing credentials in .env file.")
    return client_id, api_key


def call_endpoint(client_id: str, api_key: str, path: str, body: dict | None = None) -> dict:
    """POST to an Ozon API endpoint, return (status, body)."""
    url = f"{BASE_URL}{path}"
    headers = {
        "Client-Id": client_id,
        "Api-Key": api_key,
        "Content-Type": "application/json",
    }
    resp = requests.post(url, headers=headers, json=body or {}, timeout=TIMEOUT)
    try:
        data = resp.json()
    except Exception:
        data = {"_raw": resp.text}
    return {"status": resp.status_code, "body": data}


def main():
    client_id, api_key = load_credentials()
    print(f"Client-Id : {client_id}")
    print(f"Api-Key   : {api_key[:8]}...{api_key[-4:]}")
    print()

    # 1) Check roles / permissions
    print("─" * 55)
    print("  [1] POST /v1/roles  —  what permissions does this key have?")
    r = call_endpoint(client_id, api_key, "/v1/roles")
    print(f"  HTTP {r['status']}")
    body = r["body"]
    if isinstance(body, dict):
        print(json.dumps(body, ensure_ascii=False, indent=4))
    else:
        print(body)

    if r["status"] == 200 and isinstance(body, dict):
        expires = body.get("expires_at", "N/A")
        roles = body.get("roles") or []
        all_methods: list[str] = []
        print(f"\n  Token expires: {expires}")
        print(f"  Roles ({len(roles)}):")
        for role in roles:
            name = role.get("name", "?")
            methods = role.get("methods") or []
            all_methods.extend(methods)
            print(f"    • {name}: {methods}")
        print(f"\n  Total accessible paths: {len(all_methods)}")
        # Highlight finance-related
        finance_paths = [m for m in all_methods if "finance" in m.lower()]
        if finance_paths:
            print(f"  ✅ Finance paths found: {finance_paths}")
        else:
            print(f"  ❌ No finance paths in roles — key lacks finance permission!")

    print()

    # 2) Test v1 cash-flow endpoint
    print("─" * 55)
    print("  [2] POST /v1/finance/cash-flow-statement/list")
    r = call_endpoint(client_id, api_key, "/v1/finance/cash-flow-statement/list", {
        "date": {"from": "2026-06-01T00:00:00.000Z", "to": "2026-06-30T23:59:59.999Z"},
        "page": 1,
        "page_size": 1,
    })
    print(f"  HTTP {r['status']}")
    body = r["body"]
    if isinstance(body, dict):
        if "result" in body:
            print("  ✅ SUCCESS — got cash_flows result!")
            print(json.dumps(body, ensure_ascii=False, indent=2)[:500])
        else:
            print(json.dumps(body, ensure_ascii=False, indent=2))
    else:
        print(body)

    print()

    # 3) Test v3 transaction endpoint (comparison)
    print("─" * 55)
    print("  [3] POST /v3/finance/transaction/list  (old, being deprecated)")
    r = call_endpoint(client_id, api_key, "/v3/finance/transaction/list", {
        "filter": {"date": {"from": "2026-06-01T00:00:00.000Z", "to": "2026-06-30T23:59:59.999Z"}},
        "page": 1,
        "page_size": 1,
    })
    print(f"  HTTP {r['status']}")
    body = r["body"]
    if isinstance(body, dict):
        if "result" in body:
            print("  ✅ SUCCESS — got operations result!")
            print(json.dumps(body, ensure_ascii=False, indent=2)[:500])
        else:
            print(json.dumps(body, ensure_ascii=False, indent=2))
    else:
        print(body)


if __name__ == "__main__":
    main()
