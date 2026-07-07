"""
Ozon Seller API client — v1 finance accruals endpoints.

All POSTs to the API go through a retry wrapper with exponential backoff:
  - 429 (rate-limit) → waits for Retry-After, then retries
  - 5xx / ConnectionError / Timeout → exponential backoff (1s → 2s → 4s)
  - 4xx (non-429) → raised immediately, no retry
"""

import sys
import time
from datetime import datetime, timedelta

import requests

BASE_URL = "https://api-seller.ozon.ru"
TYPES_ENDPOINT = "/v1/finance/accrual/types"
BY_DAY_ENDPOINT = "/v1/finance/accrual/by-day"
TIMEOUT = 30
MAX_RETRIES = 3


# ---------------------------------------------------------------------------
# Low-level HTTP
# ---------------------------------------------------------------------------


def _post_with_retry(
    client_id: str,
    api_key: str,
    endpoint: str,
    body: dict,
    *,
    max_retries: int = MAX_RETRIES,
) -> dict:
    """POST to an Ozon API endpoint with retry on transient failures.

    Retry strategy:
        - 429 → honour ``Retry-After`` header (or 2ⁿ seconds fallback).
        - 5xx / ConnectionError / Timeout → exponential backoff: 1s, 2s, 4s.
        - Other 4xx → no retry — raise immediately.

    Args:
        client_id: Ozon Client-Id header value.
        api_key: Ozon Api-Key header value.
        endpoint: API path, e.g. ``/v1/finance/accrual/types``.
        body: JSON-serialisable request body.
        max_retries: Maximum number of retry attempts (default 3).

    Returns:
        Parsed JSON response body (``dict``).

    Raises:
        requests.HTTPError: On non-retryable HTTP errors or after exhausting
            all retries.
        requests.ConnectionError: After exhausting connection retries.
        requests.Timeout: After exhausting timeout retries.
    """
    url = f"{BASE_URL}{endpoint}"
    headers = {
        "Client-Id": client_id,
        "Api-Key": api_key,
        "Content-Type": "application/json",
    }

    last_exception: Exception | None = None

    for attempt in range(max_retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=body, timeout=TIMEOUT)

            # 429 — rate limited
            if resp.status_code == 429:
                if attempt < max_retries:
                    retry_after = resp.headers.get("Retry-After")
                    wait = int(retry_after) if retry_after and retry_after.isdigit() else 2 ** attempt
                    print(
                        f"  ⚠ Rate-limited (429), waiting {wait}s … "
                        f"(attempt {attempt + 1}/{max_retries + 1})",
                        file=sys.stderr,
                    )
                    time.sleep(wait)
                    continue
                resp.raise_for_status()  # exhausted — let it surface

            # 5xx — server-side transient
            if resp.status_code >= 500 and attempt < max_retries:
                wait = 2 ** attempt
                print(
                    f"  ⚠ Server error ({resp.status_code}), retrying in {wait}s … "
                    f"(attempt {attempt + 1}/{max_retries + 1})",
                    file=sys.stderr,
                )
                time.sleep(wait)
                continue

            resp.raise_for_status()
            return resp.json()

        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
            last_exception = exc
            if attempt < max_retries:
                wait = 2 ** attempt
                print(
                    f"  ⚠ Network error ({exc.__class__.__name__}), retrying in {wait}s … "
                    f"(attempt {attempt + 1}/{max_retries + 1})",
                    file=sys.stderr,
                )
                time.sleep(wait)
                continue
            raise

        except requests.exceptions.HTTPError:
            # Non-retryable 4xx (excluding 429, handled above)
            raise

    # Should be unreachable — all retries exhausted in a branch above
    if last_exception:
        raise last_exception
    raise RuntimeError("_post_with_retry: unexpected retry exhaustion")


# ---------------------------------------------------------------------------
# Accrual types
# ---------------------------------------------------------------------------


def fetch_accrual_types(client_id: str, api_key: str) -> dict[int, dict]:
    """Return ``{type_id: {name, description}}`` mapping for all fee types."""
    data = _post_with_retry(client_id, api_key, TYPES_ENDPOINT, {})
    types_list = data.get("accrual_types")
    if types_list is None:
        result = data.get("result")
        types_list = result.get("accrual_types", []) if isinstance(result, dict) else []
    mapping: dict[int, dict] = {}
    for t in types_list:
        tid = t.get("id")
        if tid is not None:
            mapping[tid] = {"name": t.get("name", ""), "description": t.get("description", "")}
    return mapping


# ---------------------------------------------------------------------------
# Accruals by day (cursor-paginated)
# ---------------------------------------------------------------------------


def fetch_accruals_for_date(
    client_id: str,
    api_key: str,
    date_str: str,
    *,
    max_pages: int = 100,
) -> list[dict]:
    """Fetch all accruals for a single date, following the ``last_id`` cursor.

    Args:
        max_pages: Safety limit to prevent infinite loops (from ``--max-pages``).

    Returns:
        List of accrual dicts for *date_str*.
    """
    all_accruals: list[dict] = []
    last_id: str = ""
    pages = 0

    while pages < max_pages:
        body: dict = {"date": date_str}
        if last_id:
            body["last_id"] = last_id

        data = _post_with_retry(client_id, api_key, BY_DAY_ENDPOINT, body)
        result = data.get("result", data)
        accruals = result.get("accruals", [])
        if not accruals:
            break

        all_accruals.extend(accruals)
        pages += 1
        last_id = result.get("last_id", "")
        if not last_id:
            break

    return all_accruals


def fetch_all_accruals(
    client_id: str,
    api_key: str,
    date_from: str,
    date_to: str,
    *,
    max_pages: int = 100,
) -> list[dict]:
    """Fetch accruals for a date range, one day at a time.

    Args:
        max_pages: Max cursor-pages per day (passed through to
            ``fetch_accruals_for_date``).
    """
    all_accruals: list[dict] = []

    start = datetime.fromisoformat(date_from).date()
    end = datetime.fromisoformat(date_to).date()
    total_days = (end - start).days + 1

    for i in range(total_days):
        day = (start + timedelta(days=i)).isoformat()
        print(f"Fetching {day} ({i + 1}/{total_days}) …", file=sys.stderr)
        day_accruals = fetch_accruals_for_date(
            client_id, api_key, day, max_pages=max_pages,
        )
        all_accruals.extend(day_accruals)
        print(f"  → {len(day_accruals)} accruals", file=sys.stderr)

    return all_accruals
