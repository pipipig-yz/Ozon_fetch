"""
Fetch Ozon transaction data via /v3/finance/transaction/list.

API docs: https://docs.ozon.ru/api/seller/#/operations/transactions

Credentials are read from environment variables:
    OZON_CLIENT_ID  — your Ozon seller Client-Id
    OZON_API_KEY    — your Ozon seller Api-Key

Usage:
    python fetch_transactions.py                          # last 7 days, print summary
    python fetch_transactions.py --days 30                # last 30 days
    python fetch_transactions.py --from 2025-06-01 --to 2025-06-25
    python fetch_transactions.py --output transactions.json
    python fetch_transactions.py --excel report.xlsx      # output as Excel file
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from typing import Optional

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

BASE_URL = "https://api-seller.ozon.ru"
ENDPOINT = "/v3/finance/transaction/list"
TIMEOUT = 30


def get_client() -> tuple[str, str]:
    """Read Client-Id and Api-Key from environment."""
    client_id = os.getenv("OZON_CLIENT_ID")
    api_key = os.getenv("OZON_API_KEY")

    if not client_id or not api_key:
        # Also try loading from a .env file in the script directory
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


def build_payload(
    date_from: str,
    date_to: str,
    page: int = 1,
    page_size: int = 1000,
    transaction_type: Optional[str] = None,
) -> dict:
    """Build the POST body for /v3/finance/transaction/list."""
    payload: dict = {
        "filter": {
            "date": {
                "from": f"{date_from}T00:00:00.000Z",
                "to": f"{date_to}T23:59:59.999Z",
            },
        },
        "page": page,
        "page_size": page_size,
    }

    if transaction_type:
        payload["filter"]["transaction_type"] = transaction_type

    return payload


def fetch_page(
    client_id: str,
    api_key: str,
    payload: dict,
) -> dict:
    """POST one page of transactions. Returns the JSON response."""
    url = f"{BASE_URL}{ENDPOINT}"
    headers = {
        "Client-Id": client_id,
        "Api-Key": api_key,
        "Content-Type": "application/json",
    }

    resp = requests.post(url, headers=headers, json=payload, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def fetch_all_transactions(
    client_id: str,
    api_key: str,
    date_from: str,
    date_to: str,
    transaction_type: Optional[str] = None,
    max_pages: int = 100,
) -> list[dict]:
    """Paginate through all transaction results and return a flat list."""
    all_operations: list[dict] = []
    page = 1

    while page <= max_pages:
        payload = build_payload(
            date_from=date_from,
            date_to=date_to,
            page=page,
            transaction_type=transaction_type,
        )

        print(f"Fetching page {page} …", file=sys.stderr)
        data = fetch_page(client_id, api_key, payload)

        result = data.get("result", {})
        ops = result.get("operations", [])
        if not ops:
            break

        all_operations.extend(ops)

        total = result.get("total_count", 0) or result.get("totalCount", 0)
        if total and len(all_operations) >= total:
            break

        page += 1

    return all_operations


def summarize(ops: list[dict]) -> None:
    """Print a quick summary of the fetched transactions."""
    if not ops:
        print("No transactions found.")
        return

    # Count by type
    type_counts: dict[str, int] = {}
    # Sum amounts (the field is usually "amount" in kopecks)
    total_amount = 0.0
    total_accruals = 0.0  # positive

    for op in ops:
        t = op.get("operation_type") or op.get("operationType", "unknown")
        type_counts[t] = type_counts.get(t, 0) + 1

        amt = op.get("amount", 0)
        total_amount += amt
        if amt > 0:
            total_accruals += amt

    print(f"\n{'='*50}")
    print(f"Total transactions: {len(ops)}")
    print(f"Sum of all amounts : {total_amount:,.2f} ₽")
    print(f"Sum of accruals    : {total_accruals:,.2f} ₽")
    print(f"\nBreakdown by type:")
    for t, c in sorted(type_counts.items(), key=lambda x: -x[1]):
        print(f"  {t:40s} {c:>6d}")
    print(f"{'='*50}")


def write_excel(ops: list[dict], filepath: str) -> None:
    """Write transactions to a formatted Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Ozon Transactions"

    # Define headers (bilingual)
    headers = [
        ("operation_id", "操作ID"),
        ("operation_type", "操作类型(代码)"),
        ("operation_type_name", "操作类型名称"),
        ("operation_date", "操作日期"),
        ("type", "交易大类"),
        ("amount", "最终金额(₽)"),
        ("accruals_for_sale", "销售收入(₽)"),
        ("sale_commission", "销售佣金(₽)"),
        ("delivery_charge", "配送费(₽)"),
        ("return_delivery_charge", "退货配送费(₽)"),
        ("delivery_schema", "配送模式"),
        ("order_date", "下单日期"),
        ("posting_number", "发货单号"),
        ("warehouse_id", "仓库ID"),
        ("item_names", "商品名称"),
        ("item_skus", "商品SKU"),
        ("service_names", "服务项"),
        ("service_prices", "服务项金额(₽)"),
    ]

    # Style definitions
    header_font = Font(name="Microsoft YaHei", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_alignment = Alignment(vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Write header row
    for col_idx, (_, header_name) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Write data rows
    for row_idx, op in enumerate(ops, 2):
        posting = op.get("posting") or {}
        items = op.get("items") or []
        services = op.get("services") or []

        # Flatten nested fields
        item_names = ", ".join(it.get("name", "") for it in items)
        item_skus = ", ".join(str(it.get("sku", "")) for it in items)
        service_names = ", ".join(sv.get("name", "") for sv in services)
        service_prices = ", ".join(str(sv.get("price", "")) for sv in services)

        row_data = [
            op.get("operation_id", ""),
            op.get("operation_type", ""),
            op.get("operation_type_name", ""),
            op.get("operation_date", ""),
            op.get("type", ""),
            op.get("amount", 0),
            op.get("accruals_for_sale", 0),
            op.get("sale_commission", 0),
            op.get("delivery_charge", 0),
            op.get("return_delivery_charge", 0),
            posting.get("delivery_schema", ""),
            posting.get("order_date", ""),
            posting.get("posting_number", ""),
            posting.get("warehouse_id", ""),
            item_names,
            item_skus,
            service_names,
            service_prices,
        ]

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = cell_alignment
            cell.border = thin_border

    # Adjust column widths (approximate)
    col_widths = {
        1: 15, 2: 38, 3: 50, 4: 14, 5: 10, 6: 14,
        7: 16, 8: 14, 9: 12, 10: 14, 11: 14, 12: 14,
        13: 22, 14: 12, 15: 50, 16: 16, 17: 35, 18: 16,
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    # Freeze the header row
    ws.freeze_panes = "A2"

    # Auto-filter on all columns
    ws.auto_filter.ref = ws.dimensions

    wb.save(filepath)
    print(f"Saved {len(ops)} transactions to {filepath}", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch Ozon finance transactions via /v3/finance/transaction/list"
    )
    parser.add_argument(
        "--from", dest="date_from",
        help="Start date (YYYY-MM-DD). Default: 7 days ago.",
    )
    parser.add_argument(
        "--to", dest="date_to",
        help="End date (YYYY-MM-DD). Default: today.",
    )
    parser.add_argument(
        "--days", type=int, default=7,
        help="Number of days back from today (used when --from is not set). Default: 7.",
    )
    parser.add_argument(
        "--type", dest="transaction_type",
        help="Filter by transaction type (e.g. all, orders, returns, services).",
    )
    parser.add_argument(
        "--output", "-o",
        help="Save raw JSON to this file instead of printing a summary.",
    )
    parser.add_argument(
        "--excel", "-x",
        help="Save as Excel (.xlsx) file instead of JSON.",
    )
    parser.add_argument(
        "--max-pages", type=int, default=100,
        help="Max pages to fetch. Default: 100.",
    )

    args = parser.parse_args()

    # Determine date range
    today = datetime.now(timezone.utc).date()

    date_to: str = args.date_to or today.isoformat()
    if args.date_from:
        date_from = args.date_from
    else:
        date_from = (today - timedelta(days=args.days)).isoformat()

    # Auth
    client_id, api_key = get_client()

    print(f"Date range: {date_from} → {date_to}", file=sys.stderr)

    # Fetch
    ops = fetch_all_transactions(
        client_id=client_id,
        api_key=api_key,
        date_from=date_from,
        date_to=date_to,
        transaction_type=args.transaction_type,
        max_pages=args.max_pages,
    )

    # Output
    if args.excel:
        write_excel(ops, args.excel)
    elif args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(ops, f, ensure_ascii=False, indent=2)
        print(f"Saved {len(ops)} transactions to {args.output}", file=sys.stderr)
    else:
        summarize(ops)


if __name__ == "__main__":
    main()
