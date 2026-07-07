"""
Fetch Ozon finance accruals via v1 LTS endpoints.

Endpoints used (all v1, long-term stable):
    POST /v1/finance/accrual/types     —  type_id → name mapping
    POST /v1/finance/accrual/by-day    —  accrued charges per day (cursor pagination)

API docs: https://docs.ozon.ru/api/seller/zh/

Credentials are read from environment variables:
    OZON_CLIENT_ID  — your Ozon seller Client-Id
    OZON_API_KEY    — your Ozon seller Api-Key

Usage:
    python fetch_transactions.py                          # last 30 days → report.xlsx
    python fetch_transactions.py --days 60                # last 60 days
    python fetch_transactions.py --from 2026-05-01        # from May 1 to today
    python fetch_transactions.py --from 2026-05-01 --to 2026-07-05
    python fetch_transactions.py --output report.xlsx     # explicit output path
    python fetch_transactions.py --json result.json       # also save raw JSON
"""

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone

from config import load_credentials
from api_client import (
    fetch_accrual_types,
    fetch_all_accruals,
    TYPES_ENDPOINT,
    BY_DAY_ENDPOINT,
)
from transformations import (
    amt,
    build_focused_rows,
    group_by_base_posting_number,
    FIELDS,
)
from excel_writer import write_excel


# ===========================================================================
# Debug helpers
# ===========================================================================


def _print_fee_summary(
    accruals: list[dict], types_map: dict[int, dict]
) -> None:
    """Print a summary of all fees found in the data (count × total per type).

    Useful for verifying that keyword-based fee matching in
    ``build_focused_rows`` will actually catch the expected fees.
    """
    fee_summary: dict[str, tuple[int, float, str]] = {}
    for a in accruals:
        item_fees = a.get("item_fees") or {}
        for sg in item_fees.get("fees") or []:
            for fee in sg.get("fees") or []:
                tid = fee.get("type_id", 0)
                tname = types_map.get(tid, {}).get("name", f"UNKNOWN_{tid}")
                prev = fee_summary.get(tname, (0, 0.0, "item_fee"))
                fee_summary[tname] = (prev[0] + 1, prev[1] + amt(fee.get("accrued")), "item_fee")
        nif = a.get("non_item_fee") or {}
        if nif:
            tid = nif.get("type_id", 0)
            tname = types_map.get(tid, {}).get("name", f"UNKNOWN_{tid}")
            prev = fee_summary.get(tname, (0, 0.0, "non_item_fee"))
            fee_summary[tname] = (prev[0] + 1, prev[1] + amt(nif.get("accrued")), "non_item_fee")
        posting = a.get("posting") or {}
        for prod in posting.get("products") or []:
            for svc in (prod.get("delivery") or {}).get("services") or []:
                tid = svc.get("type_id", 0)
                tname = types_map.get(tid, {}).get("name", f"UNKNOWN_{tid}")
                prev = fee_summary.get(tname, (0, 0.0, "delivery_svc"))
                fee_summary[tname] = (prev[0] + 1, prev[1] + amt(svc.get("accrued")), "delivery_svc")

    if not fee_summary:
        print("  (no fees found in any accrual)", file=sys.stderr)
        return

    print("  Fees found in data:", file=sys.stderr)
    for tname, (cnt, total, src) in sorted(
        fee_summary.items(), key=lambda x: -abs(x[1][1])
    ):
        print(
            f"    {tname:35s} ×{cnt:>3}  total={total:>12.2f}  [{src}]",
            file=sys.stderr,
        )


# ===========================================================================
# Terminal summary
# ===========================================================================


def summarize(accruals: list[dict], types_map: dict[int, dict]) -> None:
    """Print a high-level summary of accruals to stdout."""
    if not accruals:
        print("No accruals found.")
        return

    total_amount = 0.0
    cat_counts: dict[str, int] = {}
    for a in accruals:
        total_amount += amt(a.get("total_amount"))
        cat = a.get("accrued_category", "UNKNOWN")
        cat_counts[cat] = cat_counts.get(cat, 0) + 1

    print(f"\n{'=' * 55}")
    print(f"  应计记录数: {len(accruals)}")
    print(f"  总金额:     {total_amount:>12,.2f}")
    print(f"\n  按类别统计:")
    for cat, cnt in sorted(cat_counts.items(), key=lambda x: -x[1]):
        print(f"    {cat:30s} {cnt:>6d}")
    print(f"  类型定义数: {len(types_map)}")
    print(f"{'=' * 55}")


# ===========================================================================
# Terminal summary — orders grouped by base posting number
# ===========================================================================


def print_order_summary(
    all_focused_rows: list[dict],
    order_groups: dict[str, dict],
) -> None:
    """Print a focused per-order summary to the terminal, most recent first.

    Accepts pre-computed data so the caller can share the same
    ``all_focused_rows`` and ``order_groups`` with ``write_excel``.
    """
    if not all_focused_rows:
        print("No data to summarize.")
        return

    # Sort by latest date descending (most recent first)
    def _sort_key(item: tuple[str, dict]) -> str:
        dates = item[1]["dates"]
        return max(dates) if dates else ""

    sorted_orders = sorted(order_groups.items(), key=_sort_key, reverse=True)

    # Print table
    W_PN = 24
    W_DATE = 12
    W_SKU = 20
    W_AMT = 12
    HEADERS = ["基础单号", "日期", "SKU", "收入", "收单", "销售佣金", "配送佣金", "国际配送"]

    sep = "─" * (W_PN + W_DATE + W_SKU + W_AMT * 5 + 7 * 3 + 2)
    print(f"\n{sep}")
    print(
        f"  {HEADERS[0]:{W_PN}} │ {HEADERS[1]:{W_DATE}} │ {HEADERS[2]:{W_SKU}} │"
        f" {HEADERS[3]:>{W_AMT}} │ {HEADERS[4]:>{W_AMT}} │ {HEADERS[5]:>{W_AMT}} │"
        f" {HEADERS[6]:>{W_AMT}} │ {HEADERS[7]:>{W_AMT}}"
    )
    print(sep)

    total_row = {"收入": 0.0, "收单": 0.0, "销售佣金": 0.0, "配送佣金": 0.0, "国际配送": 0.0}

    for base, g in sorted_orders:
        dates = g["dates"]
        if len(dates) == 1:
            date_str = next(iter(dates))
        else:
            sd = sorted(dates)
            date_str = f"{sd[0]}~{sd[-1]}"
        sku_str = ", ".join(sorted(g["skus"])) if g["skus"] else "-"
        if len(sku_str) > W_SKU:
            sku_str = sku_str[: W_SKU - 2] + "…"
        pn_str = base if len(base) <= W_PN else base[: W_PN - 2] + "…"
        amts = [g[f] for f in FIELDS]

        print(
            f"  {pn_str:{W_PN}} │ {date_str:{W_DATE}} │ {sku_str:{W_SKU}} │"
            f" {amts[0]:{W_AMT},.2f} │ {amts[1]:{W_AMT},.2f} │ {amts[2]:{W_AMT},.2f} │"
            f" {amts[3]:{W_AMT},.2f} │ {amts[4]:{W_AMT},.2f}"
        )
        for i, f in enumerate(FIELDS):
            total_row[f] += g[f]

    print(sep)
    print(
        f"  {'合计':{W_PN}} │ {'':{W_DATE}} │ {'':{W_SKU}} │"
        f" {total_row['收入']:{W_AMT},.2f} │ {total_row['收单']:{W_AMT},.2f} │"
        f" {total_row['销售佣金']:{W_AMT},.2f} │ {total_row['配送佣金']:{W_AMT},.2f} │"
        f" {total_row['国际配送']:{W_AMT},.2f}"
    )
    print(sep)
    print(f"  {len(sorted_orders)} orders, {len(all_focused_rows)} line items")


# ===========================================================================
# CLI
# ===========================================================================


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch Ozon finance accruals via /v1/finance/accrual/by-day (LTS)"
    )
    parser.add_argument(
        "--from", dest="date_from",
        help="Start date (YYYY-MM-DD).  If set, --days is ignored.",
    )
    parser.add_argument(
        "--to", dest="date_to",
        help="End date (YYYY-MM-DD).  Default: today.  Only meaningful with --from.",
    )
    parser.add_argument(
        "--days", type=int, default=30,
        help="Days back from today (default 30).  Ignored when --from is given.",
    )
    parser.add_argument(
        "--output", "-o", default="report.xlsx",
        help="Excel output path. Default: report.xlsx.",
    )
    parser.add_argument(
        "--json",
        help="Also dump raw JSON to this file.",
    )
    parser.add_argument(
        "--summary", "-s", action="store_true",
        help="Print per-order summary table to terminal.",
    )
    parser.add_argument(
        "--max-pages", type=int, default=100,
        help="Max cursor pages per day. Default: 100.",
    )

    args = parser.parse_args()

    # --- Validate arguments ---
    if args.date_to and not args.date_from:
        parser.error("--to requires --from.")

    # --- Date range ---
    today = datetime.now(timezone.utc).date()
    if args.date_from:
        date_from: str = args.date_from
        date_to: str = args.date_to or today.isoformat()
    else:
        date_from = (today - timedelta(days=args.days)).isoformat()
        date_to = today.isoformat()

    # --- Auth ---
    client_id, api_key = load_credentials()

    print(f"Date range : {date_from} → {date_to}", file=sys.stderr)
    print(
        f"Endpoints  : {TYPES_ENDPOINT}  +  "
        f"{BY_DAY_ENDPOINT}  (v1 LTS)",
        file=sys.stderr,
    )

    # --- 1. Fetch accrual types ---
    types_map = fetch_accrual_types(client_id, api_key)

    # --- 2. Fetch accruals day by day ---
    accruals = fetch_all_accruals(
        client_id, api_key, date_from, date_to, max_pages=args.max_pages,
    )

    # --- 3. Debug: show all fees present in the data ---
    _print_fee_summary(accruals, types_map)

    # --- 4. Pre-compute transformed data (shared by Excel & terminal) ---
    all_focused_rows: list[dict] = []
    for a in accruals:
        for fr in build_focused_rows(a, types_map):
            all_focused_rows.append(fr)

    order_groups = group_by_base_posting_number(all_focused_rows)

    # --- 5. Always Excel ---
    write_excel(accruals, types_map, all_focused_rows, order_groups, args.output)

    # --- 6. Optional terminal summary ---
    if args.summary:
        print_order_summary(all_focused_rows, order_groups)

    # --- 7. Optional JSON ---
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(
                {"accruals": accruals, "types": types_map},
                f, ensure_ascii=False, indent=2, default=str,
            )
        print(f"Raw JSON saved to {args.json}", file=sys.stderr)

    # --- 8. Summary ---
    summarize(accruals, types_map)


if __name__ == "__main__":
    main()
