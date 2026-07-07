"""
Pure data-transformation functions for Ozon accrual data.

All functions are side-effect-free: they take dicts/lists and return
new dicts/lists.  No network, no filesystem, no global state.
"""

import re

# Top-level columns shared by the "focused" views.
FIELDS = ["收入", "收单", "销售佣金", "配送佣金", "国际配送"]

# Fee-type keyword groups used by build_focused_rows.
# When Ozon adds a new fee type that belongs to one of these categories,
# add its name (lowercase) to the appropriate list.
ACQUIRING_KEYWORDS = ["acquiring"]
DOMESTIC_DELIVERY_KEYWORDS = [
    "lastmile", "logistic", "rfbsdomesticdelivery",
    "rfbsdomesticagentfee", "rfbsglobalagentfee",
    "shipment", "fulfillment",
]
INTERNATIONAL_DELIVERY_KEYWORDS = [
    "rfbsglobaldelivery", "internationallogistic",
    "ozongloballogistics",
]


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


def amt(d: dict, default: float = 0.0) -> float:
    """Extract a numeric amount from an Ozon money object ``{amount, currency}``."""
    if not isinstance(d, dict):
        return default
    try:
        return float(d.get("amount", default))
    except (ValueError, TypeError):
        return default


def _get_commission_val(accrual: dict, sku, field: str) -> float:
    """Get a commission *field* value for a specific SKU within an accrual."""
    for prod in ((accrual.get("posting") or {}).get("products") or []):
        if prod.get("sku") == sku:
            return amt((prod.get("commission") or {}).get(field))
    return 0.0


# ---------------------------------------------------------------------------
# Fee matching
# ---------------------------------------------------------------------------


def _match_fee(
    accrual: dict,
    sku,
    types_map: dict[int, dict],
    keywords: list[str],
    *,
    include_non_item: bool = True,
) -> float:
    """Sum fees whose type name matches any *keyword* (case-insensitive).

    Searches across three locations within the accrual:
    1. ``item_fees`` — per-SKU fees (filtered to *sku*).
    2. ``non_item_fee`` — counted once, only when *include_non_item*.
    3. ``posting.products[].delivery.services`` — delivery-level fees.

    The *include_non_item* flag exists so that when iterating multiple
    SKUs for the same accrual, the non-item fee is only added to the
    *first* SKU row.
    """
    total = 0.0
    kw_lower = [k.lower() for k in keywords]

    def _hit(tid) -> bool:
        return any(
            kw in types_map.get(tid, {}).get("name", "").lower() for kw in kw_lower
        )

    # 1) item_fees (per-SKU)
    for sg in ((accrual.get("item_fees") or {}).get("fees") or []):
        if (sg.get("sku") or 0) != sku:
            continue
        for fee in sg.get("fees") or []:
            if _hit(fee.get("type_id", 0)):
                total += amt(fee.get("accrued"))

    # 2) non_item_fee — not per-SKU; only count once across all SKU rows
    if include_non_item:
        nif = accrual.get("non_item_fee") or {}
        if nif and _hit(nif.get("type_id", 0)):
            total += amt(nif.get("accrued"))

    # 3) delivery.services (per-SKU)
    for prod in ((accrual.get("posting") or {}).get("products") or []):
        if (prod.get("sku") or 0) != sku:
            continue
        for svc in ((prod.get("delivery") or {}).get("services") or []):
            if _hit(svc.get("type_id", 0)):
                total += amt(svc.get("accrued"))

    return total


# ---------------------------------------------------------------------------
# Row builders — flatten nested accrual structures
# ---------------------------------------------------------------------------


def extract_fee_rows(accrual: dict, types_map: dict[int, dict]) -> list[dict]:
    """Expand ``item_fees`` + ``non_item_fee`` + ``delivery.services`` into flat rows.

    One row per individual fee charge.  The ``source`` field records
    where the fee came from: ``"item_fees"``, ``"non_item_fee"``, or
    ``"delivery_service"``.
    """
    rows: list[dict] = []
    base = {
        "accrual_id": accrual.get("accrual_id", ""),
        "date": accrual.get("date", ""),
        "posting_number": accrual.get("unit_number", ""),
    }

    # 1) item_fees → per-SKU fees
    item_fees = accrual.get("item_fees") or {}
    for sku_group in item_fees.get("fees") or []:
        sku = sku_group.get("sku", "")
        for fee in sku_group.get("fees") or []:
            tid = fee.get("type_id", "")
            rows.append({
                **base,
                "sku": sku,
                "type_id": tid,
                "type_name": types_map.get(tid, {}).get("name", f"type_{tid}"),
                "amount": amt(fee.get("accrued")),
                "currency": (fee.get("accrued") or {}).get("currency", ""),
                "source": "item_fees",
            })

    # 2) non_item_fee
    nif = accrual.get("non_item_fee") or {}
    if nif:
        tid = nif.get("type_id", "")
        rows.append({
            **base,
            "sku": "-",
            "type_id": tid,
            "type_name": types_map.get(tid, {}).get("name", f"type_{tid}"),
            "amount": amt(nif.get("accrued")),
            "currency": (nif.get("accrued") or {}).get("currency", ""),
            "source": "non_item_fee",
        })

    # 3) delivery.services per product
    posting = accrual.get("posting") or {}
    for prod in posting.get("products") or []:
        psku = prod.get("sku", "")
        delivery = prod.get("delivery") or {}
        for svc in delivery.get("services") or []:
            tid = svc.get("type_id", "")
            rows.append({
                **base,
                "sku": psku,
                "type_id": tid,
                "type_name": types_map.get(tid, {}).get("name", f"type_{tid}"),
                "amount": amt(svc.get("accrued")),
                "currency": (svc.get("accrued") or {}).get("currency", ""),
                "source": "delivery_service",
            })

    return rows


def extract_product_rows(accrual: dict) -> list[dict]:
    """Flatten one accrual's ``posting.products`` into a list of dict rows."""
    rows: list[dict] = []
    base = {
        "accrual_id": accrual.get("accrual_id", ""),
        "date": accrual.get("date", ""),
        "posting_number": accrual.get("unit_number", ""),
    }
    posting = accrual.get("posting") or {}
    products = posting.get("products") or []
    for prod in products:
        sku = prod.get("sku", "")
        commission = prod.get("commission") or {}
        delivery = prod.get("delivery") or {}
        rows.append({
            **base,
            "sku": sku,
            "seller_price": amt(commission.get("seller_price")),
            "sale_amount": amt(commission.get("sale_amount")),
            "commission": amt(commission.get("commission")),
            "commission_ratio": commission.get("commission_ratio", ""),
            "sale_commission": amt(commission.get("sale_commission")),
            "bonus": amt(commission.get("bonus")),
            "coinvestment": amt(commission.get("coinvestment")),
            "delivery_total": amt(delivery.get("total_accrued")),
            "currency": (commission.get("sale_amount") or {}).get("currency", ""),
        })
    return rows


# ---------------------------------------------------------------------------
# Focused rows — one row per (accrual, SKU) for the summary views
# ---------------------------------------------------------------------------


def build_focused_rows(accrual: dict, types_map: dict[int, dict]) -> list[dict]:
    """Build one row per (accrual, SKU) with 收入/收单/销售佣金/配送佣金/国际配送."""
    rows: list[dict] = []
    base = {
        "accrual_id": accrual.get("accrual_id", ""),
        "date": accrual.get("date", ""),
        "posting_number": accrual.get("unit_number", ""),
    }

    # Collect all SKUs from products + item_fees
    skus: set = set()
    for prod in ((accrual.get("posting") or {}).get("products") or []):
        skus.add(prod.get("sku") or 0)
    for sg in ((accrual.get("item_fees") or {}).get("fees") or []):
        skus.add(sg.get("sku") or 0)
    skus.discard(0)

    if not skus:
        # No SKU-level data: one row with just non-item info
        rows.append({
            **base,
            "sku": "-",
            "收入": 0.0,
            "收单": _match_fee(accrual, 0, types_map, ACQUIRING_KEYWORDS),
            "销售佣金": 0.0,
            "配送佣金": _match_fee(accrual, 0, types_map, DOMESTIC_DELIVERY_KEYWORDS),
            "国际配送": _match_fee(accrual, 0, types_map, INTERNATIONAL_DELIVERY_KEYWORDS),
        })
    else:
        sorted_skus = sorted(skus)
        for i, sku in enumerate(sorted_skus):
            inc_nif = (i == 0)  # non_item_fee only on first SKU row
            rows.append({
                **base,
                "sku": sku,
                "收入": _get_commission_val(accrual, sku, "sale_amount"),
                "收单": _match_fee(accrual, sku, types_map, ACQUIRING_KEYWORDS, include_non_item=inc_nif),
                "销售佣金": _get_commission_val(accrual, sku, "sale_commission"),
                "配送佣金": _match_fee(accrual, sku, types_map, DOMESTIC_DELIVERY_KEYWORDS, include_non_item=inc_nif),
                "国际配送": _match_fee(accrual, sku, types_map, INTERNATIONAL_DELIVERY_KEYWORDS, include_non_item=inc_nif),
            })
    return rows


# ---------------------------------------------------------------------------
# Order-grouping — shared by Excel "按订单号排列" sheet and terminal output
# ---------------------------------------------------------------------------


def _base_posting_number(pn: str) -> str:
    """Strip trailing ``-N`` suffix from a posting number (e.g. ``12345-1`` → ``12345``)."""
    return re.sub(r"-\d{1,2}$", "", str(pn))


def group_by_base_posting_number(
    focused_rows: list[dict],
) -> dict[str, dict]:
    """Aggregate focused rows by base posting number.

    Returns:
        ``{base_pn: {"收入": float, "收单": float, …, "dates": set[str],
        "skus": set[str], "pns": set[str]}}``
    """
    order_groups: dict[str, dict] = {}
    for fr in focused_rows:
        pn = str(fr["posting_number"])
        base = _base_posting_number(pn)
        if base not in order_groups:
            order_groups[base] = {
                "收入": 0.0, "收单": 0.0, "销售佣金": 0.0,
                "配送佣金": 0.0, "国际配送": 0.0,
                "dates": set(), "skus": set(), "pns": set(),
            }
        g = order_groups[base]
        for field in FIELDS:
            g[field] += fr.get(field, 0.0) or 0.0
        if fr.get("date"):
            g["dates"].add(str(fr["date"]))
        sku = str(fr.get("sku", "-"))
        if sku and sku != "-":
            g["skus"].add(sku)
        g["pns"].add(pn)
    return order_groups
