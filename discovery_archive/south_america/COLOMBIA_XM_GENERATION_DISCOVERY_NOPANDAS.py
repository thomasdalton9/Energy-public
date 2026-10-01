"""
Same discovery as COLOMBIA_XM_GENERATION_DISCOVERY.py - see that
file's docstring for the full why - but with the pandas dependency
removed, for running on a phone (e.g. Pyto or a-Shell on iPhone) where
pandas' C extensions often aren't available as a pre-built wheel, even
though pure-Python `requests` works fine. Only needs `requests` (and
the standard library) - if your Python app can `pip install requests`
(or already bundles it), this should run as-is.

Does exactly the same three API calls and prints exactly the same
kind of summary (top-level keys, first item, column names, unique
values per column) - just re-implemented with plain dicts/lists
instead of a pandas DataFrame.

Run this, then share the printed output (or the saved
co_discovery_*_sample.json files) back so COLOMBIA_XM_GENERATION.py
can be written against the real response shape.
"""

print("STARTING", flush=True)

import json
import datetime as dt

import requests

HEADERS = {"Connection": "close"}
RECENT_END = dt.date.today()
RECENT_START = RECENT_END - dt.timedelta(days=6)


def try_post(label, url, body):
    print(f"\n{'=' * 70}\n{label}\nPOST {url}\nBody: {body}\n{'=' * 70}", flush=True)
    try:
        r = requests.post(url, json=body, headers=HEADERS, timeout=60)
        print(f"  status: {r.status_code}", flush=True)
        r.raise_for_status()
        payload = r.json()
        return payload
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", flush=True)
        return None


def flatten(d, prefix=""):
    """Minimal stand-in for pandas.json_normalize on a single dict -
    flattens one level of nested dicts as "parent.child" keys. Good
    enough for a discovery script's own inspection; not meant to
    handle arbitrarily deep/irregular nesting."""
    out = {}
    for k, v in d.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            out.update(flatten(v, prefix=f"{key}."))
        else:
            out[key] = v
    return out


def summarise(payload, out_file):
    if payload is None:
        return

    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    print(f"  Saved raw response -> {out_file}", flush=True)

    print(f"  Top-level keys: {list(payload.keys())}", flush=True)

    items = payload.get("Items")
    if items is None:
        print("  No 'Items' key found.", flush=True)
        return

    print(f"  Items type: {type(items).__name__}, len: {len(items) if hasattr(items, '__len__') else 'n/a'}", flush=True)
    if isinstance(items, list) and items:
        print(f"  First item: {items[0]}", flush=True)
        if isinstance(items[0], dict):
            print(f"  First item keys: {list(items[0].keys())}", flush=True)

    if not (isinstance(items, list) and items and isinstance(items[0], dict)):
        return

    rows = [flatten(item) for item in items]
    columns = []
    for row in rows:
        for k in row:
            if k not in columns:
                columns.append(k)
    print(f"  Normalised columns: {columns}", flush=True)
    print(f"  Row count: {len(rows)}", flush=True)

    print("  First 10 rows:", flush=True)
    for row in rows[:10]:
        print(f"    {row}", flush=True)

    for col in columns:
        values = [row.get(col) for row in rows if col in row]
        # Only look at string-ish columns, same as the pandas version's
        # dtype == object check - skip anything clearly numeric.
        if values and any(isinstance(v, str) for v in values):
            unique = sorted({v for v in values if v is not None}, key=lambda x: str(x))
            if len(unique) <= 60:
                print(f"  Unique values in '{col}': {unique}", flush=True)


# ----------------------------------------------------------
# 1. Full metric catalog, filtered to anything generation/resource-related -
#    confirms exact MetricId spelling and valid Entity values.
# ----------------------------------------------------------

catalog = try_post(
    "Metric catalog (Lists endpoint, no MetricId filter - some XM API "
    "variants return the full catalog this way; may 400/404, that's fine, "
    "just informational)",
    "https://servapibi.xm.com.co/lists",
    {},
)
summarise(catalog, "co_discovery_catalog_sample.json")

# ----------------------------------------------------------
# 2. Resource listing - looking for a fuel/technology-type column.
# ----------------------------------------------------------

for entity in ("Sistema", "Recurso"):
    payload = try_post(
        f"ListadoRecursos (Entity={entity})",
        "https://servapibi.xm.com.co/lists",
        {"MetricId": "ListadoRecursos", "Entity": entity},
    )
    summarise(payload, f"co_discovery_listadorecursos_{entity.lower()}_sample.json")

# ----------------------------------------------------------
# 3. Per-plant generation, last 7 days - confirms the "Gene" MetricId,
#    the per-plant resource-code field name, and record shape.
# ----------------------------------------------------------

gene_payload = try_post(
    f"Gene (Entity=Recurso, {RECENT_START}..{RECENT_END})",
    "https://servapibi.xm.com.co/daily",
    {
        "MetricId": "Gene",
        "Entity": "Recurso",
        "StartDate": RECENT_START.isoformat(),
        "EndDate": RECENT_END.isoformat(),
    },
)
summarise(gene_payload, "co_discovery_gene_recurso_sample.json")

print(
    "\nDONE. Please share the console output above (or the "
    "co_discovery_*_sample.json files) so COLOMBIA_XM_GENERATION.py can be "
    "written against the real response shape.",
    flush=True,
)
