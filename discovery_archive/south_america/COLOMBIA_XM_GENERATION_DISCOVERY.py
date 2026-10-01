"""
DONE - COLOMBIA_XM_GENERATION.py is now built from what this found
(Sep-2026): the plant list with fuels, and Gene per plant on /hourly.
Kept for re-checking the API if it changes.

NOT part of the automated pipeline (not in MASTER_SOUTH_AMERICA.py's
SCRIPTS list) - a one-off, run-it-yourself script to capture real XM
API responses for Colombia generation-by-fuel-type, so the actual pull
script can be written against confirmed data instead of guessed field
names.

Why this is needed: unlike Brazil (one ready-made ONS parquet file
with pre-split hydro/thermal/wind/solar columns) and unlike Colombia's
own hydro storage (COLOMBIA_XM_HYDRO.py - a confirmed national-level
metric, no classification needed), Colombia's generation data from XM
is per-plant ("Gene" metric, Entity="Recurso") with no confirmed
public documentation of exactly how each plant's fuel/technology type
is labelled in the resource-listing metric ("ListadoRecursos") -
different sources hint at a "Tipo" field but the exact column name and
its values (e.g. is hydro "HIDRAULICA", "HIDRO", something else?)
could not be verified from here (this environment's network can't
reach XM's API at all to test live).

Run this on a machine with real network access to
servapibi.xm.com.co, then share the printed output (or the saved
*_sample.json files) back so COLOMBIA_XM_GENERATION.py can be written
against the real shape rather than a guess.
"""

print("STARTING", flush=True)

import json

import pandas as pd
import requests
import datetime as dt

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

    try:
        df = pd.json_normalize(items)
        print(f"  Normalised columns: {list(df.columns)}", flush=True)
        print(f"  Row count: {len(df)}", flush=True)
        print(df.head(10).to_string(), flush=True)

        for col in df.columns:
            if df[col].dtype == object and df[col].nunique() <= 60:
                print(f"  Unique values in '{col}': {sorted(df[col].dropna().unique().tolist())}", flush=True)
    except Exception as e:
        print(f"  Could not normalise Items into a DataFrame: {e}", flush=True)


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
# 1b. XM's metric catalogue (MetricId=ListadoMetricas) - the exact
#     MetricIds for reservoir volume/capacity, which COLOMBIA_XM_HYDRO.py
#     needs (its VolUti/CapUti guesses get 400 on /daily).
# ----------------------------------------------------------
metrics = try_post("ListadoMetricas (metric catalogue)", "https://servapibi.xm.com.co/lists",
                   {"MetricId": "ListadoMetricas"})
summarise(metrics, "co_discovery_metric_catalogue.json")
if metrics:
    for item in metrics.get("Items", []):
        for entity in item.get("ListEntities", []):
            v = entity.get("Values", {})
            text = " ".join(str(x) for x in v.values())
            if any(k in text.lower() for k in ("volu", "capa", "embal", "reserv", "porc")):
                print("  METRIC", v, flush=True)

# candidate reservoir MetricIds (pydataxm's names), 7 days ~3 weeks back
for metric_id in ("VoluUtilDiarEner", "CapaUtilDiarEner", "PorcVoluUtilDiar"):
    payload = try_post(f"{metric_id} (Entity=Sistema)", "https://servapibi.xm.com.co/daily",
                       {"MetricId": metric_id, "Entity": "Sistema",
                        "StartDate": (RECENT_END - dt.timedelta(days=24)).isoformat(),
                        "EndDate": (RECENT_END - dt.timedelta(days=18)).isoformat()})
    summarise(payload, f"co_discovery_{metric_id.lower()}_sample.json")

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

# ----------------------------------------------------------
# 4. The /daily call above returns 400 (confirmed Sep-2026): Gene per
#    plant is an hourly metric, so try XM's /hourly endpoint. The metric
#    exists there, but 2-3 days back came back empty - XM publishes with
#    a lag - so ask for a day ~2 weeks back.
# ----------------------------------------------------------
for entity in ("Recurso", "Sistema"):
    hourly_payload = try_post(
        f"Gene hourly (Entity={entity}, {RECENT_END - dt.timedelta(days=15)}..{RECENT_END - dt.timedelta(days=14)})",
        "https://servapibi.xm.com.co/hourly",
        {
            "MetricId": "Gene",
            "Entity": entity,
            "StartDate": (RECENT_END - dt.timedelta(days=15)).isoformat(),
            "EndDate": (RECENT_END - dt.timedelta(days=14)).isoformat(),
        },
    )
    summarise(hourly_payload, f"co_discovery_gene_hourly_{entity.lower()}_sample.json")

print(
    "\nDONE. Please share the console output above (or the "
    "co_discovery_*_sample.json files) so COLOMBIA_XM_GENERATION.py can be "
    "written against the real response shape.",
    flush=True,
)
