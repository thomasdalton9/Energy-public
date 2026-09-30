"""
National hydro reservoir storage for Colombia, from XM (the Colombian
wholesale power market operator)'s free, no-auth public API
(servapibi.xm.com.co).

Pulls two daily national ("Sistema" entity) metrics directly - no
per-region weighting needed here, unlike Brazil's equivalent pull,
because XM already publishes the national aggregate:
  - VoluUtilDiarEner: national reservoir stored energy, kWh ("Volumen
    Util diario Energia por Sistema").
  - CapaUtilDiarEner: national reservoir capacity, kWh ("Capacidad Util
    Energia por Sistema").

StoragePct = stored / capacity * 100. Both MetricIds were confirmed
against live XM responses (Sep-2026); the ratio matches XM's own
PorcVoluUtilDiar. Daily totals are cached in co_hydro_daily_cache.csv,
so later runs only fetch the last two weeks.

This is the "high confidence" half of the Colombia build - see
COLOMBIA_XM_GENERATION_DISCOVERY.py's docstring for why generation-by-
fuel-type is a separate, not-yet-built, follow-up.
"""

print("STARTING", flush=True)

import datetime as dt
import time

import pandas as pd
import requests

BASE_DAILY_URL = "https://servapibi.xm.com.co/daily"
ENTITY = "Sistema"
START_DATE = dt.date(2010, 1, 1)
END_DATE = dt.date.today()

# XM's own client library caps daily-period requests at a 30-day
# window; chunk one day short of that to stay safely inside it.
CHUNK_DAYS = 29

# Confirmed against live responses (Sep-2026, COLOMBIA_XM_GENERATION_DISCOVERY.py);
# the earlier VolUti/CapUti guesses get 400 from XM. Both are kWh, and
# their ratio matches XM's own PorcVoluUtilDiar.
METRICS = {
    "VoluUtilDiarEner": "StoredEnergy",
    "CapaUtilDiarEner": "Capacity",
}


def date_chunks(start, end, size_days):
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=size_days), end)
        yield cur, chunk_end
        cur = chunk_end + dt.timedelta(days=1)


def extract_records(payload):
    """
    XM wraps results as {"Items": [...]}. The exact inner shape isn't
    documented publicly, but every source seen describes each item as
    a dict holding one list-valued key (e.g. "DailyEntities") with the
    actual rows - detect that structurally instead of hardcoding the
    wrapper key name, so a minor naming difference doesn't silently
    drop data.
    """
    items = payload.get("Items")
    if items is None:
        raise KeyError(f"No 'Items' key in response - keys were: {list(payload.keys())}")

    if isinstance(items, dict):
        items = [items]

    records = []
    for item in items:
        if isinstance(item, dict):
            list_values = [v for v in item.values() if isinstance(v, list)]
            if list_values:
                # the day sits on the item, the value on its entities:
                # {"Date": ..., "DailyEntities": [{"Id": "Sistema", "Value": ...}]}
                parent = {k: v for k, v in item.items() if not isinstance(v, list)}
                for v in list_values:
                    records.extend({**parent, **r} if isinstance(r, dict) else r for r in v)
            else:
                records.append(item)
        else:
            records.append(item)

    return records


def fetch_metric(metric_id, start, end):
    all_records = []
    failed_chunks = 0

    for chunk_start, chunk_end in date_chunks(start, end, CHUNK_DAYS):
        body = {
            "MetricId": metric_id,
            "Entity": ENTITY,
            "StartDate": chunk_start.isoformat(),
            "EndDate": chunk_end.isoformat(),
        }

        try:
            r = requests.post(
                BASE_DAILY_URL,
                json=body,
                headers={"Connection": "close"},
                timeout=60,
            )
            r.raise_for_status()
            payload = r.json()
            records = extract_records(payload)
        except Exception as e:
            failed_chunks += 1
            print(
                f"  {metric_id} {chunk_start}..{chunk_end}: FAILED ({type(e).__name__}: {e})",
                flush=True,
            )
            continue

        all_records.extend(records)
        time.sleep(0.2)

    if failed_chunks:
        print(f"  {metric_id}: {failed_chunks} chunk(s) failed", flush=True)

    return all_records


def pick(columns, *keywords, exclude=()):
    for c in columns:
        lc = c.lower()
        if all(k in lc for k in keywords) and not any(x in lc for x in exclude):
            return c
    raise KeyError(f"Could not find a column matching {keywords} in {list(columns)}")


def records_to_series(records, value_name):
    if not records:
        raise RuntimeError(
            f"No records returned for {value_name} - MetricId/Entity may no "
            "longer be valid, or the API's response shape changed. Run "
            "COLOMBIA_XM_GENERATION_DISCOVERY.py to inspect a raw response."
        )

    df = pd.json_normalize(records)

    date_col = pick(df.columns, "date")
    value_col = pick(df.columns, "value")

    df = df[[date_col, value_col]].copy()
    df.columns = ["Date", value_name]

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df[value_name] = pd.to_numeric(df[value_name], errors="coerce")
    df = df.dropna(subset=["Date"]).drop_duplicates(subset="Date", keep="last")

    return df.set_index("Date")[value_name].sort_index()


# ----------------------------------------------------------
# PULL
# ----------------------------------------------------------

# Cache: the full series is kept in CACHE_CSV, so a later run only asks
# XM for the last REFRESH_DAYS (XM can revise recent days) instead of
# re-downloading everything since 2010.
CACHE_CSV = "co_hydro_daily_cache.csv"
REFRESH_DAYS = 14
try:
    cached = pd.read_csv(CACHE_CSV, parse_dates=["Date"]).set_index("Date").sort_index()
    fetch_from = max(START_DATE, (cached.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date())
    print(f"{len(cached):,} days cached to {cached.index.max():%d-%b-%Y}; fetching from {fetch_from}", flush=True)
except (FileNotFoundError, ValueError, KeyError):
    cached = None
    fetch_from = START_DATE

series = {}

for metric_id, value_name in METRICS.items():
    print(f"Downloading {metric_id} ({ENTITY}, {fetch_from}..{END_DATE})...", flush=True)
    records = fetch_metric(metric_id, fetch_from, END_DATE)
    series[value_name] = records_to_series(records, value_name)
    print(f"  {metric_id}: {len(series[value_name]):,} daily values", flush=True)

fresh = pd.concat([series["StoredEnergy"], series["Capacity"]], axis=1).dropna(how="all")
daily = fresh if cached is None else fresh.combine_first(cached[["StoredEnergy", "Capacity"]])
daily = daily.sort_index()
daily.index.name = "Date"
daily.to_csv(CACHE_CSV)
daily["StoragePct"] = (daily["StoredEnergy"] / daily["Capacity"]) * 100
daily = daily.reset_index()

# ----------------------------------------------------------
# WRITE FILES
# ----------------------------------------------------------

daily[["Date", "StoredEnergy"]].to_csv("co_hydro_daily_energy.csv", index=False)
daily[["Date", "StoragePct"]].to_csv("co_hydro_daily_pct.csv", index=False)

print("\nLATEST NATIONAL HYDRO STORAGE")
print(daily.tail(12).to_string(index=False, float_format="%.1f"))

print(
    "\nSaved:\n"
    "  co_hydro_daily_energy.csv (StoredEnergy in kWh)\n"
    "  co_hydro_daily_pct.csv\n",
    flush=True,
)
