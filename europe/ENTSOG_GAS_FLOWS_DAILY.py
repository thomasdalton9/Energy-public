"""
Gas flows by country for Europe from the ENTSOG Transparency Platform (physical flows at every interconnection
point, API operationalData / "Physical Flow", daily), one workbook:

  output/Data and Chart Outputs/europe_gas_flows_daily.xlsx
    sheet "Country balance": date, then per country <CC>_<category>_GWhd with categories
        production        gas entering the country's grid from production points
        imports           pipeline gas entering from another country's grid (any origin, EU or not)
        exports           pipeline gas leaving to another country's grid
        lng               gas entering from LNG terminals (ALSI's terminal send-out is the better LNG series; kept for checks)
        storage_out / storage_in   gas leaving / entering storage (AGSI+ is the better storage series)
        distribution      gas leaving to distribution networks (households, commerce, small industry)
        final_consumers   gas leaving directly to large final consumers (industry, power plants)
    sheet "Imports by origin": date, one column per origin country, GWh per day entering EU27 grids from outside the EU (the UK counts as outside)
    sheet "Exports by destination": date, one column per destination country, GWh per day leaving EU27 grids to outside the EU
    sheet "Units": source and definitions

Each flow row is (point, operator, entry|exit). It is classed by the system on the other side of the point, taken from
ENTSOG's interconnections list: the operator's country is the first two letters of its key, the adjacent system gives
the type (Transmission / Production / LNG Terminals / Storage / Distribution / Final Consumers) and country.
Cross-border flows are counted on the reporting country's own side only (its entry = import, its exit = export), so a
border is not counted twice. ENTSOG reports kWh/d; converted to GWh/d. Empty (unreported) values are left blank.

Caveats: ENTSOG data are operational flows (allocations/nominations), restated for recent days (the last 45 days are
re-fetched each run) and incomplete for some points; the platform keeps about 5 years, so the committed workbook is the
history store. Germany reports consumption as aggregated final-consumer points; Spain has few demand points. A country's
balance (supply - uses) therefore does not always close - the master shows the gap rather than hiding it.

Usage: python3 ENTSOG_GAS_FLOWS_DAILY.py [--out-dir DIR] [--start 2021-10-04]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

API = "https://transparency.entsog.eu/api/v1"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_gas_flows_daily.xlsx"
START_DEFAULT = "2021-10-04"
RELOAD_DAYS = 45
WINDOW_DAYS = 31
EU27 = {"AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU",
        "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"}
CATEGORIES = ["production", "imports", "exports", "lng", "storage_out", "storage_in", "distribution", "final_consumers"]
# ENTSOG country codes shown in the master (UK = Great Britain / Northern Ireland)
NAMES = {"AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "HR": "Croatia", "CZ": "Czechia", "DK": "Denmark",
         "EE": "Estonia", "FI": "Finland", "FR": "France", "DE": "Germany", "GR": "Greece", "HU": "Hungary",
         "IE": "Ireland", "IT": "Italy", "LV": "Latvia", "LT": "Lithuania", "LU": "Luxembourg", "NL": "Netherlands",
         "PL": "Poland", "PT": "Portugal", "RO": "Romania", "SK": "Slovakia", "SI": "Slovenia", "ES": "Spain",
         "SE": "Sweden", "UK": "United Kingdom", "CH": "Switzerland"}


def get_json(path, params, tries=4):
    last = ""
    for attempt in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=HEADERS, timeout=(15, 300))
        except requests.RequestException as e:
            last = f"{type(e).__name__}"
            time.sleep(10 * (attempt + 1))
            continue
        if r.status_code in (429,) or r.status_code >= 500:
            last = f"HTTP {r.status_code}"
            time.sleep(20 * (attempt + 1))
            continue
        if r.status_code == 404:   # ENTSOG answers "no result" with 404
            return {}
        if r.status_code != 200:
            raise RuntimeError(f"{path} {params}: HTTP {r.status_code} {r.text[:150]}")
        return r.json()
    raise RuntimeError(f"{path} {params}: failed after {tries} tries ({last})")


def adjacency():
    """(pointKey, operatorKey, direction) -> (adjacent infrastructure type, adjacent country key)."""
    ics = get_json("interconnections", {"limit": -1}).get("interconnections", [])
    adj = {}
    for i in ics:
        pk = i["pointKey"]
        if i.get("toOperatorKey"):
            adj[(pk, i["toOperatorKey"], "entry")] = (i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"))
        if i.get("fromOperatorKey"):
            adj[(pk, i["fromOperatorKey"], "exit")] = (i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"))
    return adj


def classify(row, adj):
    """-> (country, category, origin) or None. origin is the adjacent country for entries (imports)."""
    op, pk, d = row.get("operatorKey") or "", row["pointKey"], row["directionKey"]
    country = op[:2]
    a = adj.get((pk, op, d))
    if not country or a is None:
        return None
    typ, acountry = a
    if typ == "Transmission":
        if acountry and acountry != country:
            return country, ("imports" if d == "entry" else "exports"), acountry
        return None   # between two systems of the same country
    if typ == "Production" and d == "entry":
        return country, "production", None
    if typ == "LNG Terminals" and d == "entry":
        return country, "lng", None
    if typ == "Storage":
        return country, ("storage_out" if d == "entry" else "storage_in"), None
    if typ == "Distribution" and d == "exit":
        return country, "distribution", None
    if typ == "Final Consumers" and d == "exit":
        return country, "final_consumers", None
    return None


def fetch_window(d0, d1, adj):
    """Daily GWh/d per (country, category) and per origin for [d0, d1]."""
    data = get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0.isoformat(),
                                        "to": d1.isoformat(), "limit": -1}).get("operationalData", [])
    cat, origin, dest, unclassified = {}, {}, {}, 0
    for r in data:
        v = r.get("value")
        if v in (None, ""):
            continue
        try:
            gwh = float(v) / 1e6 if r.get("unit", "kWh/d") == "kWh/d" else float(v) * {"MWh/d": 1e-3, "GWh/d": 1.0}.get(r["unit"], float("nan"))
        except (TypeError, ValueError):
            continue
        c = classify(r, adj)
        if c is None:
            unclassified += 1
            continue
        country, category, orig = c
        day = pd.Timestamp(r["periodFrom"][:10])
        cat[(day, f"{country}_{category}_GWhd")] = cat.get((day, f"{country}_{category}_GWhd"), 0.0) + gwh
        if category == "imports" and orig not in EU27 and country in EU27:
            origin[(day, orig)] = origin.get((day, orig), 0.0) + gwh
        if category == "exports" and orig not in EU27 and country in EU27:
            dest[(day, orig)] = dest.get((day, orig), 0.0) + gwh
    return cat, origin, dest, len(data), unclassified


def to_frame(d, name):
    if not d:
        return pd.DataFrame()
    s = pd.Series(d)
    s.index = pd.MultiIndex.from_tuples(s.index)
    return s.unstack().sort_index()


def read_sheet(path, sheet):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    return d.dropna(subset=["date"]).set_index("date").sort_index()


def merge(old, new):
    if old.empty:
        return new.sort_index()
    out = old.reindex(old.index.union(new.index))
    for c in new.columns:
        if c not in out:
            out[c] = float("nan")
        s = new[c].dropna()
        out.loc[s.index, c] = s
    return out.sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default=START_DEFAULT)
    ap.add_argument("--max-minutes", type=float, default=90.0)
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    deadline = time.time() + args.max_minutes * 60
    start = date.fromisoformat(args.start)
    end = date.today() - timedelta(days=1)
    bal, org, dst = (read_sheet(path, "Country balance"), read_sheet(path, "Imports by origin"),
                     read_sheet(path, "Exports by destination"))
    fs = start
    if len(bal):
        fs = max(start, (bal.index.max() - timedelta(days=RELOAD_DAYS)).date())
    print(f"fetching {fs} -> {end} (have {len(bal)} days)", flush=True)
    adj = adjacency()
    print(f"{len(adj)} (point, operator, direction) classifications", flush=True)
    cur = fs
    while cur <= end:
        if time.time() > deadline:
            print("time budget reached; stopping (windows done so far are saved)")
            break
        nxt = min(cur + timedelta(days=WINDOW_DAYS - 1), end)
        t0 = time.time()
        cat, origin, dest, n, unc = fetch_window(cur, nxt, adj)
        print(f"  {cur} -> {nxt}: {n} rows, {unc} unclassified, {len(cat)} country-day values in {time.time() - t0:.0f}s", flush=True)
        bal = merge(bal, to_frame(cat, "balance"))
        org = merge(org, to_frame(origin, "origin"))
        dst = merge(dst, to_frame(dest, "dest"))
        bal.index.name = org.index.name = dst.index.name = "date"
        cur = nxt + timedelta(days=1)
        time.sleep(1)
    if bal.empty:
        print("no data")
        return
    bal = bal[sorted(bal.columns)].round(3)
    org = org[sorted(org.columns)].round(3) if not org.empty else org
    dst = dst[sorted(dst.columns)].round(3) if not dst.empty else dst
    lines = ["Europe - gas flows by country (ENTSOG Transparency Platform)", "",
             "Source", "ENTSOG Transparency Platform, operational data: Physical Flow, daily, at every interconnection point "
             "(https://transparency.entsog.eu/). The TSOs' own reporting.",
             "", "Units and definitions",
             "GWh per gas day (ENTSOG kWh/d / 1e6). <CC>_<category>_GWhd: production, imports (pipeline gas entering from another country), "
             "exports (leaving to another country), lng (entering from LNG terminals), storage_out / storage_in, distribution "
             "(exits to distribution networks) and final_consumers (exits to large consumers: industry, power plants). UK = Great Britain "
             "and Northern Ireland.",
             "Each flow row is classed by the system on the other side of the point (ENTSOG interconnections list); cross-border flows "
             "are counted on the reporting country's own side only. Imports by origin / Exports by destination: pipeline gas entering EU27 grids from, or leaving "
             "to, countries outside the EU27 (the UK counts as outside), by country.",
             "Operational data: restated for recent days (last " + str(RELOAD_DAYS) + " days re-fetched each run) and unreported for "
             "some points; Germany reports aggregated final consumers, Spain has few demand points, so a country's supply and uses do "
             "not always balance. ENTSOG keeps about 5 years of history; this workbook is the history store.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(bal)} days, {bal.index.min():%Y-%m-%d} to {bal.index.max():%Y-%m-%d}"]
    sheets = {"Country balance": bal}
    if not org.empty:
        sheets["Imports by origin"] = org
    if not dst.empty:
        sheets["Exports by destination"] = dst
    xlsx_notes.write_workbook(path, sheets, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(bal)} days x {bal.shape[1]} columns; origins: {list(org.columns) if not org.empty else 'none'}")


if __name__ == "__main__":
    main()
