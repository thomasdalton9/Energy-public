"""
Norwegian gas exports by destination, daily, from Gassco (operator of the Norwegian gas transport system), one workbook:

  output/Data and Chart Outputs/norway_gassco_gas_flows_daily.xlsx
    sheet "Daily": date (gas day), GWh per day (mcm/d x 11.2 GWh per mcm):
        NO_to_GB (Langeled to Easington + Vesterled/FLAGS to St Fergus), NO_to_DE, NO_to_FR, NO_to_BE, NO_other, NO_total
    sheet "Monthly (mcm)": Gassco's own monthly delivery totals by destination (million Sm3 per month) - used as a check
    sheet "Units": source, definitions, last pull

Source: the data behind the "gas deliveries" graph on gassco.eu, public JSON, no key:
  https://gassco.eu/wp-json/gassco/v1/realTimeGraphData?year=Y&month=M&day=D   (values in million Sm3 per day)
  https://gassco.eu/wp-json/gassco/v1/deliveryNumbers?year=Y                   (monthly totals, million Sm3)
Without parameters realTimeGraphData is the live value (it has no history of its own, hence the accumulating design of the older
snapshot idea); WITH year/month/day it returns the value stored for that day, back to Oct 2020. A day is only taken once the stored
value exists (a day still "live" - today, sometimes yesterday - is skipped and picked up on a later run).
ENTSOG has no Norway flow data (Gassco's points return 404 in operationalData), so this is the only daily Norway -> GB series.

Incremental: the committed workbook is the history store; each run fetches only days not saved yet (plus the last 7 days again).
Days for which Gassco has no value are kept as empty rows so they are not requested again. Days that failed (HTTP error) are not
saved and are retried next run.

Usage: python3 NORWAY_GASSCO_DAILY.py [--out-dir DIR] [--start 2020-10-01] [--workers 3]
"""
import argparse
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "norway_gassco_gas_flows_daily.xlsx"
BASE = "https://gassco.eu/wp-json/gassco/v1/"
GWH_PER_MCM = 11.2
REVISION_DAYS = 7
SAVE_EVERY = 120
COLUMNS = ["NO_to_GB", "NO_to_DE", "NO_to_FR", "NO_to_BE", "NO_other", "NO_total"]
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
           "Accept": "application/json"}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def get_json(path, params, tries=5):
    """('ok', json) / ('empty', None) for an empty body / ('err', None) after retries."""
    for k in range(tries):
        try:
            r = SESSION.get(BASE + path, params=params, timeout=45)
            if r.status_code == 200:
                if not r.text.strip():
                    return "empty", None
                return "ok", r.json()
            code = r.status_code
        except Exception as e:  # noqa: BLE001
            code = type(e).__name__
        time.sleep(2 + 4 * k)
    print(f"  {path} {params}: failed ({code})", flush=True)
    return "err", None


def parse_day(j):
    """Graph JSON -> dict of mcm/d. The old format (to about mid 2022) has no Total: its 'Other' holds the total."""
    v = {n["name"]: n["value"] for n in j["nodes"]}
    de, fr, be, gb = (v.get(k) for k in ("Germany", "France", "Belgium", "Great Britain"))
    if None in (de, fr, be, gb):
        return None
    tot = v.get("Total")
    oth = v.get("Other")
    if tot is None:
        tot = oth
        oth = None if tot is None else max(tot - (de + fr + be + gb), 0.0)
    if tot is None or oth is None:
        return None
    return {"NO_to_GB": gb, "NO_to_DE": de, "NO_to_FR": fr, "NO_to_BE": be, "NO_other": oth, "NO_total": tot}


def fetch_day(d):
    status, j = get_json("realTimeGraphData", {"year": d.year, "month": d.month, "day": d.day})
    if status != "ok":
        return d, status, None
    ts = datetime.fromtimestamp(j["timestamp"] / 1000, tz=timezone.utc)
    if ts > datetime.now(timezone.utc) - timedelta(hours=3):
        return d, "live", None          # no stored value for this day yet: the API answers with the live reading
    row = parse_day(j)
    return d, ("ok" if row else "empty"), row


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=COLUMNS)
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=COLUMNS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in COLUMNS:
        if c not in d:
            d[c] = np.nan
    return d[COLUMNS]


def monthly_deliveries(first_year):
    rows = []
    months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
    for y in range(first_year, datetime.now(timezone.utc).year + 1):
        status, j = get_json("deliveryNumbers", {"year": y})
        if status != "ok":
            continue
        for i, m in enumerate(months, 1):
            x = j.get(m) or {}
            if not x.get("total"):
                continue
            rows.append({"month": pd.Timestamp(y, i, 1), "NO_to_GB": x.get("gb"), "NO_to_DE": x.get("germany"), "NO_to_FR": x.get("france"),
                         "NO_to_BE": x.get("belgium"), "NO_other": x.get("other"), "NO_total": x.get("total")})
    return pd.DataFrame(rows, columns=["month"] + COLUMNS).set_index("month") if rows else pd.DataFrame(columns=COLUMNS)


def save(path, daily, monthly, start, note=""):
    daily = daily.sort_index()
    daily.index.name = "date"
    have = daily.dropna(how="all")
    check = ""
    if len(monthly) and len(have):
        # daily mcm/d x days in the month vs Gassco's monthly total, for months with every day present
        m = (have / GWH_PER_MCM).resample("MS").agg(["mean", "count"])
        comp = []
        for mo, r in monthly.iterrows():
            if mo in m.index and m.loc[mo, ("NO_total", "count")] >= mo.days_in_month:
                for c in ("NO_total", "NO_to_GB"):
                    comp.append((c, m.loc[mo, (c, "mean")] * mo.days_in_month / r[c] - 1))
        if comp:
            s = pd.DataFrame(comp, columns=["c", "x"]).groupby("c")["x"].agg(["mean", "count"])
            check = "; ".join(f"{c}: daily x days vs monthly total {r['mean']:+.1%} on average over {int(r['count'])} complete months" for c, r in s.iterrows())
            print("monthly check:", check, flush=True)
    lines = ["Norway - gas exports by destination (Gassco)", "",
             "Source", "Gassco AS, operator of the integrated Norwegian gas transport system: the data behind the gas-deliveries graph on "
             "https://gassco.eu (public JSON, no key): /wp-json/gassco/v1/realTimeGraphData?year=&month=&day= (daily) and /deliveryNumbers?year= "
             "(monthly). Operator's own data. ENTSOG carries no Norway flows (Gassco points have no operational data there), so this is the "
             "source for Norway -> Great Britain.",
             "", "Units and definitions",
             f"Daily: GWh per gas day = Gassco's million Sm3 per day x {GWH_PER_MCM} GWh per million Sm3 (an assumed typical gross calorific value of "
             "Norwegian export gas; the true value varies by a few percent by blend). NO_to_GB = Langeled (Easington) + Vesterled/FLAGS "
             "(St Fergus); NO_to_DE = Dornum, Emden (Europipe, Norpipe, Norsea); NO_to_FR = Dunkerque; NO_to_BE = Zeebrugge; NO_other = "
             "Gassco 'Other' (Netherlands, Denmark and others); NO_total = Gassco's total (sum of destinations in the older format "
             "that has no total). The older format (Oct 2020 to about mid 2022) gives the total in the 'Other' field; NO_other is then "
             "total minus the four named destinations. "
             "The day value is the graph reading Gassco stores for that day; a day is saved only once it is stored (the most recent "
             "one or two days are still 'live' and are added on a later run). Compare with the GB NTS St Fergus + Easington nominations "
             "(gb_gas_nts_daily.xlsx): those include UKCS gas, so Norway -> GB is the smaller part and moves with them.",
             "Monthly (mcm): Gassco's monthly delivery totals by destination, million Sm3 per month (not converted), a cross-check of the daily series.",
             f"History from {start}; re-fetches the last {REVISION_DAYS} days each run plus days not yet saved. Days Gassco has no value for are kept as empty rows.",
             "Monthly check: " + (check or "n/a"), "",
             "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(have)} days with data, "
             f"{have.index.min():%Y-%m-%d} to {have.index.max():%Y-%m-%d}" + (f"; {note}" if note else "")]
    sheets = {"Daily": daily.round(1)}
    if len(monthly):
        mm = monthly.copy()
        mm.index.name = "month"
        sheets["Monthly (mcm)"] = mm.round(1)
    xlsx_notes.write_workbook(path, sheets, lines, {"Source", "Units and definitions", "Last pull"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2020-10-01")
    ap.add_argument("--workers", type=int, default=3)
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = read_existing(path)
    have = set(old.index.date) if len(old) else set()
    todo = [start + timedelta(days=i) for i in range((today - start).days)]
    todo = [d for d in todo if d not in have or d >= today - timedelta(days=REVISION_DAYS)]
    print(f"{len(todo)} days to fetch ({len(have)} already saved)", flush=True)
    monthly = monthly_deliveries(start.year)
    daily = old.copy()
    n_ok = n_empty = n_err = n_live = 0
    since_save = 0
    for i in range(0, len(todo), SAVE_EVERY):
        chunk = todo[i:i + SAVE_EVERY]
        with ThreadPoolExecutor(args.workers) as ex:
            res = list(ex.map(fetch_day, chunk))
        new = {}
        for d, status, row in res:
            ts = pd.Timestamp(d)
            if status == "ok":
                new[ts] = row
                n_ok += 1
            elif status == "empty":
                new[ts] = {c: np.nan for c in COLUMNS}
                n_empty += 1
            elif status == "live":
                n_live += 1
            else:
                n_err += 1
        if new:
            nd = pd.DataFrame.from_dict(new, orient="index")[COLUMNS]
            nd[COLUMNS[:-1] + ["NO_total"]] = nd[COLUMNS] * GWH_PER_MCM
            daily = pd.concat([daily[~daily.index.isin(nd.index)], nd]).sort_index()
        print(f"  {chunk[0]} .. {chunk[-1]}: ok {n_ok}, empty {n_empty}, live {n_live}, failed {n_err} so far", flush=True)
        if new and not daily.dropna(how="all").empty:
            save(path, daily, monthly, args.start)
    if daily.dropna(how="all").empty:
        print("no Norway gas data")
        return
    save(path, daily, monthly, args.start, note=f"this run: {n_ok} days saved, {n_empty} empty, {n_err} failed (retried next run)")
    print(f"saved {FILE}: {len(daily.dropna(how='all'))} days; failed this run: {n_err}")


if __name__ == "__main__":
    main()
