"""
India daily power data from CEA's National Power Portal (NPP, Central Electricity Authority),
https://npp.gov.in/publishedReports. Found via discovery_archive/subcontinent/SEA_DISCOVERY_INDIA.py and
SEA_DISCOVERY_SOUTH_ASIA2/3.py (Grid-India's own site refuses GitHub runners; NPP answers).

Date-addressed xls reports (no key), each published the day after:
  dgr2   https://npp.gov.in/public-reports/cea/daily/dgr/DD-MM-YYYY/dgr2-YYYY-MM-DD.xls - generation by region,
         state, sector, TYPE, station and unit; the TYPE subtotal rows (THERMAL = coal and lignite steam plants,
         THER (GT) = gas, THER (DG) = diesel, NUCLEAR, HYDRO) give all-India generation by type, MU ('today's
         actual'). From 2019 at least. Conventional plants only: CEA's daily report excludes wind, solar and other
         renewables (CEA's separate daily RE report stopped in Nov 2025), so India's full fuel mix stays on Ember.
  dgr6   .../dgr6-YYYY-MM-DD.xls - daily report of the major hydro reservoirs: energy content at full reservoir
         level and at the present level (MU), per reservoir.
  dailyCoal1  https://npp.gov.in/public-reports/cea/daily/fuel/DD-MM-YYYY/dailyCoal1-YYYY-MM-DD.xls - coal stock at
         thermal plants: normative and actual stock ('000 tonnes), receipt and consumption of the day.

Writes three workbooks in output/Data and Chart Outputs/:
  india_npp_generation_daily.xlsx  'Daily': Coal / Gas / Oil / Nuclear / Hydro MWh per day (conventional only)
  india_hydro_reservoirs.xlsx      'Daily': total energy content of the reported reservoirs (MU) and % of their
                                   energy content at full reservoir level; 'By reservoir': energy content per
                                   reservoir, last 400 days. Water-year chart (Oct-Sep).
  india_coal_stocks.xlsx           'Daily': all-India coal stock at power plants (actual vs normative, days of
                                   stock), receipts and consumption, '000 tonnes

Incremental: each workbook's Daily sheet is the history store; only dates not saved yet (plus REVISION_DAYS) are
fetched, several at a time. Runs on the 1st and 15th.

    python3 asia/INDIA_NPP.py
"""
import argparse
import io
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

NPP = "https://npp.gov.in/public-reports/cea/daily"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
GEN_START, RES_START, COAL_START = date(2021, 1, 1), date(2020, 10, 1), date(2022, 1, 1)
REVISION_DAYS = 3
WORKERS = 6
TYPES = {"THERMAL": "Coal", "THER (GT)": "Gas", "THER (DG)": "Oil", "NUCLEAR": "Nuclear", "HYDRO": "Hydro"}
# plausible all-India conventional generation, MWh/day (catches a wrong column: April-to-date totals are ~100x)
DAY_RANGE = (1.5e6, 8e6)


def out(*a):
    print(*a, flush=True)


def fetch(url):
    for i in range(3):
        try:
            r = requests.get(url, headers=H, timeout=T)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.content
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    return None


def frame(content):
    return pd.read_excel(io.BytesIO(content), header=None)


def text(v):
    return "" if pd.isna(v) else str(v).strip()


def gen_day(d):
    c = fetch(f"{NPP}/dgr/{d:%d-%m-%Y}/dgr2-{d:%Y-%m-%d}.xls")
    if c is None:
        return None
    df = frame(c)
    # 'TODAY'S ACTUAL' column under 'GENERATION (MU)'; its position differs between report versions
    actual = next(j for i in range(min(10, len(df))) for j in range(df.shape[1])
                  if re.search(r"today.?s\s*actual", text(df.iat[i, j]), re.I))
    tot = {}
    for i in range(len(df)):
        if text(df.iat[i, 0]).upper() != "TYPE:":
            continue
        label = next((text(v) for v in df.iloc[i, 1:8] if text(v)), "")
        key = TYPES.get(re.sub(r"\s+", " ", label.upper()))
        v = pd.to_numeric(df.iat[i, actual], errors="coerce")
        if key and pd.notna(v):
            tot[key] = tot.get(key, 0.0) + float(v) * 1000   # MU -> MWh
    if not tot or not DAY_RANGE[0] <= sum(tot.values()) <= DAY_RANGE[1]:
        return None
    return tot


def res_day(d):
    c = fetch(f"{NPP}/dgr/{d:%d-%m-%Y}/dgr6-{d:%Y-%m-%d}.xls")
    if c is None:
        return None
    df = frame(c)
    hdr = next(i for i in range(min(15, len(df))) if any("energy content" in text(v).lower() for v in df.iloc[i]))
    cols = {j: text(v).lower() for j, v in enumerate(df.iloc[hdr])}
    frl = next(j for j, t in cols.items() if "energy content at frl" in t or "energy content at f.r.l" in t)
    now = next(j for j, t in cols.items() if "energy content at present" in t)
    per = {}
    for i in range(hdr + 1, len(df)):
        name = text(df.iat[i, 0])
        if not name or name.isdigit() or re.match(r"^\d+(\.0)?$", name):
            continue
        a, b = pd.to_numeric(df.iat[i, frl], errors="coerce"), pd.to_numeric(df.iat[i, now], errors="coerce")
        if name and not name.lower().startswith(("total", "remark", "note")) and pd.notna(a) and pd.notna(b):
            per[name.title()] = (float(a), float(b))
    if not per:
        return None
    return {"Energy_content_MU": sum(b for _, b in per.values()), "Energy_at_FRL_MU": sum(a for a, _ in per.values()),
            "Reservoirs": len(per), "_per": {k: b for k, (_, b) in per.items()}}


def coal_day(d):
    c = fetch(f"{NPP}/fuel/{d:%d-%m-%Y}/dailyCoal1-{d:%Y-%m-%d}.xls")
    if c is None:
        return None
    df = frame(c)
    hdr = next(i for i in range(15) if any("actual stock" in text(v).lower() for v in df.iloc[i]))
    cols = {j: text(v).lower() for j, v in enumerate(df.iloc[hdr])}
    sub = {j: text(v).lower() for j, v in enumerate(df.iloc[hdr + 1])}
    find = lambda *k: next((j for j, t in cols.items() if all(x in t for x in k)), None)  # noqa: E731
    c_cap, c_norm, c_rec, c_con = find("capacity"), find("normative stock", "tonnes"), find("receipt"), find("consumption")
    c_tot = next((j for j, t in sub.items() if t == "total"), find("actual stock"))
    rows = df.iloc[hdr + 2:]
    lab = rows.apply(lambda r: " ".join(text(v) for v in r.iloc[:8]).upper(), axis=1)
    grand = rows[lab.str.contains("GRAND TOTAL|ALL INDIA")]
    if len(grand):
        pick, how = grand.iloc[[-1]], "grand total row"
    else:   # sum the plant rows: those with a capacity and a plant name, not subtotal rows
        plant = ~lab.str.contains("TOTAL") & pd.to_numeric(rows[c_cap], errors="coerce").notna()
        pick, how = rows[plant], "sum of plant rows"
    num = lambda j: float(pd.to_numeric(pick[j], errors="coerce").sum()) if j is not None else None  # noqa: E731
    actual = num(c_tot)
    if how == "grand total row" and c_norm is not None and c_rec is not None and not actual:
        # older layout: the 'Total' sub-header sits on a merged cell one column off its numbers. The actual-stock
        # block (indigenous, imported, total, then % of normative and days) lies between the normative-stock and
        # receipt columns; its largest number is the total.
        block = pd.to_numeric(pick.iloc[0, c_norm + 1:c_rec], errors="coerce").dropna()
        actual = float(block.max()) if len(block) else None
    return {"Actual_stock_kt": actual, "Normative_stock_kt": num(c_norm), "Receipt_kt": num(c_rec),
            "Consumption_kt": num(c_con), "Capacity_MW": num(c_cap), "_how": how}


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def todo(old, start):
    last = date.today() - timedelta(days=1)
    have = set(old.index.date) if not old.empty else set()
    revise = {last - timedelta(days=k) for k in range(REVISION_DAYS)}
    return [d for d in (start + timedelta(days=k) for k in range((last - start).days + 1)) if d not in have or d in revise]


def safe(fn):
    """A day whose file cannot be read is skipped (and refetched next run), not the whole step."""
    def f(d):
        try:
            return fn(d)
        except Exception as e:  # noqa: BLE001
            print(f"    {fn.__name__} {d}: {type(e).__name__}: {e}", flush=True)
            return None
    return f


def run(fn, days, label):
    fn = safe(fn)
    res = {}
    with ThreadPoolExecutor(WORKERS) as ex:
        for n, (d, r) in enumerate(zip(days, ex.map(fn, days))):
            if r:
                res[pd.Timestamp(d)] = r
            if n % 200 == 0:
                out(f"  {label} {d}: {len(res)} days so far")
    return res


def merge(old, new):
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def generation(path):
    old = read_sheet(path, "Daily")
    if not old.empty:   # drop implausible saved days (an earlier column mix-up) so they are fetched again
        old = old[old["Total_MWh"].between(*DAY_RANGE)]
    days = todo(old, GEN_START)
    out(f"Generation (dgr2): {len(old)} days saved, fetching {len(days)}")
    new = pd.DataFrame.from_dict(run(gen_day, days, "dgr2"), orient="index")
    if not new.empty:
        new = new.reindex(columns=["Hydro", "Gas", "Coal", "Nuclear", "Oil"]).fillna(0).add_suffix("_MWh")
        new["Total_MWh"] = new.sum(axis=1)
        new = new.round(0)
    d = merge(old, new)
    d.index.name = "date"
    notes = ["UNITS",
             "MWh per day (CEA publishes MU = GWh; x1,000), all-India, by plant type from the TYPE subtotals of CEA's "
             "daily generation report (sub-report 2): Coal = 'THERMAL' (coal and lignite steam plants), Gas = 'THER "
             "(GT)', Oil = 'THER (DG)' (diesel), Nuclear, Hydro (large hydro; small hydro is counted as renewable).",
             "", "COVERAGE",
             f"Daily from {d.index.min():%Y-%m-%d} to {d.index.max():%Y-%m-%d}. CONVENTIONAL generation only - wind, solar, "
             "biomass and small hydro are not in this report, and imports from Bhutan are excluded. India's full "
             "generation mix on the dashboard is Ember's (compiled from Grid-India) until a raw renewables feed is added.",
             "", "SOURCE",
             "CEA (Central Electricity Authority) via the National Power Portal, daily generation report sub-report 2: "
             "https://npp.gov.in/public-reports/cea/daily/dgr/DD-MM-YYYY/dgr2-YYYY-MM-DD.xls"]
    xlsx_notes.write_workbook(path, {"Daily": d}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(d)} days" + (f"\n{d.tail(3).to_string()}" if len(d) else ""))


def reservoirs(path):
    old, old_r = read_sheet(path, "Daily"), read_sheet(path, "By reservoir")
    days = todo(old, RES_START)
    out(f"Reservoirs (dgr6): {len(old)} days saved, fetching {len(days)}")
    res = run(res_day, days, "dgr6")
    per = pd.DataFrame({k: v.pop("_per") for k, v in res.items()}).T if res else pd.DataFrame()
    new = pd.DataFrame.from_dict(res, orient="index")
    if not new.empty:
        new["Pct_of_FRL"] = (100 * new["Energy_content_MU"] / new["Energy_at_FRL_MU"]).round(2)
        new = new.round(2)
    d = merge(old, new)
    per = merge(old_r, per)
    per = per[per.index >= per.index.max() - pd.Timedelta(days=400)] if not per.empty else per
    d.index.name = per.index.name = "date"
    notes = ["UNITS",
             "Energy content of the water held in CEA's monitored hydro reservoirs, MU (= GWh): Energy_content_MU = at "
             "the present level, Energy_at_FRL_MU = at full reservoir level (FRL), Pct_of_FRL = the ratio, %. Reservoirs "
             "= number reported that day. By reservoir: energy content per reservoir (MU), last 400 days.",
             "", "COVERAGE",
             f"Daily from {d.index.min():%Y-%m-%d} to {d.index.max():%Y-%m-%d}. The major hydropower reservoirs in CEA's "
             "daily reservoir report (Bhakra, Pong, Tehri, Sardar Sarovar, Indira Sagar, Srisailam, Nagarjunasagar, "
             "Koyna, Idukki, ...).",
             "", "SOURCE",
             "CEA via the National Power Portal, daily report of hydro reservoirs (sub-report 6): "
             "https://npp.gov.in/public-reports/cea/daily/dgr/DD-MM-YYYY/dgr6-YYYY-MM-DD.xls"]
    xlsx_notes.write_workbook(path, {"Daily": d, "By reservoir": per}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(d)} days" + (f"\n{d.tail(3).to_string()}" if len(d) else ""))


def coal(path):
    old = read_sheet(path, "Daily")
    if not old.empty:   # days saved with actual stock 0 (an earlier column mix-up) are fetched again
        old = old[old["Actual_stock_kt"] > 0]
    days = todo(old, COAL_START)
    out(f"Coal stocks (dailyCoal1): {len(old)} days saved, fetching {len(days)}")
    res = run(coal_day, days, "coal")
    how = sorted({v.pop("_how") for v in res.values()})
    out(f"  totals taken from: {how}")
    new = pd.DataFrame.from_dict(res, orient="index")
    if not new.empty:
        new["Days_of_stock"] = (new["Actual_stock_kt"] / new["Consumption_kt"].rolling(7, min_periods=1).mean()).round(1)
        new = new.round(1)
    d = merge(old, new)
    d.index.name = "date"
    notes = ["UNITS",
             "Thousand tonnes ('000 t), all-India thermal power plants in CEA's daily coal stock report: Actual_stock_kt "
             "(indigenous + imported coal at the plants), Normative_stock_kt (stock the plants should hold), Receipt_kt "
             "and Consumption_kt (of the day), Capacity_MW (plants covered). Days_of_stock = actual stock / 7-day average "
             "consumption.",
             "", "COVERAGE",
             f"Daily from {d.index.min():%Y-%m-%d} to {d.index.max():%Y-%m-%d} (the NPP archive of this report starts "
             "in 2022). Totals from: " + ", ".join(how or ["previous runs"]) + ".",
             "", "SOURCE",
             "CEA Fuel Management Division, daily coal stock report, via the National Power Portal: "
             "https://npp.gov.in/public-reports/cea/daily/fuel/DD-MM-YYYY/dailyCoal1-YYYY-MM-DD.xls"]
    xlsx_notes.write_workbook(path, {"Daily": d}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(d)} days" + (f"\n{d.tail(3).to_string()}" if len(d) else ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", choices=["generation", "reservoirs", "coal"])
    args = ap.parse_args()
    steps = {"generation": (generation, "india_npp_generation_daily.xlsx"),
             "reservoirs": (reservoirs, "india_hydro_reservoirs.xlsx"),
             "coal": (coal, "india_coal_stocks.xlsx")}
    for name, (fn, fname) in steps.items():
        if args.only and name != args.only:
            continue
        try:
            fn(os.path.join(OUT_DIR, fname))
        except Exception as e:  # noqa: BLE001
            out(f"{name} failed: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
