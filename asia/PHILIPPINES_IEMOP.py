"""
Philippines wholesale electricity market (WESM) data from IEMOP, the Independent Electricity Market
Operator of the Philippines, https://www.iemop.ph/market-data/. Found via discovery_archive/asia/SEA_DISCOVERY3.py
and SEA_DISCOVERY4.py.

IEMOP's market-data pages list their files through WordPress admin-ajax (action
display_filtered_market_data_files, one post_id per report); each file is served from
https://www.iemop.ph/wp-content/uploads/downloads/data/<REPORT>/<file>. About 90 days are listed:
  RTDREG  'RTD Regional Summaries', one csv per day: per 5-minute interval and region (CLUZ Luzon, CVIS
          Visayas, CMIN Mindanao) and commodity (En = energy): MKT_REQT (market requirement = demand), LOAD_BID,
          LOAD_CURTAILED, LOSSES, GENERATION, MKT_IMPORT / MKT_EXPORT (inter-regional HVDC flows)
  DIPCER  'DIPC Energy Results - Raw', one zip per hour: per 5-minute interval and resource: LMP, SCHED_MW,
          LMP_SMP (system marginal price), LMP_LOSS, LMP_CONGESTION (PHP/MWh)

Writes output/Data and Chart Outputs/philippines_power_market.xlsx:
  Daily demand   per day and region: average and peak market requirement (MW), energy (MWh = mean x 24),
                 generation (MWh), losses; Philippines totals
  Daily prices   per day and region: time-weighted average of the 5-minute system marginal price and of the
                 generation-weighted LMP (PHP/MWh); Philippines generation-weighted average
Generation by FUEL is not in these files (WESM reports by resource); the Philippines mix stays on Ember
until a resource-to-fuel map is added.

Incremental: the workbook is the history store (IEMOP lists only ~90 days, so the 1st/15th runs keep it
whole); only days not yet saved, plus REVISION_DAYS, are downloaded.

    python3 asia/PHILIPPINES_IEMOP.py
"""
import argparse
import base64
import io
import os
import sys
import time
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

AJAX = "https://www.iemop.ph/wp-admin/admin-ajax.php"
BASE = "https://www.iemop.ph"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
POSTS = {"RTDREG": 5760, "DIPCER": 5754}
REGIONS = {"CLUZ": "Luzon", "CVIS": "Visayas", "CMIN": "Mindanao", "LUZON": "Luzon", "VISAYAS": "Visayas",
           "MINDANAO": "Mindanao"}
REVISION_DAYS = 2
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "philippines_power_market.xlsx")


def out(*a):
    print(*a, flush=True)


def listing(report):
    """-> {date: [file urls]} for every file IEMOP currently lists for the report."""
    files, page = {}, 1
    while True:
        r = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                      "page": page, "post_id": POSTS[report]}, headers=H, timeout=T)
        r.raise_for_status()
        j = r.json()
        src = j.get("source") or []
        for s in src:
            path = base64.b64decode(s).decode()
            name = path.rsplit("/", 1)[-1]
            stamp = name.split("_", 1)[1].split(".")[0]
            # DIPCER_202610030000 is the hour ENDING 00:00 on the 3rd -> belongs to the 2nd
            d = pd.Timestamp(stamp[:8]).date()
            if report == "DIPCER" and stamp[8:12] == "0000":
                d -= timedelta(days=1)
            files.setdefault(d, []).append(BASE + path.split("/html", 1)[-1])
        total = int(j.get("count") or 0)
        if not src or sum(len(v) for v in files.values()) >= total or page > 200:
            return files
        page += 1


def get(url):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T)
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            if i == 3:
                out(f"  {url}: {e}")
                return None
            time.sleep(5 * (i + 1))


def demand_day(content):
    df = pd.read_csv(io.BytesIO(content))
    df.columns = [c.strip() for c in df.columns]
    df = df[df["COMMODITY_TYPE"].astype(str).str.strip().eq("En")]
    df["region"] = df["REGION_NAME"].map(REGIONS)
    for c in ("MKT_REQT", "GENERATION", "LOSSES"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["t"] = pd.to_datetime(df["TIME_INTERVAL"], format="%m/%d/%Y %I:%M:%S %p", errors="coerce")
    g = df.groupby("region")
    row = {}
    for reg, x in g:
        row[f"{reg}_demand_avg_MW"] = x["MKT_REQT"].mean()
        row[f"{reg}_demand_peak_MW"] = x["MKT_REQT"].max()
        row[f"{reg}_demand_MWh"] = x["MKT_REQT"].mean() * 24
        row[f"{reg}_generation_MWh"] = x["GENERATION"].mean() * 24
        row[f"{reg}_losses_MWh"] = x["LOSSES"].mean() * 24
    tot = df.groupby("t")["MKT_REQT"].sum()
    row["Philippines_demand_avg_MW"] = tot.mean()
    row["Philippines_demand_peak_MW"] = tot.max()
    row["Philippines_demand_MWh"] = tot.mean() * 24
    row["Intervals"] = df["t"].nunique()
    return row


def price_hour(content):
    z = zipfile.ZipFile(io.BytesIO(content))
    df = pd.read_csv(z.open(z.namelist()[0]))
    df.columns = [c.strip() for c in df.columns]
    df["region"] = df["REGION_NAME"].map(REGIONS)
    for c in ("LMP", "SCHED_MW", "LMP_SMP"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    gen = df[df["SCHED_MW"] > 0]
    out_rows = []
    for (t, reg), x in df.groupby(["TIME_INTERVAL", "region"]):
        g = gen[(gen["TIME_INTERVAL"] == t) & (gen["region"] == reg)]
        out_rows.append({"t": t, "region": reg, "SMP": x["LMP_SMP"].median(),
                         "LMP_w": (g["LMP"] * g["SCHED_MW"]).sum(), "MW": g["SCHED_MW"].sum()})
    return pd.DataFrame(out_rows)


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    old_d, old_p = read_sheet(args.out, "Daily demand"), read_sheet(args.out, "Daily prices")
    yesterday = date.today() - timedelta(days=1)
    revise = {yesterday - timedelta(days=k) for k in range(REVISION_DAYS)}

    rtd = listing("RTDREG")
    todo = sorted(d for d in rtd if d <= yesterday and (d in revise or old_d.empty or pd.Timestamp(d) not in old_d.index))
    out(f"RTDREG: {len(rtd)} days listed, fetching {len(todo)}")
    rows = {}
    for d in todo:
        c = get(rtd[d][0])
        if c:
            rows[pd.Timestamp(d)] = demand_day(c)
    dem = merge(old_d, pd.DataFrame.from_dict(rows, orient="index").round(1))

    dip = listing("DIPCER")
    todo = sorted(d for d, f in dip.items() if d <= yesterday and len(f) >= 24
                  and (d in revise or old_p.empty or pd.Timestamp(d) not in old_p.index))
    out(f"DIPCER: {len(dip)} days listed, fetching {len(todo)} (24 hourly files each)")
    rows = {}
    for n, d in enumerate(todo):
        parts = [price_hour(c) for c in (get(u) for u in dip[d]) if c]
        if not parts:
            continue
        x = pd.concat(parts)
        row = {}
        for reg, y in x.groupby("region"):
            row[f"{reg}_SMP_PHP_per_MWh"] = y["SMP"].mean()
            row[f"{reg}_LMP_genweighted_PHP_per_MWh"] = y["LMP_w"].sum() / y["MW"].sum() if y["MW"].sum() else None
        row["Philippines_LMP_genweighted_PHP_per_MWh"] = x["LMP_w"].sum() / x["MW"].sum() if x["MW"].sum() else None
        row["Intervals"] = x["t"].nunique()
        rows[pd.Timestamp(d)] = row
        if n % 10 == 0:
            out(f"  prices {d}")
    pri = merge(old_p, pd.DataFrame.from_dict(rows, orient="index").round(2))
    if dem.empty and pri.empty:
        raise SystemExit("No IEMOP data")
    dem.index.name = pri.index.name = "date"
    notes = [
        "UNITS",
        "Daily demand: per region (Luzon, Visayas, Mindanao) from the 5-minute real-time dispatch (RTD) regional "
        "summaries, energy commodity: demand_avg / demand_peak = average and maximum market requirement (MW); "
        "demand_MWh, generation_MWh, losses_MWh = mean MW x 24. Philippines_* = the three regions summed per interval. "
        "Intervals = 5-minute intervals in the day (288 = complete).",
        "Daily prices: PHP/MWh (Philippine pesos). <Region>_SMP = time-weighted average of the 5-minute system "
        "marginal price; <Region>_LMP_genweighted = locational marginal price weighted by each resource's scheduled "
        "generation; Philippines_LMP_genweighted = the same across all regions.",
        "",
        "COVERAGE",
        (f"Demand from {dem.index.min():%Y-%m-%d} to {dem.index.max():%Y-%m-%d}; " if not dem.empty else "") +
        (f"prices from {pri.index.min():%Y-%m-%d} to {pri.index.max():%Y-%m-%d}. " if not pri.empty else "") +
        "IEMOP lists about 90 days of files; older days are kept from previous runs. Grid-connected WESM regions "
        "(Mindanao joined WESM in 2023); off-grid islands are not included.",
        "",
        "SOURCE",
        "IEMOP (Independent Electricity Market Operator of the Philippines), market data: RTD Regional Summaries and "
        "DIPC Energy Results - Raw, https://www.iemop.ph/market-data/",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily demand": dem, "Daily prices": pri}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: demand {len(dem)} days, prices {len(pri)} days")
    out(dem.tail(2).T.to_string())
    out(pri.tail(2).T.to_string())


if __name__ == "__main__":
    main()
