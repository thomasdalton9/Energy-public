"""
Singapore wholesale electricity price: the Uniform Singapore Energy Price (USEP), half-hourly, from EMC's
NEMS market data download (https://www.nems.emcsg.com/nems-prices, 'USEP and Demand Forecast' = value 1;
same download service asia/SINGAPORE_POWER.py uses for metered generation, value 16). No key:
  /api/sitecore/DataSync/DataDownloadByYear?value=1&year=YYYY&tpcValue=1   whole year as a zip (one CSV per
        month; redirects to EMC's blob store). Five-year rolling window.
  /api/sitecore/DataSync/DataDownload?value=1&fromDate=&toDate=&tpcValue=1  CSV, <= 31 days per request, for
        the days after the latest yearly zip.

Writes output/Data and Chart Outputs/singapore_power_prices.xlsx:
  Daily        USEP daily average (time-weighted mean of the 48 half-hours), max and min, SGD/MWh; daily average
               demand (MW) where the CSV gives it; Periods = half-hours present
  Monthly      monthly average USEP (mean of the half-hours) and demand-weighted average where demand is given
  Half-hourly  the last 120 days

Incremental: the Daily sheet is the history store; only years holding a missing day are downloaded as zips
(plus the days after the latest zip, and the last REVISION_DAYS again). Runs on the 1st and 15th.

    python3 asia/SINGAPORE_USEP.py [--out PATH]
"""
import argparse
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

NEMS = "https://www.nems.emcsg.com"
NEMS_PAGE = f"{NEMS}/nems-prices"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (15, 180)
DATA_START = date(2021, 1, 1)
REVISION_DAYS = 18
HH_DAYS = 120
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "singapore_power_prices.xlsx")


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                raise requests.RequestException(f"HTTP {r.status_code}")
            return r
        except requests.RequestException as e:
            if i == 3:
                raise
            out(f"  retry {url[:90]}: {e}")
            time.sleep(5 * (i + 1))


def parse_dates(col):
    """NEMS writes dates as '01-Jan-2025'. Parse that format explicitly; any value it does not fit is tried as ISO
    (YYYY-MM-DD), then day-first (DD/MM/YYYY) - never 'mixed' with dayfirst, which would swap ISO month and day."""
    s = col.astype(str).str.strip()
    d = pd.to_datetime(s, format="%d-%b-%Y", errors="coerce")
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y"):
        miss = d.isna()
        if not miss.any():
            break
        d[miss] = pd.to_datetime(s[miss], format=fmt, errors="coerce")
    if d.isna().any():
        out(f"  {int(d.isna().sum())} USEP rows with unreadable dates dropped, e.g. {s[d.isna()].iloc[0]!r}")
    return d


def parse_csv(text):
    """NEMS 'USEP and Demand Forecast' CSV -> frame (time, period, USEP SGD/MWh, demand MW)."""
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip().upper() for c in df.columns]
    usep = next(c for c in df.columns if "USEP" in c)
    dem = next((c for c in df.columns if "DEMAND" in c), None)
    d = pd.DataFrame({"date": parse_dates(df["DATE"]),
                      "period": pd.to_numeric(df["PERIOD"], errors="coerce"),
                      "USEP_SGD_per_MWh": pd.to_numeric(df[usep], errors="coerce")})
    if dem:
        d["Demand_MW"] = pd.to_numeric(df[dem], errors="coerce")
    d = d.dropna(subset=["date", "period", "USEP_SGD_per_MWh"])
    d["time"] = d["date"] + pd.to_timedelta((d["period"] - 1) * 30, unit="min")
    return d


def year_zip(year):
    r = get(f"{NEMS}/api/sitecore/DataSync/DataDownloadByYear", params={"value": "1", "year": str(year),
                                                                       "tpcValue": "1"})
    if r.status_code != 200 or r.content[:2] != b"PK":
        out(f"USEP {year}: no zip (HTTP {r.status_code}, {len(r.content)} bytes)")
        return pd.DataFrame()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    parts = [parse_csv(z.read(n).decode("utf-8-sig", errors="replace")) for n in z.namelist() if n.lower().endswith(".csv")]
    df = pd.concat(parts) if parts else pd.DataFrame()
    out(f"USEP {year}: {len(z.namelist())} files -> " +
        (f"{df['date'].min():%Y-%m-%d}..{df['date'].max():%Y-%m-%d}" if not df.empty else "empty"))
    return df


def date_range(d0, d1):
    r = get(f"{NEMS}/api/sitecore/DataSync/DataDownload", params={
        "value": "1", "fromDate": d0.isoformat(), "toDate": d1.isoformat(), "tpcValue": "1"})
    if r.status_code != 200 or "csv" not in (r.headers.get("content-type") or "").lower():
        out(f"  USEP {d0}..{d1}: HTTP {r.status_code} {r.headers.get('content-type')}")
        return pd.DataFrame()
    return parse_csv(r.content.decode("utf-8-sig", errors="replace"))


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index, errors="coerce")
    return df[df.index.notna()].sort_index()


def merge(old, new):
    if old.empty:
        return new.sort_index()
    if new.empty:
        return old
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def fetch(daily):
    yesterday = date.today() - timedelta(days=1)
    complete = set() if daily.empty else {d.date() for d in daily.index[daily["Periods"] >= 48]}
    revise = {yesterday - timedelta(days=k) for k in range(REVISION_DAYS)}
    missing = [DATA_START + timedelta(days=i) for i in range((yesterday - DATA_START).days + 1)]
    missing = [d for d in missing if d not in complete or d in revise]
    out(f"USEP: {len(complete)} complete days saved; {len(missing)} to fetch")
    if not missing:
        return pd.DataFrame()
    first = min(missing)
    parts = []
    # whole-year zips for years with old gaps (cheap: one file per year); recent days by date range
    for y in sorted({d.year for d in missing if d < yesterday - timedelta(days=60)}):
        parts.append(year_zip(y))
    got = pd.concat([p for p in parts if not p.empty]) if any(not p.empty for p in parts) else pd.DataFrame()
    d0 = max(first, (got["date"].max().date() + timedelta(days=1)) if not got.empty else first)
    d0 = min(d0, min(revise))
    while d0 <= yesterday:
        d1 = min(d0 + timedelta(days=30), yesterday)
        part = date_range(d0, d1)
        out(f"  USEP {d0}..{d1}: {0 if part.empty else part['date'].nunique()} days")
        if not part.empty:
            got = pd.concat([got, part])
        d0 = d1 + timedelta(days=1)
        time.sleep(0.5)
    if got.empty:
        return got
    return got.drop_duplicates(subset=["date", "period"], keep="last").set_index("time").sort_index()


def summarise(hh):
    g = hh["USEP_SGD_per_MWh"].groupby(hh.index.normalize())
    daily = pd.DataFrame({"USEP_avg_SGD_per_MWh": g.mean(), "USEP_max_SGD_per_MWh": g.max(),
                          "USEP_min_SGD_per_MWh": g.min()})
    if "Demand_MW" in hh:
        daily["Demand_avg_MW"] = hh["Demand_MW"].groupby(hh.index.normalize()).mean()
    daily["Periods"] = g.size()
    return daily.round(2)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    old_daily, old_hh = read_sheet(a.out, "Daily"), read_sheet(a.out, "Half-hourly")
    new_hh = fetch(old_daily)
    hh = merge(old_hh, new_hh[[c for c in ("period", "USEP_SGD_per_MWh", "Demand_MW") if c in new_hh]]
               if not new_hh.empty else new_hh)
    new_daily = summarise(new_hh) if not new_hh.empty else pd.DataFrame()
    daily = merge(old_daily, new_daily)
    if daily.empty:
        raise SystemExit("No USEP data")
    daily.index.name = "date"
    # monthly from the daily averages (all days, every half-hour weight equal)
    m = daily.groupby(daily.index.to_period("M"))
    monthly = pd.DataFrame({"USEP_avg_SGD_per_MWh": m["USEP_avg_SGD_per_MWh"].mean(),
                            "USEP_max_SGD_per_MWh": m["USEP_max_SGD_per_MWh"].max(),
                            "USEP_min_SGD_per_MWh": m["USEP_min_SGD_per_MWh"].min(), "Days": m.size()}).round(2)
    monthly.index = monthly.index.to_timestamp()
    monthly.index.name = "month"
    hh = hh[hh.index >= hh.index.max() - pd.Timedelta(days=HH_DAYS)] if not hh.empty else hh
    hh.index.name = "time"
    notes = [
        "UNITS",
        "SGD/MWh (Singapore dollars per MWh). Daily: average, maximum and minimum of the day's 48 half-hourly USEP "
        "values; Demand_avg_MW = average of the half-hourly demand in the same file (MW); Periods = half-hours present "
        "(48 for a full day). Monthly: average of the daily averages, maximum and minimum half-hour. Half-hourly: last "
        f"{HH_DAYS} days, time = start of the trading period (period 1 = 00:00-00:30 Singapore time).",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. The USEP is the demand-weighted "
        "average of the nodal energy prices in the National Electricity Market of Singapore (NEMS) for each "
        "half-hour; final prices replace provisional ones about a week after trading, so the last "
        f"{REVISION_DAYS} days are re-read each run. EMC keeps a five-year rolling window; older days are kept from "
        "previous runs.",
        "",
        "SOURCE",
        f"EMC (Energy Market Company, Singapore), NEMS market data - USEP and Demand Forecast: {NEMS_PAGE} "
        "(download service /api/sitecore/DataSync/DataDownload and DataDownloadByYear, value=1).",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    xlsx_notes.write_workbook(a.out, {"Daily": daily, "Monthly": monthly, "Half-hourly": hh}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {a.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(monthly.tail(14).to_string())


if __name__ == "__main__":
    main()
