"""
Ireland (Republic) and Northern Ireland power from EirGrid's own "System and Renewable Data Reports"
  https://www.eirgrid.ie/grid/system-and-renewable-data-reports
(EirGrid / SONI, free, no key). Two whole-file publications are read:

  System-Data-Qtr-Hourly-<year>[-Vn].xlsx   quarter-hourly system data (MW) for NI, IE (Republic) and AI (all island):
        generation, demand, wind, solar, hydro, batteries, interconnector flows (Moyle; EWIC, Greenlink), SNSP
  System-and-Renewable-Data-Summary-Report-Vn.xlsx   monthly system data since 2014 (generation, demand, wind, solar per
        jurisdiction) and the rolling-12-month Fuel Mix & CO2 table

  output/Data and Chart Outputs/ireland_eirgrid_system_data.xlsx
    sheet "Daily"   : date, GWh per UTC day (mean MW x 24 over the day's quarter-hours, >= 92 of 96 required):
                      IE_Generation, IE_Demand, IE_Wind, IE_Solar, IE_Hydro, IE_Batteries, IE_EWIC, IE_Greenlink
                      (interconnector columns: positive = import into Ireland), IE_Thermal_other (= generation - wind -
                      solar - hydro; gas, oil, peat/coal and other thermal), plus the same NI_* and AI_* totals
    sheet "Monthly" : month, GWh per month from the Summary report: <region>_<Generation|Demand|Wind|Solar> for
                      Ireland, Northern Ireland and All Island
    sheet "Fuel mix": the Summary report's Fuel Mix & CO2 sheet as published (rolling 12 months, GWh and %)
    sheet "Releases": file name and Last-Modified of each source file (a file is re-read only when it changes)
    sheet "Units"   : source and definitions

Sign convention for interconnectors is as published by EirGrid; this script reports the series as given and the Units sheet
notes it. Whole-file source: a file is downloaded only when its Last-Modified differs from the one recorded on "Releases".

Usage: python3 EIRGRID_IRELAND_DAILY.py [--out-dir DIR] [--force]
"""
import argparse
import io
import os
import re
import sys
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "ireland_eirgrid_system_data.xlsx"
PAGE = "https://www.eirgrid.ie/grid/system-and-renewable-data-reports"
H = {"User-Agent": "Mozilla/5.0", "Accept": "*/*"}
MIN_QH = 92

QH_COLS = {   # quarter-hourly column -> output name (MW)
    "IE Generation": "IE_Generation", "IE Demand": "IE_Demand", "IE Wind Generation": "IE_Wind",
    "IE Solar Generation": "IE_Solar", "IE Hydro": "IE_Hydro", "IE Batteries": "IE_Batteries", "EWIC I/C": "IE_EWIC",
    "Greenlink I/C": "IE_Greenlink",
    "NI Generation": "NI_Generation", "NI Demand": "NI_Demand", "NI Wind Generation": "NI_Wind",
    "NI Solar Generation": "NI_Solar", "NI Batteries": "NI_Batteries", "Moyle I/C": "NI_Moyle",
    "AI Generation": "AI_Generation", "AI Demand": "AI_Demand", "AI Wind Generation": "AI_Wind",
    "AI Solar Generation": "AI_Solar", "AI Hydro": "AI_Hydro",
}


def links():
    page = requests.get(PAGE, headers=H, timeout=120).text
    hrefs = sorted(set(re.findall(r'href="([^"]+/publications/[^"]+\.xlsx)"', page)))
    qh = [u for u in hrefs if "System-Data-Qtr-Hourly" in u]
    # a year may be published more than once; keep the latest version of each year
    best = {}
    for u in qh:
        m = re.search(r"Qtr-Hourly-(\d{4})(?:-V(\d+))?", u)
        if m:
            y, v = int(m.group(1)), int(m.group(2) or 0)
            if y not in best or v > best[y][0]:
                best[y] = (v, u)
    # older years are still hosted but no longer linked from the page: try the plain per-year name
    base = "https://cms.eirgrid.ie/sites/default/files/publications/"
    for y in range(2014, datetime.now(timezone.utc).year):
        if y not in best:
            u = f"{base}System-Data-Qtr-Hourly-{y}.xlsx"
            try:
                if requests.head(u, headers=H, timeout=30, allow_redirects=True).status_code == 200:
                    best[y] = (0, u)
            except requests.RequestException:
                pass
    summary = [u for u in hrefs if "System-and-Renewable-Data-Summary-Report" in u]
    return [u for _, u in sorted(best.values(), key=lambda t: t[1])], (summary[-1] if summary else None)


def last_modified(u):
    try:
        return requests.head(u, headers=H, timeout=60, allow_redirects=True).headers.get("Last-Modified", "")
    except requests.RequestException:
        return ""


def daily_from_qh(content):
    d = pd.read_excel(io.BytesIO(content), sheet_name=0)
    d = d.rename(columns=lambda c: str(c).strip())
    d["DateTime"] = pd.to_datetime(d["DateTime"], errors="coerce")
    d = d.dropna(subset=["DateTime"])
    off = pd.to_numeric(d.get("GMT Offset", 0), errors="coerce").fillna(0)   # clock hours ahead of GMT in BST
    d["utc"] = d["DateTime"] - pd.to_timedelta(off, unit="h")
    cols = {k: v for k, v in QH_COLS.items() if k in d.columns}
    x = d[["utc"] + list(cols)].rename(columns=cols).set_index("utc")
    x = x.apply(pd.to_numeric, errors="coerce")
    day = x.index.floor("D")
    n = x.notna().groupby(day).sum()
    mean = x.groupby(day).mean()
    out = (mean * 24 / 1000.0).where(n >= MIN_QH)   # MW -> GWh per day
    out["IE_Thermal_other"] = out.get("IE_Generation") - out.get("IE_Wind") - out.get("IE_Solar") - out.get("IE_Hydro").fillna(0)
    out.index.name = "date"
    return out.round(3)


def monthly_from_summary(content):
    s = pd.read_excel(io.BytesIO(content), sheet_name="System Data Summary", header=None)
    dates = s.iloc[1, 3:]
    keep = [i for i, v in zip(dates.index, dates.values) if isinstance(v, (datetime, pd.Timestamp))]
    region = s.iloc[:, 0].ffill()
    item = s.iloc[:, 1].ffill()
    stat = s.iloc[:, 2]
    out = {}
    rename = {"System Generation": "Generation", "System Demand": "Demand", "Wind Generation": "Wind", "Solar Generation": "Solar"}
    reg = {"Ireland": "IE", "Northern Ireland": "NI", "All Island": "AI"}
    for r in range(2, len(s)):
        if not (isinstance(stat.iloc[r], str) and stat.iloc[r].strip().startswith("Total")):
            continue
        name = rename.get(str(item.iloc[r]).strip())
        rg = reg.get(str(region.iloc[r]).strip())
        if not name or not rg:
            continue
        vals = pd.to_numeric(s.iloc[r, keep], errors="coerce")
        out[f"{rg}_{name}"] = pd.Series(vals.values, index=pd.to_datetime([s.iloc[1, i] for i in keep]))
    m = pd.DataFrame(out).sort_index().dropna(how="all")
    m.index = m.index.strftime("%Y-%m")
    m.index.name = "Month"
    return m.round(2)


def read_old(path, sheet, index):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()
    if index in d:
        if index == "date":
            d["date"] = pd.to_datetime(d["date"], errors="coerce")
        d = d.dropna(subset=[index]).set_index(index)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    qh_urls, summary_url = links()
    print("quarter-hourly files:", qh_urls, "\nsummary:", summary_url, flush=True)
    old_rel = read_old(path, "Releases", "File")
    seen = dict(zip(old_rel.index, old_rel.get("Last-Modified", []))) if len(old_rel) else {}
    daily = read_old(path, "Daily", "date")
    monthly = read_old(path, "Monthly", "Month")
    fuel = read_old(path, "Fuel mix", "__none__") if os.path.exists(path) else pd.DataFrame()
    rel = {}
    changed = False
    for u in qh_urls:
        name = u.rsplit("/", 1)[-1]
        lm = last_modified(u)
        rel[name] = lm
        if not args.force and seen.get(name) == lm and lm:
            print(f"  {name}: unchanged ({lm})", flush=True)
            continue
        print(f"  {name}: downloading", flush=True)
        r = requests.get(u, headers=H, timeout=300)
        r.raise_for_status()
        new = daily_from_qh(r.content)
        daily = pd.concat([daily[~daily.index.isin(new.index)], new]).sort_index() if len(daily) else new
        changed = True
        print(f"    {len(new)} days {new.index.min():%Y-%m-%d} .. {new.index.max():%Y-%m-%d}; columns {list(new.columns)}", flush=True)
    if summary_url:
        name = summary_url.rsplit("/", 1)[-1]
        lm = last_modified(summary_url)
        rel[name] = lm
        if args.force or seen.get(name) != lm or not lm or monthly.empty:
            print(f"  {name}: downloading", flush=True)
            r = requests.get(summary_url, headers=H, timeout=300)
            r.raise_for_status()
            monthly = monthly_from_summary(r.content)
            fuel = pd.read_excel(io.BytesIO(r.content), sheet_name="Fuel Mix & CO2", header=None).dropna(how="all").dropna(how="all", axis=1)
            fuel = fuel.astype(object).where(fuel.notna(), None)
            changed = True
            print(f"    monthly {len(monthly)} months {monthly.index.min()} .. {monthly.index.max()}; columns {list(monthly.columns)}", flush=True)
        else:
            print(f"  {name}: unchanged ({lm})", flush=True)
    if daily.empty and monthly.empty:
        raise SystemExit("no EirGrid data read")
    if not changed and os.path.exists(path):
        print("nothing new; workbook left as is")
        return
    lines = ["Ireland and Northern Ireland - EirGrid / SONI system and renewable data", "",
             "Source", f"EirGrid, System and Renewable Data Reports: {PAGE} (quarter-hourly System Data workbooks and the Summary report). "
             "Free, no key. TSO's own data.",
             "", "Units and definitions",
             "Daily: GWh per UTC day = mean MW over the day's quarter-hours x 24 (days with fewer than 92 of 96 quarter-hours dropped). IE = Republic "
             "of Ireland, NI = Northern Ireland, AI = All Island. IE_Generation is EirGrid's system generation for the jurisdiction (metered and "
             "embedded); IE_Wind and IE_Solar include distribution-connected plants; IE_Thermal_other = generation - wind - solar - hydro "
             "(gas, oil, peat/coal and other thermal in one figure). EWIC / Greenlink / Moyle are interconnector flows as published "
             "(check the sign against the net import you expect before using).",
             "Monthly: GWh per month from the Summary report's System Data Summary (since 2014). Fuel mix: the Summary report's rolling "
             "12-month Fuel Mix & CO2 table, as published.",
             "Whole-file sources: a file is re-read only when its Last-Modified changes (see Releases).",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; daily {len(daily)} days"
             + (f", {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}" if len(daily) else "")]
    releases = pd.DataFrame({"File": list(rel), "Last-Modified": list(rel.values())}).set_index("File")
    sheets = {"Daily": daily, "Monthly": monthly, "Fuel mix": fuel, "Releases": releases}
    xlsx_notes.write_workbook(path, {k: v for k, v in sheets.items() if v is not None and len(v)}, lines,
                              {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
