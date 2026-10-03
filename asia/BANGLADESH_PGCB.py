"""
Bangladesh electricity generation by fuel from PGCB (Power Grid Bangladesh PLC, the national grid
operator / NLDC), hourly generation table: https://erp.powergrid.gov.bd/w/generations/view_generations
Found via discovery_archive/subcontinent/SEA_DISCOVERY_SOUTH_ASIA*.py.

The table (HTML, newest first, paged with ?page=N, back to 2015) gives per hour: generation (MW),
demand, load shed, and generation by source - Gas, Liquid Fuel, Coal, Hydro, Solar, Wind - plus imports
from India (Bheramara HVDC, Tripura, Adani Godda) and Nepal. The site's certificate chain is incomplete,
so requests are made with verification off (read-only public data).

Writes output/Data and Chart Outputs/bangladesh_power_generation_daily.xlsx:
  Daily    standard layout, MWh per day by fuel (mean of the day's hourly MW x 24); Oil = liquid fuel;
           Imports_MWh (India + Nepal) kept outside Total_MWh (domestic generation)
  Demand   daily peak and average generation (MW, served demand) and peak / total load shed
  Hourly   the hourly rows, last 60 days

Incremental: the Daily sheet is the history store; pages are read from the newest back until they reach
days already saved (less REVISION_DAYS) or DATA_START. Runs on the 1st and 15th.

    python3 asia/BANGLADESH_PGCB.py
"""
import argparse
import os
import re
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
URL = "https://erp.powergrid.gov.bd/w/generations/view_generations"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (15, 90)
DATA_START = date(2021, 1, 1)
REVISION_DAYS = 3
MAX_PAGES = 3000
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "bangladesh_power_generation_daily.xlsx")
# header keyword (lower case) -> column; first match wins, checked against the flattened header text
COLMAP = [("date", "time"), ("time", "time"), ("generation", "Generation_MW"), ("demand", "Demand_MW"), ("loadshed", "Loadshed_MW"),
          ("load shed", "Loadshed_MW"), ("gas", "Gas"), ("liquid", "Oil"), ("coal", "Coal"), ("hydro", "Hydro"),
          ("solar", "Solar"), ("wind", "Wind"), ("bheramara", "Import_Bheramara"), ("tripura", "Import_Tripura"),
          ("adani", "Import_Adani"), ("nepal", "Import_Nepal"), ("india", "Import_India"), ("remark", "Remarks")]
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil"]


def out(*a):
    print(*a, flush=True)


def page(n):
    for i in range(4):
        try:
            r = requests.get(URL, params={"page": n}, headers=H, timeout=T, verify=False)
            r.raise_for_status()
            return parse(r.text)
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                out(f"  page {n} failed: {e}")
                return pd.DataFrame()
            time.sleep(5 * (i + 1))


def cells(fragment, tag):
    return [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", c)).strip()
            for c in re.findall(rf"(?is)<{tag}[^>]*>(.*?)</{tag}>", fragment)]


def header_names(html):
    """Header labels -> standard names, with India's sub-headers (Bheramara HVDC, Tripura, Adani) in place of
    'India'."""
    thead = re.search(r"(?is)<thead.*?</thead>", html)
    labels = cells(thead.group(0), "th") if thead else []
    subs = [x for x in labels if any(k in x.lower() for k in ("bheramara", "tripura", "adani"))]
    flat = []
    for x in labels:
        if x in subs:
            continue
        flat.extend(subs if x.lower() == "india" and subs else [x])
    return [next((v for k, v in COLMAP if k in x.lower()), x) for x in flat]


def parse(html):
    """PGCB hourly table -> rows (time + MW columns). The page comments out the Demand and Loadshed cells in each
    row (while the header still lists them), so HTML comments are removed and, when a row has two cells fewer
    than the header, those two columns are dropped from the header."""
    html = re.sub(r"(?s)<!--.*?-->", "", html)
    names = header_names(html) or ["time", "Generation_MW", "Demand_MW", "Loadshed_MW", "Gas", "Oil", "Coal",
                                   "Hydro", "Solar", "Wind", "Import_Bheramara", "Import_Tripura", "Import_Adani",
                                   "Import_Nepal", "Remarks"]
    if names and names[0] == "time" and len(names) > 1 and names[1] in ("time", "Time"):
        names = names[1:]
    body = re.search(r"(?is)<tbody.*?</tbody>", html)
    rows, bad = [], []
    for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", body.group(0) if body else html):
        c = cells(tr, "td")
        if len(c) < 5 or not re.match(r"\d\d-\d\d-\d{4}$", c[0]):
            continue
        vals = [f"{c[0]} {c[1]}"] + c[2:]    # date and time sit in separate cells
        cols = names
        if len(vals) == len(names) - 2:
            cols = [n for n in names if n not in ("Demand_MW", "Loadshed_MW")]
        if len(vals) != len(cols):
            bad.append((len(vals), len(cols), vals[:4]))
            continue
        rows.append(dict(zip(cols, vals)))
    if bad and not rows:
        print(f"  unparsed layout: header {names}; first rows (cells, header cols, start) {bad[:2]}", flush=True)
    if not rows:
        return pd.DataFrame()
    t = pd.DataFrame(rows)
    t["time"] = pd.to_datetime(t["time"], format="%d-%m-%Y %H:%M:%S", errors="coerce")
    # keep the on-the-hour rows; an 'Evening Peak' remark row can add an off-hour reading or repeat an hour
    t = t.dropna(subset=["time"])
    t = t[t["time"].dt.minute.eq(0)].drop_duplicates("time")
    num = [x for x in t.columns if x not in ("time", "Remarks")]
    t[num] = t[num].apply(lambda s: pd.to_numeric(s.astype(str).str.replace(",", ""), errors="coerce"))
    return t[["time"] + num]


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
    old, old_dem, old_h = (read_sheet(args.out, s) for s in ("Daily", "Demand", "Hourly"))
    # read back to the last saved day; all the way to DATA_START while the saved history starts later than that
    stop = (old.index.max().date() - timedelta(days=REVISION_DAYS)
            if not old.empty and old.index.min().date() <= DATA_START + timedelta(days=7) else DATA_START)
    out(f"{len(old)} days saved; reading pages back to {stop}")
    frames = []
    empty = 0
    for n in range(1, MAX_PAGES + 1):
        p = page(n)
        if p.empty:
            empty += 1
            out(f"  page {n}: no rows")
            if empty >= 5:
                out("  5 pages without rows - stopping")
                break
            continue
        empty = 0
        mid = p["time"].median()   # a mistyped date (e.g. 2004 for 2024) must not end the walk or enter the data
        typo = (p["time"] - mid).abs() > pd.Timedelta(days=10)
        if typo.any():
            out(f"  page {n}: dropped {int(typo.sum())} row(s) dated far from the page: {p.loc[typo, 'time'].tolist()[:3]}")
        p = p[~typo]
        frames.append(p)
        oldest = p["time"].min()
        if n % 50 == 0:
            out(f"  page {n}: back to {oldest}")
        if oldest.date() < stop:
            break
        time.sleep(0.5)
    if not frames:
        raise SystemExit("No PGCB rows")
    h = pd.concat(frames).drop_duplicates("time").set_index("time").sort_index()
    h = h[h.index.date >= stop]
    h["Imports"] = h[[c for c in h.columns if c.startswith("Import_")]].sum(axis=1, min_count=1)
    day = h.index.normalize()
    g = h.groupby(day)
    mwh = g[[c for c in FUELS + ["Imports"] if c in h]].mean().mul(24)
    new = pd.DataFrame({f"{f}_MWh": mwh[f] if f in mwh else 0.0 for f in FUELS}, index=mwh.index)
    new["Other_MWh"] = 0.0
    new["Total_MWh"] = new.sum(axis=1)
    new["Imports_MWh"] = mwh.get("Imports")
    new["Hours"] = g.size()
    new = new[new["Hours"] >= 20].round(1)   # the newest, part-published day waits for the next run
    dem = pd.DataFrame({"Demand_peak_MW": g["Generation_MW"].max(), "Demand_avg_MW": g["Generation_MW"].mean().round(0),
                        "Loadshed_peak_MW": g["Loadshed_MW"].max() if "Loadshed_MW" in h else None,
                        "Loadshed_MWh": g["Loadshed_MW"].mean().mul(24).round(0) if "Loadshed_MW" in h else None})
    dem = dem[dem.index.isin(new.index)]
    daily, demand = merge(old, new), merge(old_dem, dem)
    hourly = merge(old_h, h.drop(columns=[c for c in h.columns if c == "Remarks"]))
    hourly = hourly[hourly.index >= hourly.index.max() - pd.Timedelta(days=60)]
    daily.index.name = demand.index.name = "date"
    hourly.index.name = "time"
    notes = [
        "UNITS",
        "Daily: MWh per day = mean of the day's hourly MW x 24, by source: Gas, Oil (PGCB 'Liquid Fuel': HFO and "
        "diesel), Coal, Hydro (Kaptai), Solar, Wind; Other_MWh is zero (standard layout). Total_MWh = domestic "
        "generation (sum of those). Imports_MWh = India (Bheramara HVDC, Tripura, Adani Godda) + Nepal, NOT in "
        "Total_MWh. Hours = hourly rows for the day (24 = complete; days with under 20 are left for the next run).",
        "Demand: Demand_peak_MW / Demand_avg_MW = PGCB 'Generation' column (served demand incl. imports), daily max "
        "and mean; Loadshed_peak_MW and Loadshed_MWh (mean x 24) = load shedding.",
        "Hourly: the hourly rows as published, last 60 days.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. National grid (PGCB / NLDC); "
        "rooftop and off-grid solar are not included.",
        "",
        "SOURCE",
        "PGCB (Power Grid Bangladesh PLC), hourly generation: https://erp.powergrid.gov.bd/w/generations/view_generations",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Demand": demand, "Hourly": hourly}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
