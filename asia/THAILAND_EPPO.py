"""
Thailand electricity from EPPO (Energy Policy and Planning Office, Ministry of Energy) energy statistics,
https://www.eppo.go.th/epposite/info/stat/electricity. Found via discovery_archive/asia/SEA_DISCOVERY*.py.

EPPO publishes each statistics table as .xls, re-uploaded when it updates (the upload folder changes, so
links are read from the page each run):
  Table 5.2-4  Power Generation by Fuel (Detail), GWh, monthly: Natural Gas, Lignite, Coal, Fuel Oil,
               Diesel, Hydro, Imported Electricity (mainly Lao hydro), Renewable Energy, Total - whole
               system (EGAT, IPP, SPP, VSPP). T05_02_04-1.xls = history by month (from 1992).
  Table 5.2-2  Power Generation by Fuel Type, same system and fuels (coal and lignite combined):
               T05_02_02.xls = the current year and the two before it, which extends 5.2-4.
  Table 5.2-5  Peak, generation and load factor (EGAT system), monthly: T05_02_05-1.xls / T05_02_05.xls.

Writes output/Data and Chart Outputs/thailand_power_generation_daily.xlsx (monthly rows in the standard
generation layout, dated the 1st of each month - like the other monthly-only countries):
  Daily    <Fuel>_MWh per month: Gas, Coal (coal + lignite), Oil (fuel oil + diesel), Hydro, Other (EPPO
           'Renewable Energy': solar, wind, biomass, biogas, waste - not split by EPPO), Total_MWh (domestic);
           Imports_MWh (imported electricity) outside the total
  Peak     monthly peak demand (MW), generation (GWh) and load factor (%), EGAT system

EPPO only re-publishes whole files: they are downloaded every run (two small files) and the history is the
merged result; the files' Last-Modified dates are recorded on the Units sheet. Runs on the 1st and 15th.

    python3 asia/THAILAND_EPPO.py
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

PAGE = "https://www.eppo.go.th/epposite/info/stat/electricity"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "thailand_power_generation_daily.xlsx")
MONTHS = {m: i for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV",
                                      "DEC"], start=1)}
# header text (lower case, group + sub-header) -> standard column; first match wins
GEN_COLS = [("natural gas", "Gas"), ("lignite", "Lignite"), ("coal", "Coal"), ("fuel oil", "Fuel_oil"),
            ("diesel", "Diesel"), ("hydro", "Hydro"), ("import", "Imports"), ("renewable", "Renewable"),
            ("total", None)]
PEAK_COLS = [("peak", "Peak_MW"), ("generation", "Generation_GWh"), ("load factor", "Load_factor_pct")]


def out(*a):
    print(*a, flush=True)


def table_links():
    r = requests.get(PAGE, headers=H, timeout=T)
    r.raise_for_status()
    links = {}
    for u in re.findall(r'href="([^"]+/(T05_02_0[245](?:-1)?)\.xls)"', r.text):
        links.setdefault(u[1], u[0])
    return links


def parse_monthly(content, colmap):
    """EPPO monthly table -> DataFrame indexed by month start. Year rows ('2025') set the year; month rows
    ('JAN' in the first or second column) carry values. Column names come from the two header rows (the group
    header carried right)."""
    df = pd.read_excel(io.BytesIO(content), header=None)
    hdr_i = next(i for i in range(len(df)) if any(str(v).strip().lower() in ("date", "month")
                                                    for v in df.iloc[i, :2]))   # not the 'Monthly ...' title
    top = df.iloc[hdr_i].ffill()
    sub = df.iloc[hdr_i + 1] if hdr_i + 1 < len(df) else pd.Series(dtype=object)
    names = {}
    for j in range(df.shape[1]):
        text = f"{'' if pd.isna(top.iloc[j]) else top.iloc[j]} {'' if pd.isna(sub.get(j)) else sub.get(j)}".lower()
        # a sub-header names the column ('Lignite' under 'Coal & Lignite'); else the group header does
        s = str(sub.get(j, "")).lower() if not pd.isna(sub.get(j)) else ""
        key = next((v for k, v in colmap if k in s), None) if s and s != "nan" else None
        if key is None and "total" not in s:
            key = next((v for k, v in colmap if k in text), None)
        if key and (key not in names.values() or key in ("Fuel_oil", "Diesel")):
            names[j] = key   # first column wins (Table 5.2-5 repeats 'Generation' in kWh)
    rows, year = {}, None
    for i in range(len(df)):
        c0, c1 = str(df.iat[i, 0]).strip(), str(df.iat[i, 1]).strip() if df.shape[1] > 1 else ""
        m = re.match(r"^(\d{4})(?:\.0)?\b", c0)
        if m and 1980 < int(m.group(1)) < 2100:
            year = int(m.group(1))
        mon = next((MONTHS[x[:3].upper()] for x in (c0, c1) if x[:3].upper() in MONTHS), None)
        if year and mon:
            rows[pd.Timestamp(year, mon, 1)] = {k: pd.to_numeric(df.iat[i, j], errors="coerce")
                                                for j, k in names.items()}
    return pd.DataFrame.from_dict(rows, orient="index").sort_index()


def fetch(links, keys, colmap):
    frames, stamps = [], []
    for key in keys:   # history first, then the current-year file overrides it
        if key not in links:
            out(f"  {key}: no link on the page")
            continue
        r = requests.get(links[key], headers=H, timeout=T)
        r.raise_for_status()
        f = parse_monthly(r.content, colmap)
        out(f"  {key}: {len(f)} months {f.index.min():%Y-%m}..{f.index.max():%Y-%m}" if len(f) else f"  {key}: empty")
        if f.empty:   # layout changed: show it in the log
            out(pd.read_excel(io.BytesIO(r.content), header=None).head(30).to_string(max_cols=14, max_colwidth=16))
        frames.append(f)
        stamps.append(f"{links[key].rsplit('/', 1)[-1]} (Last-Modified {r.headers.get('last-modified')})")
    if not frames:
        return pd.DataFrame(), stamps
    d = pd.concat(frames)
    return d[~d.index.duplicated(keep="last")].sort_index(), stamps


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    links = table_links()
    out(f"links: {links}")
    gen, s1 = fetch(links, ("T05_02_04-1", "T05_02_02"), GEN_COLS)
    peak, s2 = fetch(links, ("T05_02_05-1", "T05_02_05"), PEAK_COLS)
    if gen.empty:
        raise SystemExit("No EPPO generation table")
    g = gen.apply(pd.to_numeric, errors="coerce").fillna(0) * 1000   # GWh -> MWh
    daily = pd.DataFrame({"Hydro_MWh": g.get("Hydro", 0), "Gas_MWh": g.get("Gas", 0),
                          "Coal_MWh": g.get("Coal", 0) + g.get("Lignite", 0),
                          "Oil_MWh": g.get("Fuel_oil", 0) + g.get("Diesel", 0),
                          "Other_MWh": g.get("Renewable", 0)}, index=g.index)
    daily["Total_MWh"] = daily.sum(axis=1)
    daily["Imports_MWh"] = g.get("Imports", 0)
    daily = daily[(daily.index >= "2010-01-01") & (daily["Total_MWh"] > 0)].round(0)
    daily.index.name = "date"
    peak = peak[peak.index >= "2010-01-01"].dropna(how="all") if not peak.empty else peak
    peak.index.name = "date"
    out(daily.tail(4).to_string())
    out((daily.resample("YS").sum() / 1e6).round(1).tail(6).to_string())
    notes = [
        "UNITS",
        "Daily: one row per MONTH (dated the 1st), MWh in the month (EPPO publishes GWh, x1,000). Gas, Coal (coal + "
        "lignite), Oil (fuel oil + diesel), Hydro, Other = EPPO 'Renewable "
        "Energy' (solar, wind, biomass, biogas, waste - EPPO does not split it). Total_MWh = domestic generation (sum "
        "of those). Imports_MWh = imported electricity (Lao PDR hydro, some Malaysia), not in Total_MWh.",
        "Peak: monthly peak demand (MW), generation (GWh) and load factor (%), EGAT system (Table 5.2-5).",
        "",
        "COVERAGE",
        f"Monthly from {daily.index.min():%Y-%m} to {daily.index.max():%Y-%m}, whole Thai system (EGAT, IPP, SPP and "
        "VSPP generation), EPPO Table 5.2-4 'Power Generation by Fuel (Detail)', extended with Table 5.2-2 for the "
        "current year. EPPO updates about two months "
        "after the month.",
        "Files read this run: " + "; ".join(s1 + s2),
        "",
        "SOURCE",
        "EPPO (Energy Policy and Planning Office, Ministry of Energy, Thailand), Electricity Statistics: "
        "https://www.eppo.go.th/epposite/info/stat/electricity (tables 5.2-4 and 5.2-5; data from EGAT, PEA, MEA).",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Peak": peak}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} months {daily.index.min():%Y-%m}..{daily.index.max():%Y-%m}")


if __name__ == "__main__":
    main()
