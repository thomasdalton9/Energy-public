"""
Singapore electricity: system demand (half-hourly + daily), generation by
plant type (daily), monthly generation, annual consumption by sector and
annual fuel mix. Found via discovery_archive/asia/SINGAPORE_DISCOVERY*.py.

Sources (all free, no login, no key):
  1. EMA 'Half-hourly System Demand Data' - one .xls per week (Mon-Sun),
     listed by the servlet behind
     https://www.ema.gov.sg/resources/statistics/half-hourly-system-demand-data
     Columns per day: System Demand (Actual), NEM Demand (Actual),
     NEM Demand (Forecast), MW per half-hour. Weekly files since 2014;
     pulled from 2021. Published a few days after each week ends.
  2. EMC / NEMS market data download (https://www.nems.emcsg.com/nems-prices,
     'Metered Generation by Facility Type'): MWh injected per half-hour by
     facility type (CCGT/COGEN/TRIGEN, ST, GT, IGS, IMPORT, ESS/BATTERY,
     OTHERS). Whole years come as a zip (DataDownloadByYear), recent days via
     DataDownload (<=31 days per request). Five-year rolling window, from 2021.
  3. SingStat TableBuilder M890831 'Electricity Generation, Monthly' (GWh,
     from EMA), https://tablebuilder.singstat.gov.sg/table/TS/M890831
  4. EMA Singapore Energy Statistics (SES) tidy workbook: T3.2 electricity
     consumption by sector (annual GWh) and T2.2 fuel mix for electricity
     generation (annual %), https://www.ema.gov.sg/resources/singapore-energy-statistics
     (the latest year is part-year - SES 2025 has Jan-Jun 2025).

Monthly electricity CONSUMPTION by sector is not published (EMA/SES give it
annually); the monthly/daily demand series here are system demand and
generation.

Incremental: the half-hourly demand and daily generation archives are read
back from the workbook and only missing days are fetched (EMA weekly files
covering a missing day; NEMS years containing a missing day, plus the days
after the latest yearly zip). The monthly/annual tables are small and are
re-read each run.

    python3 asia/SINGAPORE_POWER.py --out "output/Data and Chart Outputs/singapore_power.xlsx"
"""
import argparse
import io
import os
import re
import sys
import time
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (15, 180)
DATA_START = date(2021, 1, 1)
EMA = "https://www.ema.gov.sg"
EMA_LIST = (f"{EMA}/bin/corporate-site/half-hourly-list?csvPath=/content/dam/corporate/statistics/"
            "half-hourly-data.csv&assetFolderPath=/content/dam/corporate/resources/statistics/half-hourly-data"
            "&startYear=2014")
EMA_PAGE = f"{EMA}/resources/statistics/half-hourly-system-demand-data"
NEMS = "https://www.nems.emcsg.com"
NEMS_PAGE = f"{NEMS}/nems-prices"
SINGSTAT = "https://tablebuilder.singstat.gov.sg/api/table/tabledata/M890831"
SES_PAGE = f"{EMA}/resources/singapore-energy-statistics/chapter3"
SES_FALLBACK = (f"{EMA}/content/dam/corporate/resources/singapore-energy-statistics/excel/"
                "SES_tidy.xlsx.coredownload.xlsx")

HH_SHEET, DAILY_SHEET, GEN_SHEET = "Half-hourly demand", "Daily demand", "Daily generation by type"
MONTHLY_SHEET, CONS_SHEET, MIX_SHEET = "Monthly generation", "Annual consumption", "Annual fuel mix"
DEMAND_COLS = {"system demand": "System_Demand_Actual_MW", "nem demand (actual)": "NEM_Demand_Actual_MW",
               "nem demand (forecast)": "NEM_Demand_Forecast_MW"}
GEN_TYPES = {"CCGT/COGEN/TRIGEN": "CCGT_Cogen_Trigen", "ST": "Steam_turbine", "GT": "Gas_turbine_OCGT",
             "IGS": "Solar_IGS", "IMPORT": "Imports", "ESS": "Battery_ESS", "BATTERY": "Battery_ESS",
             "OTHERS": "Other"}


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
            out(f"  retry {url[:100]}: {e}")
            time.sleep(5 * (i + 1))


def read_sheet(path, sheet, index_col=0):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


# ------------------------------------------------------------------ EMA half-hourly system demand

def parse_ema_week(content):
    """One weekly EMA .xls -> half-hourly frame indexed by period-ending timestamp, plus its footnotes."""
    raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None)
    date_row = next(i for i in range(10) if str(raw.iloc[i, 0]).strip().lower() == "date")
    date_cols = [c for c in range(1, raw.shape[1]) if isinstance(raw.iloc[date_row, c], (pd.Timestamp,))
                 or re.match(r"\d{4}-\d\d-\d\d|\d\d/\d\d/\d{4}", str(raw.iloc[date_row, c]))]
    hdr_row = next((i for i in range(date_row, date_row + 6)
                    if any("demand" in str(v).lower() for v in raw.iloc[i, 1:])), None)
    data_rows = [i for i in range(raw.shape[0]) if re.fullmatch(r"\d\d:\d\d(:\d\d)?", str(raw.iloc[i, 0]).strip())]
    notes = [str(v) for v in raw.iloc[data_rows[-1] + 1:, 0].dropna() if str(v).strip()]
    frames = []
    for k, c in enumerate(date_cols):
        d = pd.to_datetime(raw.iloc[date_row, c], dayfirst=True).normalize()
        end = date_cols[k + 1] if k + 1 < len(date_cols) else raw.shape[1]
        cols = {}
        for cc in range(c, end):
            name = str(raw.iloc[hdr_row, cc]).lower() if hdr_row is not None else "system demand"
            std = next((v for key, v in DEMAND_COLS.items() if name.startswith(key) or key in name), None)
            if std:
                cols[std] = pd.to_numeric(raw.iloc[data_rows, cc], errors="coerce").values
        idx = [d + pd.Timedelta(minutes=30 * (j + 1)) for j in range(len(data_rows))]
        frames.append(pd.DataFrame(cols, index=pd.DatetimeIndex(idx, name="Period_End")))
    return pd.concat(frames).sort_index(), notes


def days_of(hh):
    """Trading days fully present (48 periods) in a half-hourly frame (period ending 00:00 belongs to the day before)."""
    if hh.empty:
        return set()
    day = (hh.index - pd.Timedelta(minutes=1)).normalize()
    n = hh["System_Demand_Actual_MW"].notna().groupby(day).sum()
    return {d.date() for d in n[n >= 46].index}


def update_ema(hh):
    have = days_of(hh)
    r = get(EMA_LIST)
    items = r.json()["items"]
    out(f"EMA: {len(items)} weekly files listed ({items[-1]['Date']} .. {items[0]['Date']})")
    todo = []
    for it in items:
        start = pd.to_datetime(it["Date"], format="%d %b %Y").date()
        week = [start + timedelta(days=i) for i in range(7)]
        week = [d for d in week if DATA_START <= d < date.today()]
        if week and any(d not in have for d in week) and it.get("Excel"):
            todo.append((start, it["Excel"]))
    out(f"EMA: {len(todo)} weekly files to fetch")
    notes, new = [], []
    for k, (start, link) in enumerate(sorted(todo)):
        try:
            x = get(EMA + link if link.startswith("/") else link)
            week, notes = parse_ema_week(x.content)
            new.append(week)
        except Exception as e:  # noqa: BLE001 - keep the rest of the backfill going
            out(f"  {start}: failed ({e!r})")
        if k % 25 == 0:
            out(f"  fetched {k + 1}/{len(todo)} (week of {start})")
        time.sleep(0.3)
    if new:
        add = pd.concat(new)
        hh = pd.concat([hh[~hh.index.isin(add.index)], add]).sort_index() if not hh.empty else add
    hh = hh[hh.index > pd.Timestamp(DATA_START)]
    return hh, notes


def daily_demand(hh):
    day = (hh.index - pd.Timedelta(minutes=1)).normalize()
    g = hh.groupby(day)
    d = pd.DataFrame({
        "System_Demand_Avg_MW": g["System_Demand_Actual_MW"].mean(),
        "System_Demand_Peak_MW": g["System_Demand_Actual_MW"].max(),
        "System_Demand_Min_MW": g["System_Demand_Actual_MW"].min(),
        "System_Demand_GWh": g["System_Demand_Actual_MW"].sum() * 0.5 / 1000,
        "NEM_Demand_Actual_Avg_MW": g["NEM_Demand_Actual_MW"].mean() if "NEM_Demand_Actual_MW" in hh else None,
        "NEM_Demand_Forecast_Avg_MW": g["NEM_Demand_Forecast_MW"].mean() if "NEM_Demand_Forecast_MW" in hh else None,
        "Periods": g["System_Demand_Actual_MW"].count(),
    })
    d.index.name = "Date"
    return d.round(3)


# ------------------------------------------------------------------ NEMS metered generation by facility type

def parse_mg_csv(text):
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip().upper() for c in df.columns]
    df["DATE"] = pd.to_datetime(df["DATE"], format="%d-%b-%Y")
    for c in ("GROSS INJECTION (MWH)", "NET INJECTION (MWH)"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def mg_daily(df):
    df = df.copy()
    df["type"] = df["FACILITY TYPE"].str.strip().str.upper().map(GEN_TYPES).fillna("Other")
    gross = df.pivot_table(index="DATE", columns="type", values="GROSS INJECTION (MWH)", aggfunc="sum") / 1000
    gross.columns = [f"{c}_GWh" for c in gross.columns]
    tot = df.groupby("DATE")[["GROSS INJECTION (MWH)", "NET INJECTION (MWH)"]].sum() / 1000
    gross["Total_Gross_GWh"] = tot["GROSS INJECTION (MWH)"]
    gross["Total_Net_GWh"] = tot["NET INJECTION (MWH)"]
    gross["Periods"] = df.groupby("DATE")["PERIOD"].nunique()
    gross.index.name = "Date"
    return gross


def nems_year(year):
    r = get(f"{NEMS}/api/sitecore/DataSync/DataDownloadByYear", params={"value": "16", "year": str(year),
                                                                       "tpcValue": "1"})
    if r.status_code != 200 or r.content[:2] != b"PK":
        out(f"NEMS {year}: no zip (HTTP {r.status_code}, {len(r.content)} bytes)")
        return pd.DataFrame()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    parts = [parse_mg_csv(z.read(n).decode("utf-8-sig", errors="replace")) for n in z.namelist() if n.endswith(".csv")]
    df = pd.concat(parts) if parts else pd.DataFrame()
    out(f"NEMS {year}: {r.url.split('/')[-1]} -> {df['DATE'].min():%Y-%m-%d}..{df['DATE'].max():%Y-%m-%d}"
        if not df.empty else f"NEMS {year}: empty zip")
    return df


def nems_range(d0, d1):
    r = get(f"{NEMS}/api/sitecore/DataSync/DataDownload", params={
        "value": "16", "fromDate": d0.isoformat(), "toDate": d1.isoformat(), "tpcValue": "1"})
    if r.status_code != 200 or "csv" not in (r.headers.get("content-type") or ""):
        return pd.DataFrame()
    return parse_mg_csv(r.content.decode("utf-8-sig", errors="replace"))


def update_nems(gen):
    complete = set() if gen.empty else {d.date() for d in gen.index[gen["Periods"] >= 46]}
    yesterday = date.today() - timedelta(days=1)
    missing = [DATA_START + timedelta(days=i) for i in range((yesterday - DATA_START).days + 1)]
    missing = [d for d in missing if d not in complete]
    years = sorted({d.year for d in missing})
    out(f"NEMS: {len(missing)} missing days; years to fetch {years}")
    raw = [nems_year(y) for y in years]
    raw = [x for x in raw if not x.empty]
    got = pd.concat(raw) if raw else pd.DataFrame()
    last = got["DATE"].max().date() if not got.empty else (gen.index.max().date() if not gen.empty else DATA_START)
    d0 = last + timedelta(days=1)
    while d0 <= yesterday:   # days after the latest yearly zip
        d1 = min(d0 + timedelta(days=30), yesterday)
        part = nems_range(d0, d1)
        out(f"NEMS {d0}..{d1}: {0 if part.empty else part['DATE'].nunique()} days")
        if part.empty:
            break
        got = pd.concat([got, part])
        d0 = part["DATE"].max().date() + timedelta(days=1)
    if got.empty:
        return gen
    got = got.drop_duplicates(subset=["DATE", "PERIOD", "FACILITY TYPE"], keep="last")
    new = mg_daily(got)
    new = new[new.index >= pd.Timestamp(DATA_START)]
    gen = pd.concat([gen[~gen.index.isin(new.index)], new]).sort_index() if not gen.empty else new
    first = [c for c in ["CCGT_Cogen_Trigen_GWh", "Steam_turbine_GWh", "Gas_turbine_OCGT_GWh", "Solar_IGS_GWh",
                         "Imports_GWh", "Battery_ESS_GWh", "Other_GWh"] if c in gen]
    return gen[first + [c for c in gen.columns if c not in first]]


# ------------------------------------------------------------------ monthly + annual tables

def monthly_generation():
    r = get(SINGSTAT, params={"limit": 3000})
    d = r.json()["Data"]
    row = next(x for x in d["row"] if "generation" in x["rowText"].lower())
    s = pd.Series({pd.to_datetime(c["key"], format="%Y %b"): pd.to_numeric(c["value"], errors="coerce")
                   for c in row["columns"]}).sort_index()
    df = pd.DataFrame({"Electricity_Generation_GWh": s,
                       "Electricity_Generation_GWh_per_day": (s / s.index.days_in_month).round(2)})
    df.index.name = "Month"
    out(f"SingStat M890831: {df.index.min():%Y-%m}..{df.index.max():%Y-%m} (updated {d.get('dataLastUpdated')})")
    return df, d.get("dataLastUpdated")


def ses_workbook():
    url = SES_FALLBACK
    try:
        page = get(SES_PAGE).text
        m = re.search(r'href="([^"]*SES_tidy[^"]*\.xlsx[^"]*)"', page)
        if m:
            url = m.group(1) if m.group(1).startswith("http") else EMA + m.group(1)
    except requests.RequestException as e:
        out(f"SES page unavailable ({e}); using known link")
    r = get(url)
    r.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(r.content))
    toc = pd.read_excel(xl, sheet_name=0, header=None)
    out(f"SES: {url} ({len(r.content)} bytes) - {toc.iloc[0, 0]}")
    return xl, str(toc.iloc[0, 0]), url


def partial_note(df, col):
    """SES puts 'Data for 2025 is as at Jun 2025.' in the year column; return (clean df, note)."""
    note = next((str(v) for v in df[col] if not re.fullmatch(r"\d{4}(\.0)?", str(v)) and str(v) != "nan"), "")
    df = df[df[col].astype(str).str.fullmatch(r"\d{4}(\.0)?")].copy()
    df[col] = df[col].astype(float).astype(int)
    return df, note


def annual_consumption(xl):
    df, note = partial_note(pd.read_excel(xl, sheet_name="T3.2"), "eimh_year")
    df = df[(df["sector"] == df["sub_sector"]) | (df["sector"].str.lower() == df["sub_sector"].str.lower())]
    w = df.pivot_table(index="eimh_year", columns="sector", values="consumption_GWh", aggfunc="sum")
    ren = {"Industrial-related": "Industrial_GWh", "Commerce and Services-related": "Commerce_Services_GWh",
           "Households": "Households_GWh", "Transport-Related": "Transport_GWh", "Others": "Others_GWh",
           "Overall": "Total_GWh"}
    w = w.rename(columns=ren)[[c for c in ren.values() if c in w.rename(columns=ren)]]
    w.index = pd.to_datetime(w.index.astype(str), format="%Y")
    w.index.name = "Year"
    return w, note


def annual_fuel_mix(xl):
    df, note = partial_note(pd.read_excel(xl, sheet_name="T2.2"), "year")
    w = df.pivot_table(index="year", columns="energy_products", values="percentage", aggfunc="sum")
    w.columns = [f"{c.replace(' ', '_')}_pct" for c in w.columns]
    order = [c for c in ["Natural_Gas_pct", "Petroleum_Products_pct", "Coal_pct", "Solar_PV_pct", "Others_pct"]
             if c in w]
    w = w[order + [c for c in w.columns if c not in order]]
    w.index = pd.to_datetime(w.index.astype(str), format="%Y")
    w.index.name = "Year"
    return w, note


# ------------------------------------------------------------------ main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/singapore_power.xlsx")
    args = ap.parse_args()

    hh = read_sheet(args.out, HH_SHEET)
    gen = read_sheet(args.out, GEN_SHEET)
    out(f"archive: {len(hh)} half-hourly rows, {len(gen)} generation days")

    hh, ema_notes = update_ema(hh)
    daily = daily_demand(hh)
    gen = update_nems(gen)
    monthly, ss_updated = monthly_generation()
    xl, ses_title, ses_url = ses_workbook()
    cons, cons_note = annual_consumption(xl)
    mix, mix_note = annual_fuel_mix(xl)

    out(f"half-hourly demand: {hh.index.min()} .. {hh.index.max()} ({len(hh)} rows)")
    out(daily.tail(3).to_string())
    out(f"generation by type: {gen.index.min():%Y-%m-%d} .. {gen.index.max():%Y-%m-%d}")
    out(gen.tail(3).round(2).to_string())
    out(monthly.tail(3).to_string())
    out(cons.tail(3).to_string())
    out(mix.tail(3).to_string())

    notes = [
        "UNITS",
        f"{HH_SHEET}: MW per half-hour, indexed by period END time (Singapore time; 00:00 = end of the previous day). "
        "System_Demand_Actual_MW is EMA's system demand; NEM_Demand_Actual/Forecast_MW are the market (NEM) demand.",
        f"{DAILY_SHEET}: daily average / peak / minimum of the half-hourly System Demand (MW); System_Demand_GWh = sum of "
        "half-hourly MW x 0.5 h / 1000. Periods = half-hours present (48 = complete day).",
        f"{GEN_SHEET}: GWh per day of metered GROSS injection by facility type (sum of half-hourly MWh / 1000); "
        "Total_Net_GWh is net injection (after station load). Small negative values = auxiliary consumption of idle units.",
        f"{MONTHLY_SHEET}: total electricity generation, GWh per month (and per day = / days in month).",
        f"{CONS_SHEET}: electricity consumption by sector, GWh per year.",
        f"{MIX_SHEET}: share of fuel used for electricity generation, % per year.",
        "",
        "DEFINITIONS",
        *[f"EMA footnote: {n}" for n in ema_notes],
        "Facility types (NEMS): CCGT/COGEN/TRIGEN = combined-cycle gas turbines and cogeneration/trigeneration plants "
        "(almost all gas-fired); ST = steam turbines (incl. waste-to-energy); GT = open-cycle gas turbines; IGS = "
        "intermittent generation sources registered in the market (grid-scale solar only - most solar PV is embedded "
        "and not metered here); IMPORT = electricity imports; ESS/BATTERY = energy storage; OTHERS = other facilities.",
        "No monthly electricity consumption by sector is published - EMA/SES give it annually only.",
        "",
        "COVERAGE",
        f"Half-hourly and daily series from {DATA_START:%d %b %Y} (EMA publishes one file per week a few days after the "
        "week ends; NEMS metered generation is final about a week after the trading day and kept for a five-year "
        "rolling window). Monthly generation: full SingStat history from 1975 "
        f"(SingStat last updated {ss_updated}). Annual tables: SES ({ses_title}).",
        f"Part-year in the SES tables: {cons_note or mix_note or 'none'} - that row is NOT a full year.",
        "Incremental: half-hourly demand and daily generation archives are kept in this workbook and only missing days "
        "are fetched each run; monthly/annual tables are re-read (small).",
        "",
        "SOURCES",
        f"EMA Half-hourly System Demand Data: {EMA_PAGE}",
        f"EMC NEMS market data download (Metered Generation by Facility Type): {NEMS_PAGE}",
        "SingStat M890831 Electricity Generation, Monthly (source EMA): https://tablebuilder.singstat.gov.sg/table/TS/M890831",
        f"EMA Singapore Energy Statistics, tidy data workbook (tables T3.2, T2.2): {ses_url}",
    ]
    hh_out = hh.round(3)
    daily_out = daily.copy()
    gen_out = gen.round(3)
    for df in (daily_out, gen_out):
        df.index = df.index.strftime("%Y-%m-%d")
        df.index.name = "Date"
    mon_out = monthly.copy()
    mon_out.index = mon_out.index.strftime("%Y-%m")
    cons_out, mix_out = cons.copy(), mix.copy()
    for df, note in ((cons_out, cons_note), (mix_out, mix_note)):
        df.index = df.index.year
        df.index.name = "Year"
        m = re.search(r"(\d{4}) is (?:as )?(?:at|of) ([A-Za-z]{3})", note or "")
        df["Coverage"] = [f"Jan-{m.group(2)} only (part-year)" if m and y == int(m.group(1)) else "Full year"
                          for y in df.index]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {DAILY_SHEET: daily_out, GEN_SHEET: gen_out, MONTHLY_SHEET: mon_out,
                                         CONS_SHEET: cons_out, MIX_SHEET: mix_out, HH_SHEET: hh_out},
                              notes, {"UNITS", "DEFINITIONS", "COVERAGE", "SOURCES"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
