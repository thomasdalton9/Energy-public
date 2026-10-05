"""
MISO natural gas burn for power (Bcf/d), estimated from MISO's daily gas generation and a heat rate calibrated on
EIA-923 fuel use.

MISO does not publish gas consumption. This script converts the daily average gas MW in MISO's own real-time
fuel mix report (miso_fuel_mix_daily.xlsx, MISO_FUEL_MIX_DAILY.py) into gas burn:

    burn (MMBtu) = gas MWh x heat rate (MMBtu/MWh);   Bcf = MMBtu / (MMBtu per Mcf) / 1e6

Heat rate calibration (EIA-923, Page 1 'Generation and Fuel Data', monthly, per plant and prime mover):
  rows with Balancing Authority Code = MISO and fuel NG (natural gas; every prime mover: combined cycle, gas
  turbine, steam, engines). Fuel burned for electricity ('Elec_MMBtu', which splits a CHP plant's gas between
  power and useful heat) and net generation ('Netgen', MWh) are summed per month.
    physical heat rate  = Elec_MMBtu / Netgen              (MMBtu per net MWh, the plants' own efficiency)
    effective heat rate = Elec_MMBtu / MISO fuel-mix MWh   (the same burn over the MWh MISO's report calls gas)
  The daily estimate uses the EFFECTIVE heat rate, so each calibrated month's burn equals EIA-923's burn for
  the MISO balancing authority and MISO's day-to-day shape spreads it over the days. MISO's report counts
  generation at a different point (and its gas category has a slightly different footprint) from EIA's net
  generation; the effective rate absorbs that, the physical rate is shown for comparison.
  Heat content (MMBtu per Mcf) is likewise EIA-923's own: total gas MMBtu / total gas Mcf for the same plants.

EIA-923 lags by about two months. For months not published yet the heat rate (and heat content) of the SAME
MONTH A YEAR EARLIER is carried (summer peakers give a clear seasonal pattern), flagged 'estimated'. A month
EIA-923 later publishes is calibrated automatically on the next run.

Incremental: the committed workbook's 'Monthly' sheet is the store for the EIA-923 monthly table. A year's file
(f923_<year>.zip, ~15 MB) is downloaded only when it is not saved yet or its Last-Modified changed (recorded on
the Units sheet); the daily sheet is rebuilt each run from the fuel-mix workbook (cheap, deterministic).
"""

import argparse
import io
import os
import re
import sys
import zipfile
from datetime import date

import numpy as np
import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (15, 180)
BASE = "https://www.eia.gov/electricity/data/eia923/"
FIRST_YEAR = 2021
BA = "MISO"
FUEL = "NG"
FLAT_HR = 7.5              # the earlier flat assumption, kept for comparison
FLAT_MMBTU_PER_MCF = 1.036
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
DEFAULT_FUELMIX = "output/Data and Chart Outputs/miso_fuel_mix_daily.xlsx"
DEFAULT_OUT = "output/Data and Chart Outputs/miso_gas_burn_daily.xlsx"
FILE_TAG = "EIA923 file"


def urls_for(year):
    return [f"{BASE}xls/f923_{year}.zip", f"{BASE}archive/xls/f923_{year}.zip"]


def find_zip(year, session, want_body):
    """(url, last_modified, content or None) of a year's EIA-923 zip: the current-year location first, then the
    archive folder (EIA moves a year there once final)."""
    for u in urls_for(year):
        try:
            if want_body:
                r = session.get(u, headers=HEADERS, timeout=TIMEOUT)
            else:
                r = session.head(u, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        except requests.RequestException as e:
            print(f"  {u}: {type(e).__name__}")
            continue
        ctype = r.headers.get("Content-Type", "")
        ok = r.status_code == 200 and ("zip" in ctype or "octet" in ctype or (want_body and r.content[:2] == b"PK"))
        if ok:
            return u, r.headers.get("Last-Modified", ""), (r.content if want_body else None)
    return None, "", None


def parse_year(content, year):
    """Monthly MISO-BA natural gas table from one EIA-923 zip: rows (month) with Netgen_MWh, Elec_MMBtu, Tot_MMBtu,
    Mcf, Plants. Only the months the file has published (named in its file name) are returned."""
    zf = zipfile.ZipFile(io.BytesIO(content))
    name = next(n for n in zf.namelist() if re.search(r"Schedules_2_3_4_5", n) and n.lower().endswith(".xlsx"))
    m = re.search(r"_M_(\d{2})_(\d{4})", name)
    last_month = int(m.group(1)) if m else 12
    xl = pd.ExcelFile(zf.open(name))
    sheet = next(s for s in xl.sheet_names if s.lower().startswith("page 1 generation"))
    raw = pd.read_excel(xl, sheet_name=sheet, header=None)
    hdr = int(raw.index[raw.iloc[:, 0].astype(str).str.strip().str.lower() == "plant id"][0])
    d = raw.iloc[hdr + 1:].copy()
    d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[hdr]]
    d = d.loc[:, ~d.columns.duplicated()]
    ba_col = next(c for c in d.columns if c.lower().startswith("balancing authority code"))
    fuel_col = next(c for c in d.columns if c.lower().startswith("reported fuel type code"))
    d = d[(d[ba_col].astype(str).str.strip() == BA) & (d[fuel_col].astype(str).str.strip() == FUEL)]
    rows = []
    for i, mon in enumerate(MONTHS, start=1):
        if i > last_month:
            break
        num = lambda p: pd.to_numeric(d[f"{p} {mon}"], errors="coerce")  # noqa: E731
        ng, em, tm, q = num("Netgen"), num("Elec_MMBtu"), num("Tot_MMBtu"), num("Quantity")
        rows.append({"month": pd.Timestamp(year, i, 1), "Netgen_MWh": ng.sum(), "Elec_MMBtu": em.sum(),
                     "Tot_MMBtu": tm.sum(), "Mcf": q.sum(), "Plants": int(((em > 0) | (ng > 0)).sum())})
    out = pd.DataFrame(rows).set_index("month")
    return out[out["Elec_MMBtu"] > 0], name


def load_saved(path):
    """(saved EIA-923 monthly table, {year: 'url | Last-Modified'}) from the committed workbook."""
    if not os.path.exists(path):
        return pd.DataFrame(), {}
    try:
        mo = pd.read_excel(path, sheet_name="Monthly")
        units = pd.read_excel(path, sheet_name="Units")["Notes"].fillna("").astype(str)
    except Exception as e:  # noqa: BLE001
        print(f"saved workbook unreadable ({type(e).__name__}); rebuilding")
        return pd.DataFrame(), {}
    meta = {}
    for line in units:
        m = re.match(rf"{FILE_TAG} (\d{{4}}) \| (.*)$", line)
        if m:
            meta[int(m.group(1))] = m.group(2)
    mo["month"] = pd.to_datetime(mo["month"], errors="coerce")
    mo = mo.dropna(subset=["month"]).set_index("month")
    keep = [c for c in ("Netgen_MWh", "Elec_MMBtu", "Tot_MMBtu", "Mcf", "Plants") if c in mo.columns]
    return mo[keep].dropna(subset=["Elec_MMBtu"]), meta


def update_eia923(saved, meta):
    """Download only the years not saved yet, or whose Last-Modified changed. Returns (table, meta)."""
    table = saved.copy()
    s = requests.Session()
    this_year = date.today().year
    for year in range(FIRST_YEAR, this_year + 1):
        have = (not table.empty) and (table.index.year == year).any()
        url, lm, _ = find_zip(year, s, want_body=False)
        if url is None:
            print(f"{year}: no EIA-923 file published")
            continue
        tag = f"{url} | Last-Modified: {lm}"
        if have and meta.get(year) == tag:
            print(f"{year}: unchanged ({lm}), kept")
            continue
        url, lm, content = find_zip(year, s, want_body=True)
        if content is None:
            print(f"{year}: download failed; keeping saved months")
            continue
        t, name = parse_year(content, year)
        print(f"{year}: {name}: {len(t)} MISO gas months, {t['Netgen_MWh'].sum() / 1e6:.1f} TWh")
        table = (t if table.empty else pd.concat([table[table.index.year != year], t])).sort_index()
        meta[year] = f"{url} | Last-Modified: {lm}"
    return table, meta


def monthly_table(eia, daily):
    """Monthly sheet: EIA-923 burn and net generation, MISO fuel-mix GWh, physical and effective heat rate, MMBtu/Mcf,
    and the heat rate/heat content used (calibrated, or the same month a year earlier)."""
    m = pd.DataFrame(index=pd.date_range(eia.index.min() if len(eia) else daily.index.min().replace(day=1),
                                         daily.index.max().replace(day=1), freq="MS"))
    m = m.join(eia)
    g = daily["Gas_GWh"].resample("MS")
    m["MISO_fuelmix_GWh"] = g.sum(min_count=1)
    m["MISO_days"] = daily["Gas_GWh"].resample("MS").count()
    m["Days_in_month"] = m.index.days_in_month
    # a month with up to 10% of its days missing from the fuel-mix archive is scaled up to the full month
    full = m["MISO_days"] >= 0.9 * m["Days_in_month"]
    m["MISO_fuelmix_GWh"] = (m["MISO_fuelmix_GWh"] / m["MISO_days"] * m["Days_in_month"]).where(full)
    cal = m["Elec_MMBtu"].notna() & m["MISO_fuelmix_GWh"].notna()
    m["EIA923_GWh"] = m["Netgen_MWh"] / 1e3
    m["Physical_HR_MMBtu_per_MWh"] = m["Elec_MMBtu"] / m["Netgen_MWh"]
    m["Effective_HR_MMBtu_per_MWh"] = m["Elec_MMBtu"] / (m["MISO_fuelmix_GWh"] * 1e3)
    m["MISO_vs_EIA923_GWh_ratio"] = m["MISO_fuelmix_GWh"] / m["EIA923_GWh"]
    m["MMBtu_per_Mcf"] = m["Tot_MMBtu"] / m["Mcf"]
    m["Elec_share_of_gas_MMBtu"] = m["Elec_MMBtu"] / m["Tot_MMBtu"]

    # EIA-923's current-year monthly file holds only the plants that report monthly (about 210 of the ~470 MISO gas
    # plants; the small ones arrive with the annual file), so its burn and net generation are a partial sample
    # and burn/MISO MWh would be biased low. A month is a full census when its plant count is near the
    # previous 12 months' maximum.
    plants = m["Plants"]
    census = pd.Series(False, index=m.index)
    for i, ts in enumerate(m.index):
        if pd.isna(plants.iloc[i]):
            continue
        prev = plants.iloc[max(0, i - 12):i].dropna()
        census.iloc[i] = len(prev) == 0 or plants.iloc[i] >= 0.75 * prev.max()
    m["Full_plant_census"] = census
    phys = m["Physical_HR_MMBtu_per_MWh"]
    eff = m["Effective_HR_MMBtu_per_MWh"].where(cal & census)
    factor = (eff / phys).dropna()      # MISO-MWh basis vs EIA net-MWh basis, by month, from census months
    series_hc = m["MMBtu_per_Mcf"].where(m["Mcf"] > 0)
    hr, hc, flag = {}, {}, {}
    for ts in m.index:
        prev = ts - pd.DateOffset(years=1)
        if pd.notna(eff.get(ts)):
            hr[ts], flag[ts] = eff[ts], "calibrated (EIA-923 burn / MISO MWh)"
        elif pd.notna(phys.get(ts)) and phys.get(ts) > 0:
            f = factor.get(prev, factor.tail(12).mean() if len(factor) else 1.0)
            hr[ts] = phys[ts] * f
            flag[ts] = ("calibrated, partial plant sample: EIA-923 net-MWh heat rate x %s MISO/EIA factor" % f"{prev:%b/%y}"
                        if prev in factor.index else "calibrated, partial plant sample: EIA-923 net-MWh heat rate x trailing-12 factor")
        elif prev in hr:
            hr[ts], flag[ts] = hr[prev], f"estimated: {prev:%b/%y} rate carried"
        else:
            known = [v for k, v in hr.items() if "estimated" not in flag[k]][-12:]
            hr[ts], flag[ts] = (np.mean(known) if known else np.nan), "estimated: trailing-12-month mean"
        hc[ts] = series_hc.get(ts) if pd.notna(series_hc.get(ts)) else hc.get(prev, np.nan)
    hr, hc, flag = ([d[t] for t in m.index] for d in (hr, hc, flag))
    m["Heat_rate_used_MMBtu_per_MWh"] = hr
    m["MMBtu_per_Mcf_used"] = pd.Series(hc, index=m.index).fillna(FLAT_MMBTU_PER_MCF)
    m["Heat_rate_basis"] = flag
    return m


def daily_table(fuelmix_path):
    d = pd.read_excel(fuelmix_path, sheet_name="Data")
    d["date"] = pd.to_datetime(d["date"].astype(str), errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    out = pd.DataFrame({"Gas_MW_avg": pd.to_numeric(d["Natural Gas_MW"], errors="coerce")})
    out = out[out["Gas_MW_avg"].notna()]
    out["Gas_GWh"] = out["Gas_MW_avg"] * 24 / 1e3
    return out


def build_daily(daily, monthly):
    out = daily.copy()
    mk = out.index.to_period("M").to_timestamp()
    out["Heat_rate_MMBtu_per_MWh"] = mk.map(monthly["Heat_rate_used_MMBtu_per_MWh"]).values
    out["Heat_rate_basis"] = mk.map(monthly["Heat_rate_basis"]).values
    out["MMBtu_per_Mcf"] = mk.map(monthly["MMBtu_per_Mcf_used"]).values
    out["Gas_burn_MMBtu_d"] = out["Gas_GWh"] * 1e3 * out["Heat_rate_MMBtu_per_MWh"]
    out["Gas_burn_MMcf_d"] = out["Gas_burn_MMBtu_d"] / out["MMBtu_per_Mcf"] / 1e3
    out["Gas_burn_Bcf_per_day"] = out["Gas_burn_MMcf_d"] / 1e3
    out["Gas_burn_Bcf_per_day_7d_avg"] = out["Gas_burn_Bcf_per_day"].rolling(7, min_periods=7).mean()
    out["Flat_7.5_Bcf_per_day"] = out["Gas_GWh"] * 1e3 * FLAT_HR / FLAT_MMBTU_PER_MCF / 1e6
    out.index.name = "date"
    return out.round({"Gas_MW_avg": 1, "Gas_GWh": 3, "Heat_rate_MMBtu_per_MWh": 3, "MMBtu_per_Mcf": 4,
                      "Gas_burn_MMBtu_d": 0, "Gas_burn_MMcf_d": 1, "Gas_burn_Bcf_per_day": 4,
                      "Gas_burn_Bcf_per_day_7d_avg": 4, "Flat_7.5_Bcf_per_day": 4})


def monthly_output(monthly, daily):
    """Monthly sheet: EIA-923 table + monthly average Bcf/d."""
    m = monthly.copy()
    bcfd = daily["Gas_burn_Bcf_per_day"].resample("MS").mean().reindex(m.index)
    full = daily["Gas_burn_Bcf_per_day"].resample("MS").count().reindex(m.index) >= 0.9 * m["Days_in_month"]
    m["Gas_burn_Bcf_per_day"] = bcfd.where(full)
    m["Gas_burn_Bcf_per_day_prior_year"] = m["Gas_burn_Bcf_per_day"].shift(12)
    m["Flat_7.5_Bcf_per_day"] = daily["Flat_7.5_Bcf_per_day"].resample("MS").mean().reindex(m.index).where(full)
    m["Gas_burn_Bcf_month"] = m["Gas_burn_Bcf_per_day"] * m["Days_in_month"]
    cols = ["Heat_rate_used_MMBtu_per_MWh", "Heat_rate_basis", "MMBtu_per_Mcf_used", "Gas_burn_Bcf_per_day",
            "Gas_burn_Bcf_per_day_prior_year", "Flat_7.5_Bcf_per_day", "Gas_burn_Bcf_month",
            "MISO_fuelmix_GWh", "EIA923_GWh", "Full_plant_census", "MISO_vs_EIA923_GWh_ratio", "Physical_HR_MMBtu_per_MWh",
            "Effective_HR_MMBtu_per_MWh", "MMBtu_per_Mcf", "Elec_share_of_gas_MMBtu", "Plants", "Elec_MMBtu",
            "Tot_MMBtu", "Mcf", "Netgen_MWh", "MISO_days"]
    m = m[cols].copy()
    m.index.name = "month"
    return m.round(4)


def notes(meta, last_cal, last_day):
    lines = [
        "UNITS",
        "Daily sheet: Gas_MW_avg = MISO's daily average gas MW (miso_fuel_mix_daily.xlsx, 'Natural Gas_MW'); Gas_GWh = MW x 24 / 1000.",
        "Heat_rate_MMBtu_per_MWh = MMBtu of gas burned per MWh MISO reports as gas, monthly (see Monthly sheet; same value on every day of the month).",
        "Gas_burn_MMBtu_d = Gas_GWh x 1000 x heat rate. Gas_burn_MMcf_d = MMBtu / (MMBtu per Mcf) / 1000. Gas_burn_Bcf_per_day = MMcf_d / 1000.",
        "1 Mcf = about 1.036 MMBtu on average; the script uses EIA-923's own monthly MMBtu per Mcf for the MISO gas plants (Monthly sheet, MMBtu_per_Mcf_used).",
        "Gas_burn_Bcf_per_day_7d_avg = trailing 7-day mean. Flat_7.5_Bcf_per_day = the earlier quick estimate (7.5 MMBtu/MWh, 1.036 MMBtu/Mcf), for comparison only.",
        "",
        "METHOD",
        "MISO does not publish gas burn. Burn = MISO gas MWh x heat rate. The heat rate is calibrated on EIA-923 Page 1 (Generation and Fuel Data): all natural-gas (fuel NG) rows",
        "whose balancing authority code is MISO, every prime mover (combined cycle, gas turbine, steam, engine). Fuel for electricity (Elec_MMBtu, which excludes a CHP plant's",
        "useful heat) over MISO fuel-mix MWh = effective heat rate (used). Over EIA net generation = physical heat rate (the plants' efficiency; shown for comparison).",
        "The current year's EIA-923 monthly file covers only the plants that report monthly (about half of MISO's gas plants; the rest arrive with the annual file), so its burn is",
        "a partial sample. For those months the sample's net-MWh heat rate (representative, within about 0.2 of last year's) is converted to MISO's MWh basis with the same month's",
        "MISO/EIA factor a year earlier (Heat_rate_basis says 'partial plant sample'); they are recalibrated on the full census when EIA publishes the annual file.",
        "Months EIA-923 has not published (about 2 months lag) carry the heat rate and heat content of the same month a year earlier (Heat_rate_basis says 'estimated').",
        "Seasonality (summer peakers lift the rate) is therefore kept, a changed fleet mix is not.",
        "",
        "MISO 'NATURAL GAS' CATEGORY",
        "Gas-fired generation of units in MISO's market footprint (central US plus MISO South: Louisiana, Mississippi, Arkansas, east Texas) as MISO's real-time fuel mix report",
        "classifies it. EIA's MISO balancing authority code is the nearest EIA footprint but not identical to the market footprint; the ratio of MISO MWh to EIA-923 net MWh",
        "is shown monthly (MISO_vs_EIA923_GWh_ratio).",
        "",
        "COVERAGE",
        f"Daily: 2023-01-01 to {last_day:%Y-%m-%d} (limited by the MISO fuel mix archive). EIA-923 calibrated months: through {last_cal:%b %Y}.",
        "",
        "SOURCE",
        "MISO daily real-time generation fuel mix report (docs.misoenergy.org/marketreports/<YYYYMMDD>_sr_gfm.xlsx); EIA-923 Monthly Generation and Fuel Consumption Time Series",
        "File (https://www.eia.gov/electricity/data/eia923/), Page 1 Generation and Fuel Data. Estimate, not a measured gas burn.",
        "",
        "EIA-923 FILES (incremental: a year is re-downloaded only when its Last-Modified changes)",
    ]
    lines += [f"{FILE_TAG} {y} | {v}" for y, v in sorted(meta.items())]
    titles = {"UNITS", "METHOD", "MISO 'NATURAL GAS' CATEGORY", "COVERAGE", "SOURCE",
              "EIA-923 FILES (incremental: a year is re-downloaded only when its Last-Modified changes)"}
    return lines, titles


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--fuelmix", default=DEFAULT_FUELMIX)
    ap.add_argument("--offline", action="store_true", help="skip EIA downloads (use the saved monthly table)")
    args = ap.parse_args()

    saved, meta = load_saved(args.out)
    if args.offline:
        eia = saved
    else:
        eia, meta = update_eia923(saved, meta)
    if eia.empty:
        sys.exit("no EIA-923 data for the MISO balancing authority")
    daily0 = daily_table(args.fuelmix)
    mon = monthly_table(eia, daily0)
    daily = build_daily(daily0, mon)
    mout = monthly_output(mon, daily)
    last_cal = mon.index[mon["Heat_rate_basis"].str.startswith("calibrated")].max()
    lines, titles = notes(meta, last_cal, daily.index.max())

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Monthly": mout}, lines, titles)
    print(f"Saved {args.out}: {len(daily)} days, EIA-923 calibrated through {last_cal:%b %Y}")
    pd.set_option("display.width", 250)
    print(mout[["Heat_rate_used_MMBtu_per_MWh", "Heat_rate_basis", "Physical_HR_MMBtu_per_MWh", "MMBtu_per_Mcf_used",
                "MISO_vs_EIA923_GWh_ratio", "Gas_burn_Bcf_per_day", "Flat_7.5_Bcf_per_day"]].tail(40).to_string())


if __name__ == "__main__":
    main()
