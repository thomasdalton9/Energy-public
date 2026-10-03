"""
Singapore natural gas: demand by sector, imports by pipeline vs LNG, town
gas sales, and a monthly estimate of gas burnt for power. Found via
discovery_archive/asia/SINGAPORE_DISCOVERY*.py.

What is published (free, no login):
  1. EMA Singapore Energy Statistics (SES) tidy workbook - ANNUAL only:
       T1.1 imports of energy products: Pipeline NG and Liquefied NG (ktoe), from 2005
       T2.1 energy inputs into main power producers + autoproducers: natural gas (ktoe),
            from 2005, latest year part-year (SES 2025 has Jan-Jun 2025)
       T3.7 final natural gas consumption by sector / sub-sector (TJ), from 2009
     https://www.ema.gov.sg/resources/singapore-energy-statistics
  2. SingStat M890371 'Piped Gas Sales, Quarterly' (town gas, million kWh = GWh,
     domestic / non-domestic), https://tablebuilder.singstat.gov.sg/table/TS/M890371
  3. ESTIMATE, daily and monthly: gas burnt for power = NEMS metered CCGT/COGEN/TRIGEN
     generation (daily, from the singapore_power.xlsx workbook written by
     asia/SINGAPORE_POWER.py) x a heat-rate factor calibrated each year as
     SES natural gas input to power generation (main power producers +
     autoproducers) / CCGT-cogen metered generation in the same period.
     Power takes ~85-90% of Singapore's gas supply, so this is the only
     monthly view of gas demand; it is labelled as an estimate throughout.

Not published: monthly or daily natural gas consumption, sendout or imports
(pipeline or LNG). EMA/SES give these annually only; SLNG and the pipeline
operators do not publish flows. Metered offtake data exists (PowerGas meters every
transmission offtake point and shares readings with shippers over its GTSS system),
but the Gas Network Code (Section K 4.2.3) makes Metering Data Confidential
Information; PowerGas publishes only maintenance and network-development plans and a
monthly shrinkage factor (discovery_archive/asia/SINGAPORE_GAS_DISCOVERY3.py). Town
gas (City Energy) is published quarterly only (SingStat M890371). The monthly Enterprise Singapore trade
dataset (SingStat T010002) has no public API.

    python3 asia/SINGAPORE_GAS.py --out "output/Data and Chart Outputs/singapore_gas.xlsx"
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (15, 180)
EMA = "https://www.ema.gov.sg"
SES_PAGE = f"{EMA}/resources/singapore-energy-statistics/chapter3"
SES_FALLBACK = (f"{EMA}/content/dam/corporate/resources/singapore-energy-statistics/excel/"
                "SES_tidy.xlsx.coredownload.xlsx")
TOWN_GAS = "https://tablebuilder.singstat.gov.sg/api/table/tabledata/M890371"
POWER_XLSX = "output/Data and Chart Outputs/singapore_power.xlsx"
TJ_PER_KTOE = 41.868
MCM_PER_KTOE = 1 / 0.9   # IEA rule of thumb: 1 bcm of natural gas ~ 0.9 Mtoe
TJ_PER_MCM = TJ_PER_KTOE / MCM_PER_KTOE


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


def ses_workbook():
    url = SES_FALLBACK
    try:
        m = re.search(r'href="([^"]*SES_tidy[^"]*\.xlsx[^"]*)"', get(SES_PAGE).text)
        if m:
            url = m.group(1) if m.group(1).startswith("http") else EMA + m.group(1)
    except requests.RequestException as e:
        out(f"SES page unavailable ({e}); using known link")
    r = get(url)
    r.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(r.content))
    title = str(pd.read_excel(xl, sheet_name=0, header=None).iloc[0, 0])
    out(f"SES: {url} ({len(r.content)} bytes) - {title}")
    return xl, title, url


def clean_years(df, col):
    note = next((str(v) for v in df[col] if not re.fullmatch(r"\d{4}(\.0)?", str(v)) and str(v) != "nan"), "")
    df = df[df[col].astype(str).str.fullmatch(r"\d{4}(\.0)?")].copy()
    df[col] = df[col].astype(float).astype(int)
    return df, note


def partial_year(note):
    """'Data for 2025 is as at Jun 2025.' -> (2025, 6); None if the note doesn't say."""
    m = re.search(r"(\d{4}) is (?:as )?(?:at|of) ([A-Za-z]{3})", note or "")
    return (int(m.group(1)), pd.to_datetime(m.group(2), format="%b").month) if m else None


def imports(xl):
    df, _ = clean_years(pd.read_excel(xl, sheet_name="T1.1"), "year")
    df = df[df["energy_products"].str.contains("Natural Gas", case=False)]
    w = df.pivot_table(index="year", columns="sub_products", values="value_ktoe", aggfunc="sum")
    w = w.rename(columns={"Pipeline NG": "Pipeline_ktoe", "Liquefied NG": "LNG_ktoe"})[["Pipeline_ktoe", "LNG_ktoe"]]
    w["Total_ktoe"] = w.sum(axis=1)
    for c in ("Pipeline", "LNG", "Total"):
        w[f"{c}_TJ"] = (w[f"{c}_ktoe"] * TJ_PER_KTOE).round(0)
    for c in ("Pipeline", "LNG", "Total"):
        days = pd.Series([366 if y % 4 == 0 else 365 for y in w.index], index=w.index)
        w[f"{c}_mcm_per_day_approx"] = (w[f"{c}_ktoe"] * MCM_PER_KTOE / days).round(2)
    w["LNG_share_pct"] = (100 * w["LNG_ktoe"] / w["Total_ktoe"]).round(1)
    w.index.name = "Year"
    return w


def power_gas(xl):
    df, note = clean_years(pd.read_excel(xl, sheet_name="T2.1"), "year")
    df = df[df["energy_products"].str.strip().str.lower() == "natural gas"]
    w = df.pivot_table(index="year", columns="energy_flow", values="value_ktoe", aggfunc="sum")
    mpp = next(c for c in w.columns if "main power" in c.lower())
    auto = next((c for c in w.columns if "autoproducer" in c.lower()), None)
    res = pd.DataFrame({"Main_power_producers_TJ": w[mpp] * TJ_PER_KTOE})
    res["Autoproducers_TJ"] = w[auto] * TJ_PER_KTOE if auto else 0.0
    res["Power_generation_TJ"] = res.sum(axis=1)
    res.index.name = "Year"
    return res.round(0), note


def final_consumption(xl):
    df, _ = clean_years(pd.read_excel(xl, sheet_name="T3.7"), "year")
    sec = df.pivot_table(index="year", columns="sector", values="ng_consumption_tj", aggfunc="sum")
    ren = {"Industrial-related": "Industrial_TJ", "Commerce and Service-related": "Commerce_Services_TJ",
           "Households": "Households_TJ", "Transport-related": "Transport_TJ", "Others": "Others_TJ"}
    sec = sec.rename(columns=ren)
    sec = sec[[c for c in ren.values() if c in sec]]
    sub = df.pivot_table(index="year", columns="sub_sector", values="ng_consumption_tj", aggfunc="sum")
    sub.columns = [re.sub(r"[^A-Za-z]+", "_", c).strip("_") + "_TJ" for c in sub.columns]
    sec.index.name = sub.index.name = "Year"
    return sec, sub


def town_gas():
    d = get(TOWN_GAS, params={"limit": 3000}).json()["Data"]
    cols = {}
    for row in d["row"]:
        name = {"total piped gas sales (town gas)": "Total_GWh", "domestic": "Domestic_GWh",
                "non-domestic": "Non_domestic_GWh"}.get(row["rowText"].strip().lower())
        if not name:
            continue
        s = {}
        for c in row["columns"]:
            y, q = c["key"].split()
            s[pd.Timestamp(int(y), 3 * int(q[0]) - 2, 1)] = pd.to_numeric(c["value"], errors="coerce")
        cols[name] = pd.Series(s)
    w = pd.DataFrame(cols).sort_index()
    w.index.name = "Quarter_start"
    out(f"SingStat M890371 town gas: {w.index.min():%Y-%m}..{w.index.max():%Y-%m} "
        f"[{d['row'][0].get('uoM')}] updated {d.get('dataLastUpdated')}")
    return w, d.get("dataLastUpdated"), d["row"][0].get("uoM")


def power_burn_estimate(power_xlsx, pg, part):
    """Monthly and daily gas burn for power, from metered CCGT/cogen generation x a yearly calibrated factor."""
    try:
        gen = pd.read_excel(power_xlsx, sheet_name="Daily generation by type", index_col=0)
    except (FileNotFoundError, ValueError) as e:
        out(f"power workbook not available ({e}); skipping the power-burn estimate")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    gen.index = pd.to_datetime(gen.index)
    ccgt = gen["CCGT_Cogen_Trigen_GWh"]
    # calibration periods: each full SES year, plus the part-year (Jan..month) if SES has one
    cal = []
    for y, row in pg.iterrows():
        if part and y == part[0]:
            months = part[1]
        else:
            months = 12
        g = ccgt[(ccgt.index.year == y) & (ccgt.index.month <= months)]
        full = g.index.nunique() >= 0.95 * sum(pd.Period(f"{y}-{m:02d}").days_in_month for m in range(1, months + 1))
        if full:
            cal.append({"Year": y, "Months": months, "Gas_for_power_TJ": row["Power_generation_TJ"],
                        "CCGT_Cogen_Trigen_GWh": g.sum(),
                        "Factor_TJ_per_GWh": row["Power_generation_TJ"] / g.sum()})
    cal = pd.DataFrame(cal).set_index("Year") if cal else pd.DataFrame()
    if cal.empty:
        out("no overlapping year between SES gas-for-power and metered generation; no estimate")
        return pd.DataFrame(), cal, pd.DataFrame()
    m = ccgt.resample("MS").agg(["sum", "count"])
    m = m[m["count"] >= 0.9 * m.index.days_in_month]   # complete months only
    fac = pd.Series([cal["Factor_TJ_per_GWh"].get(d.year, cal["Factor_TJ_per_GWh"].iloc[-1]) for d in m.index],
                    index=m.index)
    est = pd.DataFrame({"CCGT_Cogen_Trigen_GWh": m["sum"].round(1), "Factor_TJ_per_GWh": fac.round(3)})
    est["Gas_for_power_TJ_est"] = (est["CCGT_Cogen_Trigen_GWh"] * est["Factor_TJ_per_GWh"]).round(0)
    est["Gas_for_power_TJ_per_day_est"] = (est["Gas_for_power_TJ_est"] / est.index.days_in_month).round(1)
    est["Gas_for_power_mcm_per_day_est"] = (est["Gas_for_power_TJ_per_day_est"] / TJ_PER_MCM).round(2)
    est["Factor_basis"] = ["SES year" if d.year in cal.index else f"latest SES year ({cal.index[-1]})"
                           for d in est.index]
    est.index.name = "Month"
    # daily: the same factor on each day's metered CCGT/cogen generation
    dfac = pd.Series([cal["Factor_TJ_per_GWh"].get(d.year, cal["Factor_TJ_per_GWh"].iloc[-1]) for d in ccgt.index],
                     index=ccgt.index)
    day = pd.DataFrame({"CCGT_Cogen_Trigen_GWh": ccgt.round(2), "Factor_TJ_per_GWh": dfac.round(3)})
    day["Gas_for_power_TJ_est"] = (ccgt * dfac).round(1)
    day["Gas_for_power_mcm_per_day_est"] = (day["Gas_for_power_TJ_est"] / TJ_PER_MCM).round(2)
    day = day[day["CCGT_Cogen_Trigen_GWh"].notna()]
    day.index.name = "date"
    return est, cal.round(3), day


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/singapore_gas.xlsx")
    ap.add_argument("--power-xlsx", default=POWER_XLSX)
    args = ap.parse_args()

    xl, ses_title, ses_url = ses_workbook()
    imp = imports(xl)
    pg, pg_note = power_gas(xl)
    part = partial_year(pg_note)
    sec, sub = final_consumption(xl)
    tg, tg_updated, tg_unit = town_gas()
    est, cal, day = power_burn_estimate(args.power_xlsx, pg, part)

    demand = pd.concat([pg[["Power_generation_TJ"]], sec], axis=1)
    demand["Final_consumption_TJ"] = sec.sum(axis=1, min_count=1)
    demand["Total_TJ"] = demand["Power_generation_TJ"] + demand["Final_consumption_TJ"]
    demand["Power_share_pct"] = (100 * demand["Power_generation_TJ"] / demand["Total_TJ"]).round(1)
    demand["Coverage"] = [f"Jan-{pd.Timestamp(2000, part[1], 1):%b} only (part-year)" if part and y == part[0]
                          else "Full year" for y in demand.index]
    demand = demand[demand.index >= 2009]
    demand.index.name = "Year"

    out(imp.tail(3).to_string())
    out(demand.tail(4).to_string())
    out(tg.tail(4).to_string())
    if not est.empty:
        out(cal.to_string())
        out(est.tail(4).to_string())

    notes = [
        "UNITS",
        "TJ = terajoules (EMA's unit for gas). ktoe = thousand tonnes of oil equivalent (1 ktoe = 41.868 TJ). "
        "*_mcm_per_day_approx / *_est in mcm: million cubic metres per day, APPROXIMATE, converted with the IEA rule "
        f"of thumb 1 bcm ~ 0.9 Mtoe (1 mcm ~ {TJ_PER_MCM:.1f} TJ).",
        "Annual demand by sector: TJ per year. Power_generation_TJ = natural gas input to main power producers + "
        "autoproducers (SES T2.1); the other sectors are final end-use consumption (SES T3.7).",
        f"Town gas quarterly: piped gas (town gas) sales, {tg_unit} per quarter (million kWh = GWh). Town gas in "
        "Singapore is made from natural gas.",
        "Power burn monthly / daily (estimate): see METHOD. Daily = the same calculation on each day's metered "
        "generation (TJ per day and mcm per day).",
        "",
        "METHOD (power-burn estimate)",
        "Gas_for_power_TJ_est = monthly NEMS metered gross generation of CCGT/COGEN/TRIGEN plants (GWh, from "
        "singapore_power.xlsx) x Factor_TJ_per_GWh. The factor is calibrated for each year as SES natural gas input "
        "to power generation (main power producers + autoproducers, TJ) / metered CCGT/COGEN/TRIGEN generation (GWh) "
        "over the same months (sheet 'Calibration'); months after the latest SES year use the latest factor. It "
        "folds in autoproducer and embedded gas use, so it is a scaled proxy, not a measured flow.",
        "",
        "COVERAGE",
        f"Annual SES tables from 2005 (imports, gas for power) / 2009 (final consumption by sector). {ses_title}. "
        f"Part-year: {pg_note or 'none'}",
        f"Town gas: quarterly from 1994 (SingStat last updated {tg_updated}). Power-burn estimate: daily and monthly "
        "from 2021 (monthly: complete months only), to the latest NEMS metered day (final about a week after the day).",
        "NOT PUBLISHED: monthly/daily natural gas consumption, sendout or imports (pipeline or LNG) - EMA/SES publish "
        "these annually only and SLNG / pipeline operators do not publish flows. PowerGas meters every offtake point, but "
        "the Gas Network Code (Section K 4.2.3) makes metering data confidential (shippers only); town gas is "
        "published quarterly only.",
        "Each run re-reads the (small) source tables; the power-burn estimate is rebuilt from the power workbook.",
        "",
        "SOURCES",
        f"EMA Singapore Energy Statistics tidy workbook (T1.1, T2.1, T3.7): {ses_url}",
        "SingStat M890371 Piped Gas Sales, Quarterly: https://tablebuilder.singstat.gov.sg/table/TS/M890371",
        "EMC NEMS metered generation by facility type, via asia/SINGAPORE_POWER.py: https://www.nems.emcsg.com/nems-prices",
    ]
    sheets = {"Annual demand by sector": demand, "Annual imports": imp, "Final consumption by subsector": sub}
    tg_out = tg.copy()
    tg_out.index = tg_out.index.strftime("%Y-%m")
    if not est.empty:
        est_out = est.copy()
        est_out.index = est_out.index.strftime("%Y-%m")
        sheets["Power burn monthly (est)"] = est_out
        day_out = day.copy()
        day_out.index = day_out.index.strftime("%Y-%m-%d")
        sheets["Power burn daily (est)"] = day_out
        sheets["Calibration"] = cal
    sheets["Town gas quarterly"] = tg_out
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "METHOD (power-burn estimate)", "COVERAGE", "SOURCES"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
