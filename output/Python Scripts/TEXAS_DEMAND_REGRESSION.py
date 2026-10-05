"""
Texas gas and power demand REGRESSIONS (monthly), weather-normalised forecast to Dec 2033. A standalone workbook:
it reads texas_gas_monthly.xlsx, eia930_fuel_mix_daily.xlsx, henry_hub_daily.xlsx and (fallback only) macro_drivers.xlsx and
changes none of them or of the Texas forecast scripts; the owner may wire it into the existing forecast later.
Runs in GitHub Actions (.github/workflows/texas_demand_regression.yml, 1st and 15th, after texas_gas); NASA POWER, FRED and the
EIA-930 data are not reachable from the Claude sandbox.

Models (ordinary least squares, numpy; no statsmodels needed)
  1. Texas gas consumption by sector, Bcf/d, monthly from 2015: electric power, residential, commercial, industrial (published
     EIA sectors) on weather (HDD/CDD per day, base 18C, population-weighted over Houston, Dallas-Fort Worth, San Antonio, Austin and
     El Paso; NASA POWER daily T2M, Open-Meteo fallback) OR month dummies, Texas population, a linear trend and (electric power
     only) Henry Hub and the ERCOT wind+solar share. Every combination is a candidate; the chosen one has the lowest error on a
     hold-out of the last 12 months (fit on the months before) among candidates whose signs are plausible (HDD>0 for
     residential/commercial, CDD>0 for electric power, population>=0). All candidates are listed on the 'Candidates' tab.
  2. ERCOT load (EIA-930 ERCOT total net generation as the load proxy, average GW) on HDD, CDD and population. The BASELINE is fit on a
     pre-step-up window (editable, default Jan 2019 - Dec 2022) and extrapolated; actual minus baseline is the 'unexplained load
     growth' (the data-centre signal). A full-sample model with a trend is reported for reference with a hold-out.
  3. Forecast to Dec 2033 at normal weather (average HDD/CDD by calendar month over an editable window of years), population
     growing at an editable rate, Henry Hub and wind+solar share held at editable values, +/-1 prediction standard error band.
  4. LOAD BREAKOUT ('Load breakout' + 'Load sources' tabs + chart): the unexplained load set against the one OBSERVED figure ERCOT publishes
     (large loads >=75 MW, monthly peak, all types together) and labelled SOURCED / JUDGEMENT / NOT ATTRIBUTED. No guessed shares; per-category
     figures (data centres, crypto, oil & gas, industrial) exist only as ERCOT FORECASTS or QUEUE requests, listed on 'Load sources' (with Oncor / CenterPoint /
     AEP Texas disclosures and named crypto / data-centre sites), never mixed with the observed series; further charts show the ERCOT queue by type and the observed split by zone / connection type (Feb, Mar 2026).
Edit the yellow cells of the Assumptions tab and the monthly ERCOT large-load column of the Load breakout tab (committed workbook is read back on every run) and re-run.

Usage: python3 TEXAS_DEMAND_REGRESSION.py [--out "output/Data and Chart Outputs/texas_demand_regression.xlsx"]
"""
import argparse
import datetime as dt
import io
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from itertools import product

import numpy as np
import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(OUT_DIR, "texas_demand_regression.xlsx")
GAS_XLSX = os.path.join(OUT_DIR, "texas_gas_monthly.xlsx")
E930_XLSX = os.path.join(OUT_DIR, "eia930_fuel_mix_daily.xlsx")
HH_XLSX = os.path.join(OUT_DIR, "henry_hub_daily.xlsx")
MACRO_XLSX = os.path.join(OUT_DIR, "macro_drivers.xlsx")

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
NASA = "https://power.larc.nasa.gov/api/temporal/daily/point"
OPEN_METEO = "https://archive-api.open-meteo.com/v1/archive"
FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=TXPOP"
START = pd.Timestamp("2015-01-01")
END = pd.Timestamp("2033-12-01")
BASE_C = 18.0
# metro populations in millions, approx. 2020 Census (weights only; from memory, unverified)
CITIES = [("Houston", 29.76, -95.37, 7.12), ("Dallas-Fort Worth", 32.78, -96.80, 7.64), ("San Antonio", 29.42, -98.49, 2.56),
          ("Austin", 30.27, -97.74, 2.28), ("El Paso", 31.76, -106.49, 0.87)]
SECTORS = ["Electric power", "Residential", "Commercial", "Industrial"]
HOLD = 12
TEXAS_SHARE_OF_US_2020 = 0.0879   # fallback only: Texas 29.1m / US 331m


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------ data pulls
def get_json(url, params, tries=3, timeout=300):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, params=params, headers=UA, timeout=timeout)
            r.raise_for_status()
            return r.json()
        except Exception as exc:  # noqa: BLE001
            last = exc
            time.sleep(5 * (i + 1))
    raise last


def fetch_city(job):
    name, lat, lon, start, end = job
    try:
        js = get_json(NASA, {"parameters": "T2M", "community": "RE", "longitude": lon, "latitude": lat,
                             "start": start.strftime("%Y%m%d"), "end": end.strftime("%Y%m%d"), "format": "JSON"})
        t = js["properties"]["parameter"]["T2M"]
        s = pd.Series({pd.Timestamp(k): (np.nan if v is None or v <= -998 else float(v)) for k, v in t.items()}).dropna()
        if s.empty:
            raise RuntimeError("no data")
        return name, s, "NASA POWER"
    except Exception as exc:  # noqa: BLE001
        log(f"  NASA POWER failed for {name}: {exc}; trying Open-Meteo")
    try:
        time.sleep(8)
        js = get_json(OPEN_METEO, {"latitude": lat, "longitude": lon, "start_date": start.date().isoformat(),
                                   "end_date": end.date().isoformat(), "daily": "temperature_2m_mean", "timezone": "UTC"})
        d = js["daily"]
        s = pd.Series(d["temperature_2m_mean"], index=pd.to_datetime(d["time"]), dtype=float).dropna()
        return name, s, "Open-Meteo ERA5"
    except Exception as exc:  # noqa: BLE001
        log(f"  Open-Meteo failed for {name}: {exc}")
        return name, pd.Series(dtype=float), "none"


def update_temps(saved):
    """Incremental city store: fetch only days after the last saved one (minus a 10-day revision window)."""
    end = pd.Timestamp(dt.date.today() - dt.timedelta(days=1))
    jobs = []
    for name, lat, lon, _ in CITIES:
        if name in saved and saved[name].notna().any():
            start = max(START, saved[name].dropna().index.max() - pd.Timedelta(days=10))
        else:
            start = START
        if start <= end:
            jobs.append((name, lat, lon, start, end))
    log(f"weather: {len(jobs)} city requests")
    new, srcs = {}, set()
    with ThreadPoolExecutor(max_workers=5) as ex:
        for name, s, src in ex.map(fetch_city, jobs):
            if not s.empty:
                new[name] = s.round(2)
                srcs.add(src)
    out = pd.DataFrame(new).combine_first(saved) if new else saved.copy()
    out = out.sort_index()
    out.index.name = "date"
    return out[[c[0] for c in CITIES if c[0] in out]], (", ".join(sorted(srcs)) or "saved store only")


def degree_days(temps):
    """Population-weighted daily HDD/CDD (a day counts only when every city has data), summed to months."""
    w = pd.Series({c[0]: c[3] for c in CITIES})
    t = temps[list(w.index)].dropna()
    hdd = ((BASE_C - t).clip(lower=0) * w).sum(axis=1) / w.sum()
    cdd = ((t - BASE_C).clip(lower=0) * w).sum(axis=1) / w.sum()
    m = pd.DataFrame({"HDD_total": hdd.resample("MS").sum(), "CDD_total": cdd.resample("MS").sum(), "Days": hdd.resample("MS").count()})
    m["Days in month"] = m.index.days_in_month
    m["Complete"] = m["Days"] == m["Days in month"]
    m["HDD per day"] = m["HDD_total"] / m["Days"]
    m["CDD per day"] = m["CDD_total"] / m["Days"]
    m.index.name = "Month"
    return m


def get_population(prev, prem=0.9):
    """Annual Texas population (million). FRED TXPOP (Census, July 1 estimates); else the saved store; else IMF/World Bank US x share."""
    try:
        r = requests.get(FRED, headers=UA, timeout=20)
        r.raise_for_status()
        d = pd.read_csv(io.StringIO(r.text))
        d.columns = ["date", "v"]
        d["v"] = pd.to_numeric(d["v"], errors="coerce")
        d = d.dropna()
        s = pd.Series(d["v"].values / 1000.0, index=pd.to_datetime(d["date"]).dt.year)
        s = s[s.index >= 2000]
        if len(s) > 10:
            return s, "FRED TXPOP (US Census Bureau resident population, July 1 estimates)"
    except Exception as exc:  # noqa: BLE001
        log("population: FRED failed:", exc)
    if prev is not None and len(prev) and "FRED" in str(prev["Source"].iloc[-1]) and "unreachable" not in str(prev["Source"].iloc[-1]):
        return prev["Population_m"], "saved copy from an earlier run (FRED not reachable): " + str(prev["Source"].iloc[-1])
    try:
        m = pd.read_excel(MACRO_XLSX, sheet_name="Population_IMF_m").set_index("year")["USA"].dropna()
        g = m.pct_change().fillna(0) + prem / 100
        lvl = {2020: 29.145}   # Texas 2020 Census, million
        for y in range(2021, int(m.index.max()) + 1):
            lvl[y] = lvl[y - 1] * (1 + g[y])
        for y in range(2019, 1999, -1):
            lvl[y] = lvl[y + 1] / (1 + g[y + 1])
        s = pd.Series(lvl).sort_index()
        return s, f"FALLBACK (FRED/Census unreachable): Texas = 2020 Census 29.145m carried with IMF WEO US population growth + {prem} pp a year (editable Texas premium); NOT Census estimates"
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f"no population source: {exc}")


def pop_monthly(idx, annual, g):
    """Linear interpolation between July-1 annual points; beyond the last point compound growth g per year."""
    pts = pd.Series(annual.values, index=[pd.Timestamp(int(y), 7, 1) for y in annual.index])
    x = np.array([p.toordinal() for p in pts.index], float)
    out = []
    for m in idx:
        mid = m + pd.Timedelta(days=m.days_in_month // 2)
        o = mid.toordinal()
        if o <= x[-1]:
            out.append(float(np.interp(o, x, pts.values)))
        else:
            out.append(float(pts.values[-1] * (1 + g) ** ((o - x[-1]) / 365.25)))
    return pd.Series(out, index=idx)


# ------------------------------------------------------------------ regression helpers
def ols(X, y):
    X, y = np.asarray(X, float), np.asarray(y, float)
    n, k = X.shape
    inv = np.linalg.pinv(X.T @ X)
    b = inv @ X.T @ y
    res = y - X @ b
    dof = max(n - k, 1)
    s2 = float(res @ res) / dof
    se = np.sqrt(np.diag(inv) * s2)
    sst = float(((y - y.mean()) ** 2).sum())
    r2 = 1 - float(res @ res) / sst
    return {"b": b, "se": se, "r2": r2, "adj": 1 - (1 - r2) * (n - 1) / dof, "s": np.sqrt(s2), "n": n, "k": k, "inv": inv, "s2": s2}


def predict(fit, X):
    X = np.asarray(X, float)
    mu = X @ fit["b"]
    se = np.sqrt(fit["s2"] * (1 + np.einsum("ij,jk,ik->i", X, fit["inv"], X)))
    return mu, se


def design(d, weather, pt, extras):
    X = pd.DataFrame({"const": 1.0}, index=d.index)
    if weather == "temp":
        X["HDD per day"] = d["hdd"]
        X["CDD per day"] = d["cdd"]
    else:
        for m in range(2, 13):
            X[f"Month {m:02d} (vs Jan)"] = (d.index.month == m).astype(float)
    if "pop" in pt:
        X["Population (million)"] = d["pop"]
    if "trend" in pt:
        X["Trend (years)"] = d["trend"]
    if "hh" in extras:
        X["Henry Hub ($/MMBtu)"] = d["hh"]
    if "re" in extras:
        X["ERCOT wind+solar share"] = d["re"]
    return X


def plausible(sector, names, b):
    ok, why = True, []
    for n, v in zip(names, b):
        if n == "HDD per day" and sector in ("Residential", "Commercial") and v <= 0:
            ok = False; why.append("HDD coefficient <= 0")
        if n == "CDD per day" and sector == "Electric power" and v <= 0:
            ok = False; why.append("CDD coefficient <= 0")
        if n == "Population (million)" and v < 0:
            ok = False; why.append("population coefficient < 0")
    return ok, "; ".join(why)


def fit_candidate(sector, d, y, weather, pt, extras):
    X = design(d, weather, pt, extras)
    rows = d.index[y.notna() & X.notna().all(axis=1)]
    last = y.dropna().index.max()
    train = rows[rows <= last - pd.DateOffset(months=HOLD)]
    test = rows[rows > last - pd.DateOffset(months=HOLD)]
    if len(train) < len(X.columns) + 12 or len(test) < 6:
        return None
    f_tr = ols(X.loc[train], y.loc[train])
    p, _ = predict(f_tr, X.loc[test])
    err = p - y.loc[test].values
    full = ols(X.loc[rows], y.loc[rows])
    ok, why = plausible(sector, list(X.columns), full["b"])
    return {"weather": weather, "pt": pt, "extras": extras, "X": X, "rows": rows, "fit": full,
            "rmse": float(np.sqrt((err ** 2).mean())), "mape": float(np.mean(np.abs(err) / y.loc[test].values) * 100),
            "bias": float(err.mean()), "ok": ok, "why": why, "test": test, "pred_test": p}


def label(c):
    w = "HDD/CDD" if c["weather"] == "temp" else "month dummies"
    pt = {"pop+trend": "population + trend", "pop": "population", "trend": "trend"}[c["pt"]]
    ex = "".join(f" + {e}" for e in c["extras"]).replace("hh", "Henry Hub").replace("re", "wind+solar share")
    return f"{w}, {pt}{ex}"


MEANING = {
    "const": "intercept (level when every driver is zero; not physical on its own)",
    "HDD per day": "Bcf/d per +1 average daily heating degree-day (18C) in the month; one cold month at 10 HDD/day vs 0 adds 10x this",
    "CDD per day": "Bcf/d per +1 average daily cooling degree-day (18C); a hot month at 10 CDD/day vs 0 adds 10x this",
    "Population (million)": "Bcf/d per +1 million Texans (with a trend in the model, the two share the same growth and are hard to separate)",
    "Trend (years)": "Bcf/d per year beyond the other drivers (efficiency, electrification, new load, anything omitted)",
    "Henry Hub ($/MMBtu)": "Bcf/d per +$1/MMBtu monthly Henry Hub (negative = switching to coal/renewables when gas is dear)",
    "ERCOT wind+solar share": "Bcf/d per +1.0 (=100 percentage points) wind+solar share; divide by 10 for +10 points",
}


def meaning(n):
    return MEANING.get(n, "Bcf/d versus January, all else equal (seasonal shape that weather does not explain)" if n.startswith("Month") else "")


# ------------------------------------------------------------------ load breakout (SOURCED large-load data vs the unexplained load)
BO, BOV, LS = "Load breakout", "Load breakout values", "Load sources"
PROD_XLSX = os.path.join(OUT_DIR, "texas_production_forecast.xlsx")
OV = "https://www.ercot.com/files/docs/"
OVN = "ERCOT-Monthly-Operational-Overview-"
# ERCOT Monthly Operational Overview (Board/System Planning), slide 'Loads Approved to Energize - Observations', read 5 Oct 2026 in an Actions probe
# (discovery_archive/ERCOT_LOAD_SOURCES_PROBE.py). (report month, approved-to-energize MW, observed NON-SIMULTANEOUS monthly peak MW, month ERCOT's
# own sentence names, file path). NB: ERCOT's text names the following month in the Sep-Dec 2025 reports (published mid-month, so a typo): the file's month is used.
LL_OBS = [
    ("2025-04", 6779, 3691, "April 2025", "2025/05/15/" + OVN + "April-2025.pdf"),
    ("2025-05", 6874, 3489, "May 2025", "2025/06/17/" + OVN + "May-2025.pdf"),
    ("2025-06", 6874, 3605, "June 2025", "2025/07/17/" + OVN + "June-2025.pdf"),
    ("2025-07", 7150, 3820, "July 2025", "2025/08/15/" + OVN + "July-2025.pdf"),
    ("2025-08", 7502, 3694, "August 2025", "2025/09/22/" + OVN + "August-2025.pdf"),
    ("2025-09", 7502, 4126, "October 2025 (sic)", "2025/10/17/" + OVN + "September-2025.pdf"),
    ("2025-10", 7502, 3960, "November 2025 (sic)", "2025/11/17/" + OVN + "October-2025.pdf"),
    ("2025-11", 7502, 3827, "December 2025 (sic)", "2025/12/16/" + OVN + "November-2025.pdf"),
    ("2025-12", 8787, 3349, "January 2026 (sic)", "2026/01/15/" + OVN + "December-2025.pdf"),
    ("2026-01", 8787, 4049, "January 2026", "2026/02/18/" + OVN + "January-2026.pdf"),
    ("2026-02", 9043, 4037, "February 2026", "2026/03/18/" + OVN + "February-2026.pdf"),
    ("2026-03", 9042, 4008, "March 2026", "2026/04/16/" + OVN + "March-2026.pdf"),
    ("2026-04", 9012, 4006, "April 2026", "2026/05/19/" + OVN + "April-2026.pdf"),
    ("2026-05", 9062, 4145, "May 2026", "2026/06/15/" + OVN + "Final-May-2026.pdf"),
    ("2026-06", 8926, 3966, "June 2026", "2026/07/17/" + OVN + "June-2026.pdf"),
    ("2026-07", 9456, 4370, "July 2026", "2026/08/17/" + OVN + "July-2026.pdf"),
    ("2026-08", 9456, 4316, "August 2026", "2026/09/16/" + OVN + "August-2026.pdf"),
]
CDR_URL = OV + "2025/12/19/CapacityDemandandReservesReport_December2025.xlsx"
# ERCOT CDR Dec 2025, tab 'LoadResourceScenarios', 'ERCOT Large Load Forecast' by type, summer 2026 cumulative new loads (MW): FORECAST, not observed
CDR_TYPES = {"Data centres": 2433.2, "Crypto mining": 2102.5, "Oil & gas": 1699.2, "Hydrogen": 1387.5, "Industrial (other)": 5.0}
DC_PEAK_RATIO = 0.498          # ERCOT 2025 LTLF report: observed data-centre site peak / requested MW, sites in service 2022-24
QBT, OSV = "Queue by type values", "Observed split values"
TAC_FEB = OV + "2026/03/05/February-TAC-Report.pdf"
TAC_MAR = OV + "2026/03/12/March-TAC-Report.pdf"
TAC_MAR_UPD = OV + "2026/03/27/March-TAC-Report-Updated_03262026.pptx"
Q_COLS = ["Data centre", "Crypto", "Industrial", "Data centre / crypto (mixed label)", "Hydrogen", "Type not given ('None')"]
# ERCOT 'Large Load Interconnection Status Update' (TAC), slide 12 'Large Load Project Distribution by Type' (an IMAGE: pie + bar; pymupdf / python-pptx extracted the picture,
# tesseract read the labels, and the picture was inspected by eye, 5 Oct 2026). The pie prints the data-centre share and its MW; the bar prints the other shares.
# (key, deck date, data-centre MW printed, printed shares %: DC, None, Crypto, Industrial, DC/Crypto, Hydrogen, url, queue total of the same deck's status chart MW, month shown on the chart or None)
QUEUE_TYPE = [
    ("2026-02-25", "25 Feb 2026", 175124, (74.4, 13.2, 7.5, 3.3, 1.1, 0.5), TAC_FEB, 237712, "2026-02"),
    ("2026-03-13", "13 Mar 2026", 183469, (77.5, 11.3, 6.9, 2.7, 1.1, 0.6), TAC_MAR, 238630, None),
    ("2026-03-26", "26 Mar 2026 (updated deck, incl. ~140 GW of new submissions)", 355830, (87.6, 6.3, 3.8, 1.3, 0.6, 0.3), TAC_MAR_UPD, None, "2026-03"),
]
# ERCOT 'Approved to Energize by Load Zone / by Project Type' (same decks, slide 5, two bar images): observed non-simultaneous peak and approved MW
# (key, deck date, url, LZ_WEST obs, LZ_WEST approved, other-zones obs, other approved, co-located obs, co-located approved, standalone obs, standalone approved, chart month or None)
OBS_SPLIT = [
    ("2026-02-25", "25 Feb 2026", TAC_FEB, 2350, 5136, 1652, 3741, 1243, 1878, 2758, 6999, "2026-02"),
    ("2026-03-13", "13 Mar 2026", TAC_MAR, 2253, 5136, 1634, 3907, 1238, 2044, 2648, 6999, None),
    ("2026-03-26", "26 Mar 2026 (updated deck)", TAC_MAR_UPD, 2364, 5136, 1644, 3907, 1329, 2044, 2677, 6999, "2026-03"),
]


def queue_type_df():
    """ERCOT queue by type in GW at the snapshots shown on the chart: data-centre MW as printed on the pie; the other types = printed share x (printed DC MW / printed DC share)
    (DERIVED; shares are rounded to 0.1 point)."""
    rows = {}
    for _k, _d, dc, sh, _u, _t, cm in QUEUE_TYPE:
        if not cm:
            continue
        tot = dc / (sh[0] / 100.0)
        dcg, none, cry, ind, mix, hyd = dc / 1000.0, *[tot * x / 100.0 / 1000.0 for x in sh[1:]]
        rows[pd.Timestamp(cm + "-01")] = dict(zip(Q_COLS, [dcg, cry, ind, mix, hyd, none]))
    d = pd.DataFrame(rows).T
    d.index.name = "Snapshot"
    return d.sort_index()


def observed_split_df():
    """ERCOT-observed non-simultaneous peak of approved large loads (GW) by load zone and by project type, at the chart snapshots (printed MW)."""
    rows = {}
    for _k, _d, _u, wo, _wa, oo, _oa, co, _ca, so, _sa, cm in OBS_SPLIT:
        if cm:
            rows[pd.Timestamp(cm + "-01")] = {"LZ_WEST (West Texas)": wo / 1000.0, "All other load zones": oo / 1000.0, "Co-located with generation": co / 1000.0, "Standalone": so / 1000.0}
    d = pd.DataFrame(rows).T
    d.index.name = "Snapshot"
    return d.sort_index()

_Q = queue_type_df().iloc[-1]                 # latest ERCOT queue snapshot by type (26 Mar 2026 updated deck), GW (derived from printed shares)
CAT_ROWS = [  # category, status, (what the references are), CDR forecast GW (None = n/a), ERCOT queue GW (None = n/a), source
    ("Large loads >=75 MW, all types (data centres, crypto, hydrogen, industrial): ERCOT-observed energised",
     "SOURCED - observed", "Observed: monthly peaks, block D. Forecast column here = approved to energise (not necessarily operating). Queue column = total queue implied by the 26 Mar 2026 pie (355,830 MW / 87.6%), all statuses", 9.456, QUEUE_TYPE[-1][2] / (QUEUE_TYPE[-1][3][0] / 100.0) / 1000.0,
     "ERCOT Monthly Operational Overview, 'Loads Approved to Energize - Observations' (block D); ERCOT TAC report (updated) 26 Mar 2026 slide 12 (queue total)"),
    ("Data centres", "NO SOURCED OBSERVED FIGURE", "CDR Dec 2025 FORECAST of new load, summer 2026; ERCOT QUEUE (requests, all statuses) Mar 2026",
     CDR_TYPES["Data centres"] / 1000, _Q["Data centre"], "ERCOT CDR Dec 2025 'LoadResourceScenarios'; ERCOT TAC report (updated) 26 Mar 2026 slide 12 (image)"),
    ("Crypto mining", "NO SOURCED OBSERVED FIGURE", "same, crypto (ERCOT's MORA also assumes crypto self-curtailment of 0.16-4.03 GW: modelled)",
     CDR_TYPES["Crypto mining"] / 1000, _Q["Crypto"], "ERCOT CDR Dec 2025; ERCOT TAC report (updated) 26 Mar 2026 slide 12; MORA Oct 2026"),
    ("Oil & gas / Permian electrification", "NO SOURCED OBSERVED FIGURE", "CDR forecast only: ERCOT's queue chart has no oil & gas label (Oncor's own queue: 4 GW, Feb 2026)",
     CDR_TYPES["Oil & gas"] / 1000, None, "ERCOT CDR Dec 2025; Oncor 8-K 26 Feb 2026 Ex 99.2 (Sempra Texas slide 22)"),
    ("Industrial incl. hydrogen (LNG terminal load not itemised by ERCOT)", "NO SOURCED OBSERVED FIGURE", "CDR: hydrogen 1.39 + industrial 0.005 GW; queue: industrial + hydrogen",
     (CDR_TYPES["Hydrogen"] + CDR_TYPES["Industrial (other)"]) / 1000, _Q["Industrial"] + _Q["Hydrogen"], "ERCOT CDR Dec 2025; ERCOT TAC report (updated) 26 Mar 2026 slide 12"),
    ("Data centre / crypto (mixed label in the queue)", "NO SOURCED OBSERVED FIGURE", "queue label 'Data Center/Crypto'; no CDR equivalent", None, _Q["Data centre / crypto (mixed label)"],
     "ERCOT TAC report (updated) 26 Mar 2026 slide 12"),
    ("Type not given ('None' in the queue)", "NO SOURCED OBSERVED FIGURE", "queue label 'None'", None, _Q["Type not given ('None')"], "ERCOT TAC report (updated) 26 Mar 2026 slide 12"),
]
SPEC_NAMES = ["ERCOT-observed large loads >=75 MW (SOURCED, observed)", "Unexplained - not attributed (no source)", "Unexplained load, 12-month mean (total)"]
SRC_HDR, SRC0 = 6, 7                     # block A: category status table
SUM_HDR, SUM0 = 18, 19                   # block B: sourced vs unattributed
DC_HDR, DC0 = 28, 29                     # block C: data centres
MON_HDR, MON0 = 41, 42                   # block D: monthly block from Jan 2022
MON_START = pd.Timestamp("2022-01-01")
CHART_FROM = pd.Timestamp("2025-04-01")
COL_TR = ["G", "H", "I"]                 # chart series columns on 'Load breakout' (add_charts.py links the chart to them)


def read_breakout_prior(out):
    """Observed large-load GW typed into column C of the committed workbook's monthly block ({Timestamp: GW}); read BEFORE the workbook is rewritten."""
    try:
        from openpyxl import load_workbook
        ws = load_workbook(out, data_only=False)[BO]
        res = {}
        for r in range(MON0, ws.max_row + 1):
            a, c = ws.cell(r, 1).value, ws.cell(r, 3).value
            if isinstance(a, (dt.datetime, pd.Timestamp)) and isinstance(c, (int, float)):
                res[pd.Timestamp(a)] = float(c)
        return res
    except Exception:  # noqa: BLE001
        return {}


def scenario_dec26():
    try:
        d = pd.read_excel(PROD_XLSX, sheet_name="Demand to 2033", index_col=0, parse_dates=True)
        return [float(d.loc[pd.Timestamp("2026-12-01"), f"Data centres GW (year-end path), {c}"]) for c in ("LOW", "BASE", "HIGH")]
    except Exception as exc:  # noqa: BLE001
        log(f"  scenario path unreadable ({type(exc).__name__}: {exc})")
        return [np.nan] * 3


def breakout_model(unexp, prior):
    """Monthly block in Python (mirrors the Excel formulas): unexplained, ERCOT-observed large loads (sourced; prior typed values fill months the code does not hold),
    12-month mean of the unexplained, the sourced part and the not-attributed remainder."""
    idx = pd.date_range(MON_START, unexp.index.max(), freq="MS")
    u = unexp.reindex(idx)
    obs = {pd.Timestamp(m + "-01"): (ap / 1000, ns / 1000, lab, f) for m, ap, ns, lab, f in LL_OBS}
    L = pd.Series({m: v[1] for m, v in obs.items()}).reindex(idx)
    for m, v in (prior or {}).items():
        if m in L.index and pd.isna(L.loc[m]):
            L.loc[m] = v
    appr = pd.Series({m: v[0] for m, v in obs.items()}).reindex(idx)
    u12 = u.rolling(12, min_periods=10).mean()      # a calendar-year window; months EIA-930 lacks (incomplete) are skipped, at least 10 of 12 needed
    ok = L.notna() & u12.notna()
    tr = pd.DataFrame({SPEC_NAMES[0]: L.where(ok), SPEC_NAMES[1]: (u12 - L).where(ok), SPEC_NAMES[2]: u12.where(ok)})
    last12 = u.dropna().index[-12:]
    U12 = float(u.loc[last12].mean())
    Lw = L.loc[last12].dropna()
    return {"u": u, "L": L, "appr": appr, "obs": obs, "tr": tr, "last12": last12, "idx": idx, "U12": U12,
            "L12": float(Lw.mean()) if len(Lw) else np.nan, "nL12": int(len(Lw)), "scen": scenario_dec26()}


def breakout_values_sheet(m):
    d = pd.DataFrame(index=m["idx"])
    d.index.name = "Month"
    d["Unexplained, monthly (GW)"] = m["u"]
    d["ERCOT-observed large loads, monthly peak (GW, SOURCED)"] = m["L"]
    for c in m["tr"].columns:
        d[c] = m["tr"][c]
    d["Row on 'Load breakout'"] = [MON0 + i for i in range(len(d))]
    return d


def write_breakout(out, m):
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    head = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    bold = Font(bold=True)
    wb = load_workbook(out)
    for nm in (BO, LS):
        if nm in wb.sheetnames:
            del wb[nm]
    ws = wb.create_sheet(BO)
    wb.move_sheet(ws, offset=wb.sheetnames.index("Assumptions") + 1 - wb.sheetnames.index(BO))
    wl = wb.create_sheet(LS)
    wb.move_sheet(wl, offset=wb.sheetnames.index(BO) + 1 - wb.sheetnames.index(LS))

    def hdr(sh, r, labels, c0=1):
        for j, t in enumerate(labels, start=c0):
            c = sh.cell(r, j, t)
            c.font, c.fill, c.alignment = bold, head, Alignment(wrap_text=True, vertical="top")

    ws["A1"] = "ERCOT unexplained load growth: what is SOURCED and what is not (GW, average load)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("NO GUESSED SHARES. The unexplained load (EIA-930 ERCOT net generation minus a weather + population baseline fitted 2019-22) is split only where a published figure exists. "
                "ERCOT publishes ONE observed large-load number: the sum of each >=75 MW load's monthly peak consumption (block D, monthly, all types together). It publishes NO observed split by data centre / crypto / "
                "oil & gas / industrial. By type it publishes only FORECASTS (CDR, block A column D) and the QUEUE of requests (TAC decks, an image chart; column E) - plans and requests, not load, so they are shown in "
                "separate columns and a separate chart ('Queue by type'), never stacked with the observed series. The only OBSERVED splits are by load zone and by connection type (charts 'Observed by zone' / 'Observed by connection'), and an IMM figure for crypto alone (4.6 GW, 2025 peak, Load sources). SOURCED = ERCOT-observed large loads; JUDGEMENT = none; the rest is 'Unexplained - not attributed (no source)'. "
                "Every figure with publisher, document, URL, as-of date and type (observed / queue / forecast / contract) is on the 'Load sources' tab, including Oncor, CenterPoint and AEP Texas disclosures and named crypto / data-centre sites.")
    ws["A3"] = ("Comparability (not like-for-like, so the remainder can be negative and is never clipped): ERCOT's figure is a PEAK (sum of each load's monthly maximum), the unexplained load is an AVERAGE; "
                "ERCOT's large loads include some energised before 2023, i.e. inside the 2019-22 baseline; and the unexplained load also holds the baseline's own error and any load below 75 MW.")
    for r in (2, 3):
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(f"A{r}:G{r}")
        ws.row_dimensions[r].height = 110 if r == 2 else 58
    ra, rb = MON0 + (m["last12"].min().year - 2022) * 12 + m["last12"].min().month - 1, MON0 + (m["last12"].max().year - 2022) * 12 + m["last12"].max().month - 1
    a0, a1 = m["last12"].min(), m["last12"].max()
    # A. status by category
    ws.cell(SRC_HDR - 1, 1, "A. Status by category (SOURCED = observed, published; everything else has no observed figure and is NOT apportioned)").font = bold
    hdr(ws, SRC_HDR, ["Category", "Status", f"Observed GW, latest 12 months ({a0:%b/%y}-{a1:%b/%y})", "ERCOT CDR FORECAST (Dec 2025), GW: planned, NOT observed", "ERCOT QUEUE by type (26 Mar 2026 updated deck), GW: requests under study, NOT load", "What the references are", "Source"])
    ws.row_dimensions[SRC_HDR].height = 48
    for i, (cat, status, what, ref, qref, src) in enumerate(CAT_ROWS):
        r = SRC0 + i
        ws.cell(r, 1, cat)
        c = ws.cell(r, 2, status)
        c.font, c.fill = bold, (green if status.startswith("SOURCED") else red)
        if i == 0:
            ws.cell(r, 3, f"=AVERAGE(C{ra}:C{rb})").number_format = "0.00"
        else:
            ws.cell(r, 3, "no sourced figure")
        ws.cell(r, 4, ref if ref is not None else "n/a").number_format = "0.00"
        ws.cell(r, 5, float(qref) if qref is not None else "n/a").number_format = "0.0"
        ws.cell(r, 6, what)
        ws.cell(r, 7, src)
        for cc in (1, 6, 7):
            ws.cell(r, cc).alignment = Alignment(wrap_text=True, vertical="top")
    r = SRC0 + len(CAT_ROWS)
    ws.cell(r, 1, "Unexplained - not attributed (no source)")
    c = ws.cell(r, 2, "NOT ATTRIBUTED")
    c.font, c.fill = bold, red
    ws.cell(r, 3, f"=C{SUM0}-C{SUM0 + 1}").number_format = "0.00"
    ws.cell(r, 6, "Unexplained load less the ERCOT-observed large loads; includes the baseline error, sub-75 MW load and the peak-vs-average difference")
    ws.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="top")
    # B. sourced vs judgement vs unattributed
    ws.cell(SUM_HDR - 1, 1, f"B. How much of the unexplained load is sourced (latest 12 months {a0:%b/%y}-{a1:%b/%y})").font = bold
    hdr(ws, SUM_HDR, ["Item", "Status", "GW", "Share of unexplained", "Note"])
    rowsB = [("Unexplained load growth (EIA-930 ERCOT minus baseline), 12-month mean", "data", f"=AVERAGE(B{ra}:B{rb})", None, "this run's regression; same months as below"),
             ("SOURCED: ERCOT-observed large loads >=75 MW (peak basis)", "SOURCED - observed", f"=AVERAGE(C{ra}:C{rb})", f"=C{SUM0 + 1}/C{SUM0}",
              f"mean over the months of the window ERCOT figures exist for ({m['nL12']} of 12)"),
             ("JUDGEMENT", "JUDGEMENT", 0, f"=C{SUM0 + 2}/C{SUM0}", "none: the earlier judgement shares were removed"),
             ("UNATTRIBUTED (no source)", "NOT ATTRIBUTED", f"=C{SUM0}-C{SUM0 + 1}-C{SUM0 + 2}", f"=C{SUM0 + 3}/C{SUM0}", "may be negative: see the comparability note in row 3")]
    for i, (a, b, c_, d_, e_) in enumerate(rowsB):
        r = SUM0 + i
        ws.cell(r, 1, a)
        x = ws.cell(r, 2, b)
        x.font = bold
        if b.startswith("SOURCED"):
            x.fill = green
        elif b != "data":
            x.fill = red
        ws.cell(r, 3, c_).number_format = "0.00"
        if d_:
            ws.cell(r, 4, d_).number_format = "0%"
        else:
            ws.cell(r, 4, 1).number_format = "0%"
        ws.cell(r, 5, e_)
    ws.cell(SUM0 + 4, 1, f"Approved to energise but not necessarily operating (latest, {max(m['obs']):%b/%y}): {m['obs'][max(m['obs'])][0]:.2f} GW (ERCOT); an approval, not a measurement - not counted above.")
    # C. data centres
    ws.cell(DC_HDR - 1, 1, "C. Data centres: no sourced observed figure, so the scenario path cannot be tested").font = bold
    hdr(ws, DC_HDR, ["Item", "Status", "LOW", "BASE", "HIGH", "Note"])
    s = m["scen"]
    rowsC = [("Observed data-centre GW", "NO SOURCED OBSERVED FIGURE", "no sourced figure", "no sourced figure", "no sourced figure",
              "ERCOT publishes no observed split of its large loads by type"),
             ("Data-centre gas burn on observed load, Bcf/d", "NO SOURCED OBSERVED FIGURE", "no sourced figure", "no sourced figure", "no sourced figure",
              "needs observed data-centre GW; the earlier figure rested on judgement shares and was removed"),
             ("Scenario path, Dec 2026 energised data-centre GW ('Demand to 2033', read only)", "SCENARIO (not data)", None if pd.isna(s[0]) else round(s[0], 2),
              None if pd.isna(s[1]) else round(s[1], 2), None if pd.isna(s[2]) else round(s[2], 2), "scenario LOW/BASE/HIGH from texas_production_forecast.xlsx; not tested against observed data"),
             ("ERCOT QUEUE: data-centre requests, all statuses, 26 Mar 2026 (updated TAC deck slide 12), GW", "QUEUE (requests, not load)", None, round(float(_Q["Data centre"]), 1), None,
              "printed on ERCOT's pie as 355,830 MW (87.6% of the queue, after ~140 GW of new submissions); 13 Mar 2026: 183,469 MW; Feb 2026: 175,124 MW"),
             ("ERCOT FORECAST of new data-centre load, summer 2026 (TSP-reported, CDR Dec 2025), GW", "FORECAST (not observed)", None, round(CDR_TYPES["Data centres"] / 1000, 2), None,
              "cumulative new loads in ERCOT's April 2025 adjusted load forecast"),
             (f"... times ERCOT's observed data-centre peak / requested MW ({DC_PEAK_RATIO:.1%}, sites in service 2022-24), GW", "DERIVED from forecast (not observed)", None,
              round(CDR_TYPES["Data centres"] / 1000 * DC_PEAK_RATIO, 2), None, "ERCOT applies this factor in its own adjusted forecast (2025 LTLF report); arithmetic, not a measurement")]
    for i, row in enumerate(rowsC):
        r = DC0 + i
        for j, v in enumerate(row):
            c = ws.cell(r, 1 + j, v)
            if j == 1:
                c.font, c.fill = bold, red
            if j in (2, 3, 4) and isinstance(v, float):
                c.number_format = "0.00"
        ws.cell(r, 1).alignment = ws.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="top")
    # D. monthly
    ws.cell(MON_HDR - 2, 1, "D. Monthly block (GW). Column C = ERCOT-observed large loads, SOURCED (yellow: type a newer ERCOT Operational Overview figure here; blank = no source). Columns G-I feed the chart").font = bold
    hdr(ws, MON_HDR, ["Month", "ERCOT unexplained load, monthly (GW)", "ERCOT-observed large loads, non-simultaneous monthly peak (GW) SOURCED", "Source (ERCOT Operational Overview, month as named in ERCOT's text)",
                      "ERCOT approved to energise (GW)", "Unexplained, 12-month mean (GW; >=10 of 12 months)"] + SPEC_NAMES)
    ws.row_dimensions[MON_HDR].height = 76
    for i, mth in enumerate(m["idx"]):
        r = MON0 + i
        ws.cell(r, 1, mth.to_pydatetime()).number_format = "mmm/yy"
        if pd.notna(m["u"].iloc[i]):
            ws.cell(r, 2, float(m["u"].iloc[i])).number_format = "0.00"
        lv = m["L"].iloc[i]
        c = ws.cell(r, 3)
        c.fill = yellow
        if pd.notna(lv):
            c.value = round(float(lv), 4)
            c.number_format = "0.000"
        if mth in m["obs"]:
            ap, ns, lab, f = m["obs"][mth]
            ws.cell(r, 4, f"{lab}: {OV}{f}")
            ws.cell(r, 5, ap).number_format = "0.00"
        elif pd.notna(lv):
            ws.cell(r, 4, "typed in the workbook (source not recorded by the script)")
        if i >= 11:
            ws.cell(r, 6, f'=IF(COUNT(B{r - 11}:B{r})>=10,AVERAGE(B{r - 11}:B{r}),"")').number_format = "0.00"
            ws.cell(r, 7, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(F{r})),C{r},"")').number_format = "0.00"
            ws.cell(r, 8, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(F{r})),F{r}-C{r},"")').number_format = "0.00"
            ws.cell(r, 9, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(F{r})),F{r},"")').number_format = "0.00"
    for col, w in zip("ABCDEFGHI", (74, 30, 22, 44, 44, 60, 40, 26, 26)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ---- Load sources tab
    wl["A1"] = "Load sources: every figure behind the Load breakout tab, and what was tried and not found"
    wl["A1"].font = Font(bold=True, size=13)
    wl["A2"] = ("Type: OBSERVED = measured by ERCOT; APPROVAL = permission to energise, not consumption; FORECAST = planned/expected, not observed; QUEUE = requests under study; MODEL ASSUMPTION = an input to an ERCOT forecast. "
                "Only OBSERVED rows enter the chart. Read from ERCOT/EIA documents by an Actions probe on 5 Oct 2026 (discovery_archive/ERCOT_LOAD_SOURCES_PROBE.py).")
    wl["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    wl.merge_cells("A2:J2")
    wl.row_dimensions[2].height = 48
    hdr(wl, 4, ["ID", "Our category", "Value", "Unit", "As-of", "Type", "Publisher", "Document", "URL", "How it maps to our categories / caveat"])
    last = max(m["obs"])
    lastv = m["obs"][last]
    mx = lambda k: OV + k  # noqa: E731
    rows = [
        ("S1", "Large loads >=75 MW, all types", round(lastv[1] * 1000), "MW", f"{last:%b %Y}", "OBSERVED", "ERCOT", "Monthly Operational Overview, 'Loads Approved to Energize - Observations': sum of each load's monthly peak (non-simultaneous)",
         mx(m["obs"][last][3]), "All types together; no split by type. Peak basis, loads >=75 MW approved through ERCOT's large-load process. Monthly series 2025-26 in block D"),
        ("S2", "Large loads >=75 MW, all types", round(lastv[0] * 1000), "MW", f"{last:%b %Y}", "APPROVAL", "ERCOT", "same", mx(m["obs"][last][3]), "Approved to energise, of which only the S1 share is observed operating"),
        ("S3", "Large loads >=75 MW, all types", 3282, "MW", "Sep 2024", "OBSERVED", "ERCOT", "LLI Queue Status Update 2024-9-6 (all-time non-simultaneous peak; simultaneous peak 2,815 MW)",
         "https://www.ercot.com/files/docs/2024/09/05/LLI%20Queue%20Status%20Update%20-%202024-9-6.pdf", "Of 5,496 MW approved: 4,421 MW standalone, 1,075 MW co-located; 2,854 MW in LZ_WEST. Different basis (all-time peak) from S1, so not in the monthly series"),
        ("S4", "Large loads >=75 MW, all types", 3801, "MW", "Mar 2026 (ERCOT text says 2025)", "OBSERVED", "ERCOT", "March TAC Report 2026-03-12: simultaneous monthly peak (non-simultaneous 3,883 MW of 9,042 MW approved)",
         mx("2026/03/12/March-TAC-Report.pdf"), "Simultaneous peak = the most large load served at one instant; ~2% below the non-simultaneous figure"),
        ("S5", "Data centres", CDR_TYPES["Data centres"], "MW", "Summer 2026 (CDR Dec 2025)", "FORECAST", "ERCOT", "Capacity, Demand and Reserves Dec 2025, tab LoadResourceScenarios: ERCOT Large Load Forecast, data centres, cumulative new loads",
         CDR_URL, "Planned additions (TSP-reported contracts + officer letters), not observed load. Total all types 7,627 MW"),
        ("S6", "Crypto mining", CDR_TYPES["Crypto mining"], "MW", "Summer 2026 (CDR Dec 2025)", "FORECAST", "ERCOT", "same", CDR_URL, "Forecast, not observed"),
        ("S7", "Oil & gas / Permian electrification", CDR_TYPES["Oil & gas"], "MW", "Summer 2026 (CDR Dec 2025)", "FORECAST", "ERCOT", "same", CDR_URL, "Forecast, not observed"),
        ("S8", "Other industrial (incl. hydrogen)", CDR_TYPES["Hydrogen"] + CDR_TYPES["Industrial (other)"], "MW", "Summer 2026 (CDR Dec 2025)", "FORECAST", "ERCOT", "same: hydrogen 1,387.5 + industrial 5.0", CDR_URL,
         "Forecast, not observed; ERCOT has no LNG-terminal category"),
        ("S9", "Data centres", DC_PEAK_RATIO, "ratio", "sites in service 2022-24", "OBSERVED", "ERCOT", "2025 Long-Term Demand and Energy Forecast report: average data-centre site peak consumption / requested MW",
         mx("2025/04/08/2025_LTLF_Report.docx"), "An observed ratio, not a GW; used only to scale the S5 forecast (block C). Officer-letter loads: 55.4% of 2024 in-service projects had energised"),
        ("S10", "Crypto mining", "0.16-4.03", "GW", "Oct 2026 MORA", "MODEL ASSUMPTION", "ERCOT", "MORA Oct 2026 (Monthly Outlook): crypto self-curtailment now modelled all year; reduces forecast system load by 160 to 4,026 MW depending on the hour",
         mx("2026/08/07/MORA_October2026.xlsx"), "Implies crypto load of at least that size at those hours, but is a forecast-method assumption, not a measurement"),
        ("S11", "Large loads (queue)", 140000, "MW", "Mar 2026", "QUEUE", "ERCOT", "March TAC Report: 137 new large-load submissions, ~140,000 MW by 2036 (not yet in the by-type chart)",
         mx("2026/03/12/March-TAC-Report.pdf"), "Requests under study, not load"),
        ("S12", "Crypto mining", 1530, "MW", "c. Feb 2024", "PROGRAM ENROLMENT", "EIA (Today in Energy)", "'Tracking electricity consumption from US cryptocurrency mining operations': ERCOT Large Flexible Load program enlisted up to 1,530 MW",
         "https://www.eia.gov/todayinenergy/detail.php?id=61364", "Enrolment in a curtailment program, not consumption; same article: ERCOT had 41 GW of crypto requests, 9 GW with planning studies approved (NERC LTRA 2023)"),
        ("S13", "Large loads", 4817, "MW", "Jan 2026 winter peak (forecast)", "FORECAST", "ERCOT", "ERCOT Adjusted Load Forecast Winter 2025-2026, tab Peak: Large Load Additions",
         mx("2025/10/06/ERCOT-Adjusted-Load-Forecast-Winter-2025-2026-for-RS-Magnitude-2025.10.07-.xlsx"), "Forecast addition to base load for the winter peak, not observed"),
    ]
    # ERCOT queue by type (image chart read by OCR + eye) - one row per type and snapshot
    qn = 14
    for key, ddate, dcmw, sh, qurl, qtot, _cm in sorted(QUEUE_TYPE, key=lambda x: x[0], reverse=True):
        tot = dcmw / (sh[0] / 100.0)
        for name, share in zip(Q_COLS, (sh[0], sh[2], sh[3], sh[4], sh[5], sh[1])):
            printed = name == "Data centre"
            mw = dcmw if printed else round(tot * share / 100.0)
            rows.append((f"S{qn}", name, mw, "MW", ddate.split(" (")[0], "QUEUE",
                         "ERCOT", f"Large Load Interconnection Status Update to TAC, {ddate}, slide 12 'Large Load Project Distribution by Type' (an image: pie + bar). Printed share {share}%"
                         + (f" and printed MW {dcmw:,}" if printed else f" (MW = printed share x {tot:,.0f} MW, the total implied by the printed data-centre MW and share: DERIVED)"),
                         qurl, "Requests under study at every status (incl. 'No Studies Submitted'), not load and not observed. "
                         + (f"Pie base is ~1% below the status chart total ({qtot:,} MW) for a reason ERCOT does not state; " if qtot else "This deck adds ~140 GW of 137 new submissions (the 13 Mar deck said they were still being processed); ")
                         + "shares are rounded to 0.1 point"))
            qn += 1
    # observed non-simultaneous peaks by zone / project type (same decks, slide 5)
    for key, ddate, ourl, wo, wa, oo, oa, co, ca, so, sa, _cm in sorted(OBS_SPLIT, key=lambda x: x[0], reverse=True):
        for name, obs_, app_, what in (("Large loads in LZ_WEST (West Texas)", wo, wa, "Approved to Energize by Load Zone"), ("Large loads in all other load zones", oo, oa, "Approved to Energize by Load Zone"),
                                        ("Large loads co-located with generation", co, ca, "Approved to Energize by Project Type"), ("Standalone large loads", so, sa, "Approved to Energize by Project Type")):
            rows.append((f"S{qn}", name, obs_, "MW", ddate.split(" (")[0], "OBSERVED", "ERCOT", f"Large Load Interconnection Status Update to TAC, {ddate}, slide 5 '{what}' (image): observed non-simultaneous peak; approved to energise {app_:,} MW",
                         ourl, "Observed split by ZONE or by CONNECTION TYPE, not by load type (no data-centre / crypto label); the two splits cover the same loads, do not add them. 13 Mar values were revised upward in the 26 Mar update"))
            qn += 1
    # IMM / ERCOT operations: observed crypto
    imm = "https://www.potomaceconomics.com/wp-content/uploads/2026/06/2025-State-of-the-Market-Report-for-ERCOT.pdf"
    for cat, val, unit, asof, typ, pub, doc_, url_, cav in [
        ("Crypto mining", 4600, "MW", "2025 (peak)", "OBSERVED", "Potomac Economics (ERCOT Independent Market Monitor)", "2025 State of the Market Report for ERCOT, June 2026, p.31 (PDF p.53): 'cryptocurrency mines accounted for a substantial aggregate load, with peak demand reaching 4,600 MW in 2025'; executive summary p.14: 'Aggregate demand from cryptocurrency mines in ERCOT is 4.6 GW'", imm,
         "Annual peak of crypto load as the IMM measures it (method not stated in the report); not comparable with ERCOT's >=75 MW large-load figure (S1), which is 4.3 GW for ALL types"),
        ("Crypto mining", 900, "MW", "2024 to 2025 (increase in the annual maximum)", "OBSERVED", "Potomac Economics (ERCOT Independent Market Monitor)", "same report p.28 (PDF p.50): 'the maximum load from cryptocurrency operations increased by more than 900 MW from 2024 to 2025'", imm, "Increase, not a level"),
        ("Crypto mining (Controllable Load Resources)", 240, "MW", "2025 (monthly average)", "OBSERVED", "Potomac Economics (ERCOT Independent Market Monitor)", "same report p.29: 'CLRs still averaged more than 240 MW of demand in 2025'; all 12 registered CLRs are crypto mines; crypto was ~56% of ERS volume in summer 2025 and 64% by the end of 2025", imm, "Only the part registered as CLR; most mines moved to the Emergency Reserve Service"),
        ("Crypto mining", 3151, "MW", "Summer 2025 (maximum reduction)", "OBSERVED (curtailment)", "ERCOT", "Summer 2025 Operational and Market Review (15 Sep 2025), p.10: 'maximum reduction from cryptomining loads was about 3,151 MW and the maximum 4CP reduction about 1,501 MW'",
         OV + "2025/09/15/12-Summer-2025-Operational-and-Market-Review.pdf", "A reduction, i.e. a lower bound on crypto load at that hour; not a level"),
        ("Large loads, all types (queue)", 238600, "MW", "Dec 2025", "QUEUE", "ERCOT", "2025 Report on Existing and Potential Electric System Constraints and Needs, p.13: 'ERCOT continues to track nearly 238.6 GW of large load interconnection requests'",
         OV + "2025/12/23/2025-Report-on-Existing-and-Potential-Electric-System-Constraints-and-Needs.pdf", "Requests; same as the TAC deck total"),
        ("Large loads, all types (queue incl. new submissions)", 410000, "MW", "Mar 2026 (quoted in the IMM's June 2026 report)", "QUEUE", "Potomac Economics (ERCOT Independent Market Monitor)", "2025 State of the Market Report p.27: 'ERCOT's load interconnection queue has grown to more than 410 GW of additional load by 2030, of which 87.6% is data centers' (cites ERCOT's updated March 2026 TAC report)", imm, "Matches the 26 Mar 2026 updated deck (355,830 MW data centres = 87.6%)"),
    ]:
        rows.append((f"S{qn}", cat, val, unit, asof, typ, pub, doc_, url_, cav)); qn += 1
    # utilities (SEC filings): Texas wires companies' own large-load disclosures
    sec = lambda cik, acc, doc: f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{doc}"  # noqa: E731
    onc_feb = sec(1193311, "000119312526073624", "d17515dex992.htm")
    onc_aug = sec(1193311, "000119312526336875", "d171086dex991.htm")
    onc_aug2 = sec(1193311, "000119312526336875", "d171086dex992.htm")
    util = [
        ("Data centres", 255000, "MW", "26 Feb 2026", "QUEUE", "Oncor (Sempra Texas) 8-K, Ex 99.2 slide 'LC&I queue nearly doubled year-over-year to 273 GW': Data Center / IT, 384 active requests", onc_feb,
         "Oncor service territory only; requests, not load. Same slide: 'at least 38 GW expected in Oncor's 2026 RTP filing'"),
        ("Oil & gas", 4000, "MW", "26 Feb 2026", "QUEUE", "same slide: Oil + Gas, 114 requests", onc_feb, "Oncor only; requests"),
        ("Industrial (manufacturing, >100 MW)", 6000, "MW", "26 Feb 2026", "QUEUE", "same slide: Manufacturing, 32 requests", onc_feb, "Oncor only; requests"),
        ("Utility + government", 3000, "MW", "26 Feb 2026", "QUEUE", "same slide: Utility + Government, 94 requests", onc_feb, "Oncor only; requests"),
        ("Other", 5000, "MW", "26 Feb 2026", "QUEUE", "same slide: Other, 26 requests; total 650 requests, 273 GW", onc_feb, "Oncor only; requests"),
        ("Data centres", 282000, "MW", "6 Aug 2026", "QUEUE", "Oncor 8-K Ex 99.1 (Q2 2026 results): 'approximately 282 gigawatts from data centers and over 16 gigawatts of load from various other industrial sectors'", onc_aug,
         "Oncor only; requests (271 GW in May 2026, 255 GW in Feb 2026). Oncor's overall queue 298 GW (Ex 99.2)"),
        ("Industrial (other sectors)", 16000, "MW", "6 Aug 2026", "QUEUE", "same sentence: 'over 16 gigawatts' of load from other industrial sectors", onc_aug, "Oncor only; requests (over 18 GW in May and Feb 2026)"),
        ("Large loads, all types (Batch Zero eligible)", 44000, "MW", "6 Aug 2026", "QUEUE / BATCH ZERO", "Oncor 8-K Ex 99.1: ~44 GW eligible as base (27 GW) or studied (17 GW) load, collateralised by $2B+; includes ~8 GW existing interconnected large load ramping up to its authorised capacity",
         onc_aug, "Oncor only; 8 GW is authorised capacity of loads already connected (not a measured consumption); no split by type"),
        ("Large loads, all types (already energised)", 8000, "MW", "6 Aug 2026", "ENERGISED - authorised capacity", "Oncor Ex 99.2 (Sempra slide): '~8 GW already energized and continuing to ramp'", onc_aug2,
         "Oncor only; authorised capacity, not measured consumption; no split by type"),
        ("Data centres", 8000, "MW", "23 Apr 2026", "PLANNED (expected energised by 2029)", "CenterPoint Energy 8-K Ex 99.1 Q1 2026: 'expecting to energize 8 gigawatts of [data centre] projects in the Greater Houston area by 2029, with 3.5 gigawatts already under construction'",
         sec(1130310, "000110465926047123", "tm2612248d1_ex99-1.htm"), "CenterPoint Houston Electric only; plan, not load"),
        ("Industrial (refining, exports, life sciences, advanced manufacturing)", 12200, "MW", "23 Apr 2026", "CONTRACTED ('firmly committed')", "same release: '12.2 gigawatts of firmly committed industrial load' (7.5 GW in Q4 2025)",
         sec(1130310, "000110465926047123", "tm2612248d1_ex99-1.htm"), "CenterPoint Houston Electric only; contract/plan, not load"),
        ("Large loads, all types (Batch Zero)", 17000, "MW", "28 Jul 2026", "QUEUE / BATCH ZERO", "CenterPoint Energy 8-K Ex 99.1 Q2 2026: 'submitted over 17 gigawatts of large load projects through ERCOT's Batch Zero process, of which approximately 14 gigawatts are expected to be eligible'",
         sec(1130310, "000110465926087279", "tm2621002d1_ex99-1.htm"), "CenterPoint Houston Electric only; no split by type"),
        ("Data centres and mega-sized developers", 45000, "MW", "30 Jun 2026", "CONTRACT (Letters of Agreement)", "AEP 10-Q: 'AEP Texas has executed Letters of Agreement for approximately 45 gigawatts of incremental load by 2030, including approximately 40 gigawatts ... Batch Zero'",
         sec(4904, "000000490426000059", "aep-20260630.htm"), "AEP Texas only; 36 GW at Feb 2026 'backed by signed letters of agreement with well-capitalized hyperscalers and mega-sized data center developers' (AEP 8-K 12 Feb 2026 Ex 99.1); contract, not load"),
    ]
    named = [  # named sites, SEC filings: company, text, MW, as-of, type, url, caveat
        ("Crypto / compute - Riot Platforms Rockdale", 700, "10-Q 30 Jun 2026", "DEVELOPED CAPACITY", "Rockdale Facility 'provides up to approximately 700 MW of developed capacity for Bitcoin Mining and data center leasing'; AMD lease 50 MW IT (up to 200); Aug 2026 lease 191 MW critical IT",
         sec(1167419, "000110465926093448", "riot-20260630x10q.htm"), "Capacity, not metered load; part is being converted to data-centre leasing"),
        ("Crypto - Riot Platforms Corsicana", 400, "10-Q 30 Jun 2026", "DEVELOPED CAPACITY", "'currently equipped to provide up to approximately 400 MW ... for Bitcoin Mining'; ~1 GW at full build-out",
         sec(1167419, "000110465926093448", "riot-20260630x10q.htm"), "Capacity, not metered load"),
        ("Crypto - Cipher Mining Odessa", 207, "10-Q 30 Jun 2026", "OPERATING CAPACITY", "'one wholly owned bitcoin mining data center, a 207 MW site located in Odessa, Texas' (fixed-price PPA); Black Pearl, Wink: 300 MW, previously bitcoin; 700 MW HPC in development over three sites",
         sec(1819989, "000181998926000041", "cifr-20260630.htm"), "Capacity of an operating site, not metered load"),
        ("Data centre - Core Scientific Denton", 297, "8-K 10 Sep 2026", "ERCOT-APPROVED (conditional)", "Denton: 297 MW 'conditionally approved by ERCOT as Base Load, Pathway (a)'; 74 MW more to be energised (in ERCOT's 2025 RTP)",
         sec(1839341, "000183934126000023", "core-20260910.htm"), "Approval/plan, not load"),
        ("Data centre - Core Scientific Hunt County", 431, "8-K 10 Sep 2026", "ERCOT-APPROVED (Batch Zero)", "Hunt: 431 MW 'conditionally approved in Batch Zero as Base Load, Advancing Large Load, Pathway (e)'; construction started",
         sec(1839341, "000183934126000023", "core-20260910.htm"), "Approval/plan, not load"),
        ("Data centre - Galaxy Helios (Haskell/Throckmorton, West Texas)", 133, "10-Q 30 Jun 2026", "DEVELOPED (critical IT load)", "'developed the first 133 MW of critical IT load, utilizing approximately 200 MW of gross power capacity, for CoreWeave'; ERCOT has approved over 1.6 GW gross capacity",
         sec(1859392, "000185939226000091", "glxy-20260630.htm"), "Critical IT MW built; 1.6 GW is approved gross capacity, not load"),
        ("Data centre - TeraWulf / Fluidstack Abernathy JV", 168, "10-Q 30 Jun 2026", "UNDER CONSTRUCTION", "JV 'will construct and operate a 168 MW critical IT load datacenter campus in Abernathy, Texas'",
         sec(1083301, "000108330126000166", "wulf-20260630.htm"), "Plan, not load"),
        ("Data centre - Hut 8 Beacon Point (Nueces County)", 352, "8-K 20 Jul 2026", "CONTRACTED (lease)", "second phase of the 1 GW Beacon Point campus: 15-year lease for 352 MW of IT capacity; 352 MW critical IT financing in 10-Q 4 Aug 2026",
         sec(1964789, "000110465926084862", "tm2620835d1_8k.htm"), "Contract/plan, not load"),
        ("Compute - IREN Childress / Sweetwater 1 / Sweetwater 2", 2750, "10-K FY to 30 Jun 2026", "PLANNED (grid connection agreements)", "'Childress, Texas 750MW; Sweetwater 1, Texas 1,400MW; Sweetwater 2, Texas 600MW' of planned power capacity (gross)",
         sec(1878848, "000187884826000052", "iren-20260630.htm"), "Plan, not load; 750 + 1,400 + 600 MW"),
        ("Crypto / compute - MARA (acquisition)", 2000, "10-Q 30 Jun 2026", "CONTRACTED power, conditional", "acquisition may be terminated if 'ERCOT does not approve the allocation of all or a portion of the contracted 2,000 MW of power to the site'",
         sec(1507605, "000150760526000022", "mara-20260630.htm"), "Contract subject to ERCOT approval; site not named in the extract"),
        ("Data centre - CleanSpark Brazoria County", 300, "10-Q 31 Mar 2026", "PLANNED (framework)", "'secured a framework for approximately 300 megawatts for power capacity, with potential expansion to approximately 600 megawatts'",
         sec(827876, "000119312526217036", "clsk-20260331.htm"), "Plan, not load"),
        ("Data centre - Constellation / CyrusOne, Freestone County", 380, "10-Q 30 Jun 2026", "CONTRACTED", "'signed a new 380 MW agreement with Dallas-based CyrusOne ... to connect and serve a new data center adjacent to the Freestone Energy Center'",
         sec(1868275, "000186827526000104", "ceg-20260630.htm"), "Contract, not load"),
    ]
    for cat, val, unit, asof, typ, doc_, url_, cav in util:
        pub = "Oncor (SEC filing)" if "Oncor" in doc_ or url_ in (onc_feb, onc_aug, onc_aug2) else ("CenterPoint (SEC filing)" if "CenterPoint" in doc_ else "AEP (SEC filing)")
        rows.append((f"S{qn}", cat, val, unit, asof, typ, pub, doc_, url_, cav)); qn += 1
    for cat, val, asof, typ, doc_, url_, cav in named:
        rows.append((f"S{qn}", cat, val, "MW", asof, typ, "Company (SEC filing)", doc_, url_, cav)); qn += 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = wl.cell(5 + i, 1 + j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j == 5:
                c.font = bold
                c.fill = green if v == "OBSERVED" else red
            if j == 2 and isinstance(v, (int, float)) and v > 1:
                c.number_format = "#,##0"
    r0 = 5 + len(rows) + 2
    wl.cell(r0, 1, "Tried; nothing usable for a category split (exactly what was found)").font = bold
    hdr(wl, r0 + 1, ["ID", "Source", "Reachable from Actions?", "What was found", "", "", "", "", "URL", "Result"])
    tried = [
        ("T1", "ERCOT TAC 'Large Load Interconnection Status Update' decks and Monthly Operational Overviews: by-type chart images (OCR + pymupdf)", "yes",
         "Feb and Mar 2026 TAC decks (February-TAC-Report.pdf, March-TAC-Report.pdf, March-TAC-Report-Updated_03262026.pptx) hold slide 12 'Large Load Project Distribution by Type' and slide 5 'Approved to Energize by Load Zone / Project Type' as pictures: pymupdf / python-pptx extracted them, tesseract read the labels, the pictures were checked by eye. Result = the QUEUE-by-type rows (requests, not load) and the OBSERVED zone / connection-type rows. The Operational Overviews (Jun 2025, Nov 2025, Jan 2026, Jun 2026, Aug 2026 scanned) carry the queue by STATUS and year only, no type; slide 'Loads Approved to Energize - Observations' is all types.",
         "https://www.ercot.com/files/docs/2026/03/12/March-TAC-Report.pdf", "queue by type for 2 months; no observed by type"),
        ("T2", "Other months of the TAC deck (Jan 2025 - Sep 2026) and the older 'LLI Queue Status Update' decks", "partly",
         "Brute-force over file names (Month-TAC-Report.pdf, -Final, TAC-Report-Month-Year; 'LLI Queue Status Update - Y-M-D.pdf', 14,128 candidate URLs): only Feb and Mar 2026 TAC reports (+ the 26 Mar 'Updated' pptx named in the IMM report) and the Jun, Jul and Sep 2024 LLI decks exist under those names. The 2024 decks (8 pages) carry status/zone charts, no by-type slide. The Operational Overviews carry no by-zone / by-type approved-to-energise slide (a text scan of the Apr 2025 - Aug 2026 decks found only credit-by-type charts). Other months are posted under names not guessed or on the TAC meeting pages, which list no documents to scripts.",
         "https://www.ercot.com/files/docs/2024/09/05/LLI%20Queue%20Status%20Update%20-%202024-9-6.pdf", "2 monthly by-type snapshots only"),
        ("T3", "ERCOT Capacity, Demand and Reserves reports: prior editions", "yes",
         "Dec 2023, May 2024 (Revised), May 2025 and Dec 2025 opened. Only Dec 2025 splits new large load by type (data centres, industrial, hydrogen, oil & gas, crypto: S5-S8); the earlier ones give only 'New Contracted Loads' and 'TSP Officer Letter Loads' totals (May 2025: contracted 4,894 MW and officer-letter 2,734 MW for summer 2026). No 'observed / in service' column in any edition. Dec 2024 and the 2026 editions are not under the file names tried (May 2026 not published yet at those names).",
         "https://www.ercot.com/gridinfo/resource", "forecast by type for one edition"),
        ("T4", "ERCOT 2025 Long-Term Load Forecast report and adjusted forecast files", "yes",
         "The docx gives the adjustment factors (data centres 49.8% of requested MW, officer letters 55.4%) and no MW by type; the adjusted-forecast workbooks are .xlsb (not read); the 2026 LTLF is not posted at the listing page.",
         "https://www.ercot.com/gridinfo/load/forecast", "no MW by type"),
        ("T5", "ERCOT Batch Zero Load Information Form, Readiness FAQs, Verification RFI exhibit list", "yes",
         "Blank templates (categories Data Centers (non-crypto), Hydrogen and Electrofuel, Cryptocurrency Mining); no aggregates.",
         "https://www.ercot.com/files/docs/2026/06/18/Batch-Zero-Load-Information-Form-06172026.xlsx", "no figures"),
        ("T6", "ERCOT MIS public reports (CLR / Large Flexible Load telemetry)", "partly",
         "MIS list endpoint answers but no CLR/LFL MW report was found; CDR only carries assumed CLR.", "https://www.ercot.com/misapp/servlets/IceDocListJsonWS", "no observed CLR/LFL MW"),
        ("T7", "SEC filings (Oncor, CenterPoint, AEP, Sempra, TXNM, Vistra, NRG, Entergy, Xcel; Riot, Cipher, Core Scientific, TeraWulf, Galaxy, MARA, CleanSpark, Hut 8, IREN, Applied Digital, Constellation)", "yes",
         "Latest 10-K/10-Q and Ex 99 earnings releases/slides read. Utilities give queues and contracts, with a category split only for Oncor (S-rows). Miners give capacity and contracts, only Cipher Odessa (207 MW) described as operating. None gives metered consumption. Bitdeer files 20-F/6-K and was not opened; Lancium / Crusoe / Stargate are private (no filings); Applied Digital's Texas site text was not matched.",
         "https://www.sec.gov/cgi-bin/browse-edgar", "queue / contract / capacity"),
        ("T8", "Cambridge CBECI mining map", "page yes, data no", "Data load client-side from Firestore (403); no Texas GW.", "https://ccaf.io/cbnsi/cbeci/mining_map", "no Texas GW"),
        ("T9", "EIA-860M, EIA Today in Energy, PUCT SB6 project 58317, Texas Comptroller, LBNL 2024 report", "partly", "EIA-860M is a generator inventory; the EIA crypto article gives S12 only; PUCT index loads but individual filings were not read; Comptroller 404; LBNL pdf answered an empty 202.",
         "https://www.eia.gov/todayinenergy/detail.php?id=61364", "no usable totals"),
        ("T10", "Grid Strategies, Brattle, E3, Texas Blockchain Council, Potomac Economics", "see log", "Opened by an Actions probe (discovery_archive/ERCOT_LARGE_LOAD_THINKTANKS_PROBE.py); nothing with MW by type for ERCOT was found beyond the items above. Any number seen only in a search snippet is NOT used.",
         "https://gridstrategiesllc.com/reports/", "none used"),
    ]
    for i, row in enumerate(tried):
        r = r0 + 2 + i
        vals = [row[0], row[1], row[2], row[3], None, None, None, None, row[4], row[5]]
        for j, v in enumerate(vals):
            c = wl.cell(r, 1 + j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        wl.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        wl.row_dimensions[r].height = 118
    for col, w in zip("ABCDEFGHIJ", (6, 34, 14, 12, 22, 20, 14, 62, 60, 62)):
        wl.column_dimensions[col].width = w
    wl.freeze_panes = "A5"
    root, ext = os.path.splitext(out)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)



# ------------------------------------------------------------------ main
def read_prev(out, sheet, **kw):
    try:
        return pd.read_excel(out, sheet_name=sheet, **kw)
    except Exception:  # noqa: BLE001
        return None


def assumptions_in(out, auto):
    """Yellow cells: keep the user's edit when it differs from last run's auto default, else take the new auto default."""
    prev = read_prev(out, "Assumptions")
    used = {}
    for k, v in auto.items():
        val = v
        if prev is not None and k in set(prev["Assumption"]):
            r = prev[prev["Assumption"] == k].iloc[0]
            try:
                pu, pa = float(r["Value used (edit)"]), float(r["Auto default (this run)"])
                if abs(pu - pa) > 1e-9 * max(1, abs(pa)):
                    val = pu
            except Exception:  # noqa: BLE001
                pass
        used[k] = val
    return used


def run(out):
    t0 = time.time()
    bprior = read_breakout_prior(out)   # editable Load breakout cells, read before the workbook is rewritten
    # ---- history inputs
    cons = pd.read_excel(GAS_XLSX, sheet_name="Consumption by sector")
    cons["Month"] = pd.to_datetime(cons["Month"])
    cons = cons.set_index("Month").sort_index()[SECTORS]
    cons = cons[cons.index >= START].dropna()
    last_gas = cons.index.max()
    log(f"gas sectors {cons.index.min():%b/%y}-{last_gas:%b/%y}")

    prev_t = read_prev(out, "Texas T2M daily", index_col=0)
    saved = pd.DataFrame()
    if prev_t is not None:
        prev_t.index = pd.to_datetime(prev_t.index)
        saved = prev_t.apply(pd.to_numeric, errors="coerce")
    temps, wsrc = update_temps(saved)
    if temps.empty:
        raise SystemExit("no temperature data (NASA POWER and Open-Meteo both failed and no saved store)")
    dd = degree_days(temps)
    ddc = dd[dd["Complete"]]
    log(f"degree days: complete months {ddc.index.min():%b/%y}-{ddc.index.max():%b/%y} ({wsrc})")

    prev_p = read_prev(out, "Texas population")
    prev_p = prev_p.set_index("Year") if prev_p is not None else None
    prem = 0.9
    pa = read_prev(out, "Assumptions")
    if pa is not None:
        r = pa[pa["Assumption"].astype(str).str.startswith("Texas population growth premium")]
        if len(r):
            prem = float(r["Value used (edit)"].iloc[0])
    annual, psrc = get_population(prev_p, prem)
    cagr = float((annual.iloc[-1] / annual.iloc[-6]) ** (1 / 5) - 1)

    hh = pd.read_excel(HH_XLSX, sheet_name="Chart - Monthly")
    hh["Date"] = pd.to_datetime(hh["Date"])
    hh = hh.set_index("Date")["Henry Hub spot"].dropna()
    hh = hh[hh.index.day == 1]

    e = pd.read_excel(E930_XLSX, sheet_name="ERCOT", usecols=["date", "Wind_MWh", "Solar_MWh", "Total_MWh"])
    e["date"] = pd.to_datetime(e["date"])
    e = e.drop_duplicates("date").set_index("date").sort_index()
    e = e[e["Total_MWh"] > 0]
    g = e.resample("MS")
    cnt = g["Total_MWh"].count()
    er = pd.DataFrame({"GWh": g["Total_MWh"].sum() / 1000, "re": (e["Wind_MWh"].fillna(0) + e["Solar_MWh"].fillna(0)).resample("MS").sum() / g["Total_MWh"].sum()})
    er = er[cnt == er.index.days_in_month]
    er["GW"] = er["GWh"] / (er.index.days_in_month * 24)
    log(f"ERCOT load {er.index.min():%b/%y}-{er.index.max():%b/%y}")

    # ---- assumptions
    last_full_year = int(ddc.index.year[ddc.groupby(ddc.index.year)["Days"].transform("count") == 12].max())
    auto = {
        "Texas population growth after the last Census year (% per year)": round(cagr * 100, 3),
        "Texas population growth premium over US, pp a year (used only if Census/FRED unreachable)": 0.9,
        "Henry Hub, forecast months ($/MMBtu)": round(float(hh.iloc[-12:].mean()), 2),
        "ERCOT wind+solar share of generation, forecast months": round(float(er["re"].iloc[-12:].mean()), 3),
        "Normal-weather window, first year": last_full_year - 9,
        "Normal-weather window, last year": last_full_year,
        "ERCOT baseline fit window, last month (YYYYMM)": 202212,
        "Prediction band, number of standard errors": 1.0,
    }
    A = assumptions_in(out, auto)
    gpop = A["Texas population growth after the last Census year (% per year)"] / 100
    hh_f = A["Henry Hub, forecast months ($/MMBtu)"]
    re_f = A["ERCOT wind+solar share of generation, forecast months"]
    y0, y1 = int(A["Normal-weather window, first year"]), int(A["Normal-weather window, last year"])
    base_end = pd.Timestamp(f"{int(A['ERCOT baseline fit window, last month (YYYYMM)']) // 100}-{int(A['ERCOT baseline fit window, last month (YYYYMM)']) % 100:02d}-01")
    nse = A["Prediction band, number of standard errors"]

    # ---- driver frame
    idx = pd.date_range(START, END, freq="MS")
    D = pd.DataFrame(index=idx)
    win = ddc[(ddc.index.year >= y0) & (ddc.index.year <= y1)]
    norm = win.groupby(win.index.month)[["HDD per day", "CDD per day"]].mean()
    D["hdd_act"] = ddc["HDD per day"].reindex(idx)
    D["cdd_act"] = ddc["CDD per day"].reindex(idx)
    last_obs = max(last_gas, er.index.max())
    D["Type"] = np.where(D.index <= last_gas, "History", "Forecast")
    D["hdd"] = [D.loc[m, "hdd_act"] if (m <= last_obs and pd.notna(D.loc[m, "hdd_act"])) else norm.loc[m.month, "HDD per day"] for m in idx]
    D["cdd"] = [D.loc[m, "cdd_act"] if (m <= last_obs and pd.notna(D.loc[m, "cdd_act"])) else norm.loc[m.month, "CDD per day"] for m in idx]
    D["pop"] = pop_monthly(idx, annual, gpop)
    D["trend"] = [(m - START).days / 365.25 for m in idx]
    D["hh"] = hh.reindex(idx)
    D.loc[D.index > hh.index.max(), "hh"] = hh_f
    D["re"] = er["re"].reindex(idx)
    D.loc[D.index > er.index.max(), "re"] = re_f
    # months inside the observed span lacking a driver stay NaN (not filled), except forecast months
    log("normals window", y0, y1, "pop growth", round(gpop * 100, 2), "HH", hh_f, "RE", re_f)

    # ---- model 1: gas sectors
    chosen, cand_rows, fit_rows, hold_rows = {}, [], [], []
    fc_idx = D.index[D.index > last_gas]
    for s in SECTORS:
        y = cons[s].reindex(idx)
        extras_set = [(), ("hh",), ("hh", "re")] if s == "Electric power" else [()]
        cands = []
        for w, pt, ex in product(("temp", "dummy"), ("pop+trend", "pop", "trend"), extras_set):
            c = fit_candidate(s, D, y, w, pt, ex)
            if c:
                cands.append(c)
        ok = [c for c in cands if c["ok"]]
        pick = min(ok or cands, key=lambda c: c["rmse"])
        naive = (y.shift(12)).loc[pick["test"]]
        naive_rmse = float(np.sqrt(((naive - y.loc[pick["test"]]) ** 2).mean()))
        pick["naive"] = naive_rmse
        pick["all_implausible"] = not ok
        chosen[s] = pick
        for c in sorted(cands, key=lambda c: c["rmse"]):
            cand_rows.append({"Sector": s, "Candidate": label(c), "Hold-out RMSE (Bcf/d)": c["rmse"], "Hold-out MAPE %": c["mape"],
                              "Hold-out bias (Bcf/d)": c["bias"], "In-sample R2": c["fit"]["r2"], "Adj R2": c["fit"]["adj"], "n": c["fit"]["n"],
                              "Signs plausible": "yes" if c["ok"] else "NO: " + c["why"], "Chosen": "CHOSEN" if c is pick else ""})
        f = pick["fit"]
        ybar = float(y.dropna().iloc[-12:].mean())
        Xf = design(D.loc[fc_idx], pick["weather"], pick["pt"], pick["extras"])
        mu, se = predict(f, Xf)
        pick["fc_mu"], pick["fc_se"] = mu, se
        pick["fc_flag"] = ""
        a33 = float(pd.Series(mu, index=fc_idx)[fc_idx.year == 2033].mean())
        if abs(a33 / ybar - 1) > 0.5:
            pick["fc_flag"] = f"2033 average {a33:.2f} is {a33 / ybar - 1:+.0%} vs the last 12 months: trend/population extrapolation, treat with caution"
        for n_, b_, se_ in zip(pick["X"].columns, f["b"], f["se"]):
            fit_rows.append({"Model": f"{s} (gas, Bcf/d)", "Specification": label(pick), "Term": n_, "Coefficient": b_, "Std error": se_,
                             "t-stat": b_ / se_ if se_ else np.nan, "What it means": meaning(n_), "n": f["n"], "R2": f["r2"], "Adj R2": f["adj"],
                             "Residual std": f["s"], "Hold-out RMSE (last 12m)": pick["rmse"], "Hold-out MAPE %": pick["mape"],
                             "Naive same-month-last-year RMSE": naive_rmse, "Flags": ("ALL candidates implausible: " if not ok else "") + pick["fc_flag"]})
        for m, a, p in zip(pick["test"], y.loc[pick["test"]].values, pick["pred_test"]):
            hold_rows.append({"Month": m, "Model": s, "Actual": a, "Predicted (fit before the hold-out)": p, "Error": p - a})
        log(f"{s}: {label(pick)} R2={f['r2']:.3f} holdout RMSE={pick['rmse']:.3f} ({pick['mape']:.1f}%) naive={naive_rmse:.3f} {pick['fc_flag']}")

    # fitted / forecast frames
    F = pd.DataFrame(index=idx)
    F.index.name = "Month"
    F["Type"] = D["Type"]
    F["HDD per day (actual; normal in forecast months)"] = D["hdd"]
    F["CDD per day (actual; normal in forecast months)"] = D["cdd"]
    F["Texas population (million)"] = D["pop"]
    F["Henry Hub ($/MMBtu; assumption in forecast months)"] = D["hh"]
    F["ERCOT wind+solar share"] = D["re"]
    var_sum = pd.Series(0.0, index=idx)
    tot_act, tot_fit = cons.sum(axis=1).reindex(idx), pd.Series(np.nan, index=idx)
    parts = []
    for s in SECTORS:
        c = chosen[s]
        Xall = design(D, c["weather"], c["pt"], c["extras"])
        okr = Xall.notna().all(axis=1)
        mu = pd.Series(np.nan, index=idx)
        se = pd.Series(np.nan, index=idx)
        m_, s_ = predict(c["fit"], Xall[okr])
        mu[okr], se[okr] = m_, s_
        F[f"{s} actual"] = cons[s].reindex(idx)
        F[f"{s} fitted (actual weather) / forecast (normal weather)"] = mu
        F[f"{s} lower"] = mu - nse * se
        F[f"{s} upper"] = mu + nse * se
        parts.append((mu, se))
    mu_sum = sum(p[0] for p in parts)
    se_sum = np.sqrt(sum(p[1] ** 2 for p in parts))
    F["Other sectors actual (sum of four)"] = tot_act
    F["Other sectors fitted / forecast (sum of four)"] = mu_sum
    F["Other sectors lower (band in quadrature, sectors independent)"] = mu_sum - nse * se_sum
    F["Other sectors upper (band in quadrature, sectors independent)"] = mu_sum + nse * se_sum
    fv = pd.read_excel(GAS_XLSX, sheet_name="Forecast values")
    fv["Month"] = pd.to_datetime(fv["Month"])
    fv = fv.set_index("Month")
    ex = fv[["Electric power", "Industrial", "Residential", "Commercial"]].sum(axis=1)
    F["Existing damped-trend forecast, four sectors (texas_gas_monthly, to Dec 2028)"] = ex.where(fv["Actual / forecast"].eq("Forecast")).reindex(idx)

    # ---- model 2: ERCOT load
    Er = D.join(er[["GW"]]).copy()
    Er = Er[(Er.index <= er.index.max()) | (Er.index > er.index.max())]
    y = Er["GW"]
    Xa = design(Er, "temp", "pop", ())
    rows_a = Er.index[(Er.index <= base_end) & y.notna() & Xa.notna().all(axis=1)]
    fa = ols(Xa.loc[rows_a], y.loc[rows_a])
    obs = Er.index[y.notna() & Xa.notna().all(axis=1)]
    mu_a, se_a = predict(fa, Xa[Xa.notna().all(axis=1)])
    base = pd.Series(np.nan, index=idx)
    bse = pd.Series(np.nan, index=idx)
    ok_a = Xa.notna().all(axis=1)
    base[ok_a], bse[ok_a] = mu_a, se_a
    unexp = (y - base).where(y.notna())
    unexp_roll = unexp.rolling(12, min_periods=12).mean()
    post = unexp[(unexp.index >= "2023-01-01") & unexp.notna()]
    sl = ols(np.column_stack([np.ones(len(post)), (post.index - pd.Timestamp("2023-01-01")).days / 365.25]), post.values)
    last12 = unexp.dropna().iloc[-12:]
    stepup = float(last12.mean())
    # reference: full-sample model with trend and hold-out
    Xb = design(Er, "temp", "pop+trend", ())
    last_er = y.dropna().index.max()
    rows_b = Er.index[y.notna() & Xb.notna().all(axis=1)]
    trb = rows_b[rows_b <= last_er - pd.DateOffset(months=HOLD)]
    teb = rows_b[rows_b > last_er - pd.DateOffset(months=HOLD)]
    fb_tr = ols(Xb.loc[trb], y.loc[trb])
    pb, _ = predict(fb_tr, Xb.loc[teb])
    rmse_b = float(np.sqrt(((pb - y.loc[teb].values) ** 2).mean()))
    mape_b = float(np.mean(np.abs(pb - y.loc[teb].values) / y.loc[teb].values) * 100)
    fb = ols(Xb.loc[rows_b], y.loc[rows_b])
    pre = rows_a
    err_in = (y.loc[pre] - base.loc[pre])
    corr_pt = float(np.corrcoef(Er["pop"], Er["trend"])[0, 1])
    F["ERCOT load actual (GW, EIA-930 net generation)"] = y
    F["ERCOT baseline (weather + population, fit on window; actual weather, normal weather in forecast months)"] = base
    F["ERCOT baseline lower"] = base - nse * bse
    F["ERCOT baseline upper"] = base + nse * bse
    F["ERCOT unexplained load growth (GW) = actual - baseline"] = unexp
    F["ERCOT unexplained, 12-month average (GW)"] = unexp_roll
    F["ERCOT baseline + unexplained held at last-12-month level (reference)"] = (base + stepup).where(base.index > last_er)
    for n_, b_, se_ in zip(Xa.columns, fa["b"], fa["se"]):
        fit_rows.append({"Model": "ERCOT load baseline (GW)", "Specification": f"HDD/CDD + population, fit {rows_a.min():%b/%y}-{rows_a.max():%b/%y}", "Term": n_,
                         "Coefficient": b_, "Std error": se_, "t-stat": b_ / se_ if se_ else np.nan,
                         "What it means": meaning(n_).replace("Bcf/d", "GW average load"), "n": fa["n"], "R2": fa["r2"], "Adj R2": fa["adj"],
                         "Residual std": fa["s"], "Hold-out RMSE (last 12m)": np.nan, "Hold-out MAPE %": np.nan, "Naive same-month-last-year RMSE": np.nan,
                         "Flags": "baseline is extrapolated beyond the window; the gap to actual is the unexplained growth" + ("; population coefficient negative" if fa["b"][3] < 0 else "")})
    for n_, b_, se_ in zip(Xb.columns, fb["b"], fb["se"]):
        fit_rows.append({"Model": "ERCOT load, full sample with trend (reference)", "Specification": f"HDD/CDD + population + trend, {rows_b.min():%b/%y}-{rows_b.max():%b/%y}", "Term": n_,
                         "Coefficient": b_, "Std error": se_, "t-stat": b_ / se_ if se_ else np.nan,
                         "What it means": meaning(n_).replace("Bcf/d", "GW average load"), "n": fb["n"], "R2": fb["r2"], "Adj R2": fb["adj"],
                         "Residual std": fb["s"], "Hold-out RMSE (last 12m)": rmse_b, "Hold-out MAPE %": mape_b, "Naive same-month-last-year RMSE": np.nan,
                         "Flags": f"population and trend correlation {corr_pt:.3f}: their coefficients are not separately identified; the trend term absorbs the step-up"})
    for m, a, p in zip(teb, y.loc[teb].values, pb):
        hold_rows.append({"Month": m, "Model": "ERCOT load (full-sample model with trend)", "Actual": a, "Predicted (fit before the hold-out)": p, "Error": p - a})
    unexpl = {"slope": float(sl["b"][1]), "slope_se": float(sl["se"][1]), "last12": stepup, "since": post.index.min(), "n": len(post),
              "peak": float(unexp_roll.dropna().iloc[-1]), "first": float(post.iloc[:12].mean())}
    log(f"ERCOT baseline R2={fa['r2']:.3f} resid std={fa['s']:.2f} GW; unexplained last12={stepup:.2f} GW, trend since 2023 {unexpl['slope']:.2f}+/-{unexpl['slope_se']:.2f} GW/yr; full-sample holdout RMSE {rmse_b:.2f} GW ({mape_b:.1f}%)")

    # ---- load breakout (apportioned unexplained load; the shares are judgement)
    bm = breakout_model(unexp.dropna(), bprior)
    log(f"load breakout: unexplained latest-12-month mean {bm['U12']:.2f} GW; ERCOT-observed large loads (sourced) {bm['L12']:.2f} GW over {bm['nL12']} of 12 months; "
        f"not attributed {bm['U12'] - bm['L12']:.2f} GW")

    # ---- summary numbers
    def ann(col, yr):
        s_ = F[col][F.index.year == yr]
        return float(s_.mean())
    summ = []
    oc, ec = "Other sectors fitted / forecast (sum of four)", "Existing damped-trend forecast, four sectors (texas_gas_monthly, to Dec 2028)"
    for yr in (2026, 2028, 2033):
        dec = pd.Timestamp(f"{yr}-12-01")
        summ.append({"Period": f"Dec {yr}", "Regression other sectors (Bcf/d)": F.loc[dec, oc],
                     "Regression lower": F.loc[dec, "Other sectors lower (band in quadrature, sectors independent)"],
                     "Regression upper": F.loc[dec, "Other sectors upper (band in quadrature, sectors independent)"],
                     "Existing forecast, four sectors (Bcf/d)": F.loc[dec, ec],
                     "ERCOT baseline (GW)": F.loc[dec, "ERCOT baseline (weather + population, fit on window; actual weather, normal weather in forecast months)"]})
        summ.append({"Period": f"Year {yr} average", "Regression other sectors (Bcf/d)": ann(oc, yr),
                     "Existing forecast, four sectors (Bcf/d)": ann(ec, yr) if yr <= 2028 else np.nan,
                     "ERCOT baseline (GW)": ann("ERCOT baseline (weather + population, fit on window; actual weather, normal weather in forecast months)", yr)})
    S = pd.DataFrame(summ).set_index("Period")
    log(S.round(2).to_string())

    # ---- workbook
    fit_df = pd.DataFrame(fit_rows).set_index("Model")
    cand_df = pd.DataFrame(cand_rows).set_index("Sector")
    hold_df = pd.DataFrame(hold_rows).set_index("Month")
    notes_a = {
        "Texas population growth after the last Census year (% per year)": ("% per year", f"Applies after the last annual point ({int(annual.index.max())}). Default = 5-year CAGR of the Census series ({cagr * 100:.2f}%). The IMF outlook for the whole US is far lower; Texas has grown faster than the US."),
        "Texas population growth premium over US, pp a year (used only if Census/FRED unreachable)": ("pp", "Fallback only: Texas grew about 1.4% a year 2015-24 against about 0.5% for the US (from memory, unverified), so the default is 0.9 pp on top of IMF's US growth."),
        "Henry Hub, forecast months ($/MMBtu)": ("$/MMBtu", "Held flat in forecast months; default = mean of the last 12 months of the Henry Hub monthly series. Only the electric-power model uses it, if the chosen candidate includes it."),
        "ERCOT wind+solar share of generation, forecast months": ("share", "Held flat; default = last-12-month mean. Rising wind/solar/battery output would lower gas burn: a flat share is a conservative assumption. Used only if the chosen electric-power candidate includes it."),
        "Normal-weather window, first year": ("year", "Normal weather = average HDD/CDD per day by calendar month over these complete years (default: the last 10)."),
        "Normal-weather window, last year": ("year", "Last complete year in the temperature store."),
        "ERCOT baseline fit window, last month (YYYYMM)": ("YYYYMM", "The ERCOT baseline (weather + population) is fit on data up to this month and extrapolated; default Dec 2022, before the recent step-up in load."),
        "Prediction band, number of standard errors": ("x", "Band = fitted/forecast +/- this x the prediction standard error (residual noise + coefficient uncertainty)."),
    }
    assum = pd.DataFrame([{"Assumption": k, "Value used (edit)": A[k], "Auto default (this run)": auto[k], "Unit": notes_a[k][0], "Note": notes_a[k][1]} for k in auto]).set_index("Assumption")
    popdf = pd.DataFrame({"Population_m": annual, "Source": psrc}); popdf.index.name = "Year"
    dd_out = dd.copy()
    sheets = {"Fit summary": fit_df, "Candidates": cand_df, "Assumptions": assum, "Summary": S, "Forecast": F, "Hold-out": hold_df,
              "Load breakout values": breakout_values_sheet(bm), QBT: queue_type_df(), OSV: observed_split_df(), "Texas degree days": dd_out, "Texas population": popdf, "Texas T2M daily": temps}
    sel = "; ".join(f"{s}: {label(chosen[s])} (R2 {chosen[s]['fit']['r2']:.2f}, hold-out RMSE {chosen[s]['rmse']:.2f} Bcf/d = {chosen[s]['mape']:.1f}%)" for s in SECTORS)
    notes = [
        "Notes", "", "UNITS",
        "Gas sectors: Bcf/d (billion cubic feet per day), monthly average (EIA monthly volumes / days, texas_gas_monthly.xlsx 'Consumption by sector'). ERCOT load: GW, monthly average of EIA-930 daily ERCOT net generation (the repo has no ERCOT demand series; net generation is the proxy). Degree days: HDD/CDD per day, base 18C, population-weighted over Houston, Dallas-Fort Worth, San Antonio, Austin, El Paso (weights are approximate 2020 metro populations). Population: million.",
        "", "METHOD",
        "Ordinary least squares, monthly, from Jan 2015 (gas) / Jan 2019 (ERCOT). Standalone: no existing Texas workbook or script is changed; texas_gas_monthly.xlsx is read for history and, for the comparison chart only, its 'Forecast values' damped-trend forecast.",
        "Gas sectors: for each sector every combination of (HDD+CDD per day | month dummies) x (population+trend | population | trend), and for electric power also +Henry Hub and +Henry Hub + ERCOT wind+solar share, was fitted. The CHOSEN specification has the lowest RMSE on a hold-out of the last 12 months (fit on the months before) among candidates with plausible signs (HDD>0 residential/commercial, CDD>0 electric power, population>=0). One 12-month hold-out is a noisy basis for choosing, and the choice is made on the same data as the fit - read hold-out errors as optimistic. 'Naive' on the Fit summary is the error of simply repeating the same month a year earlier.",
        "Chosen: " + sel,
        "Population and a linear trend are nearly collinear over 2015-2026 (correlation 0.99+), so with both in the model neither coefficient is separately meaningful and the 2033 extrapolation adds both; candidates with population only or trend only are on the Candidates tab, and any chosen forecast whose 2033 average departs more than 50% from the last 12 months is flagged on the Fit summary.",
        f"Population: {psrc}; monthly values are linear interpolation between July-1 annual points, then the editable growth rate after the last point. Weather: {wsrc}; incremental store on 'Texas T2M daily' (this workbook only; macro_drivers.xlsx is untouched); a month enters only when every day has data for all five cities.",
        "ERCOT: the baseline is HDD + CDD + population fitted on the window ending at the Assumptions cell (default Dec 2022) and extrapolated at actual weather; actual minus baseline is the 'unexplained load growth', which includes data centres, crypto mining, electrified industry/oil and gas load, and anything else the baseline omits, plus the population coefficient's own error (4 years of data only). The reference full-sample model with a trend is on the Fit summary with its hold-out.",
        f"ERCOT unexplained load growth: last 12 months {unexpl['last12']:.2f} GW above the baseline; fitted trend since Jan 2023 {unexpl['slope']:+.2f} +/- {unexpl['slope_se']:.2f} GW per year (OLS on {unexpl['n']} months from {unexpl['since']:%b/%y}).",
        "Forecast to Dec 2033: normal weather (editable window), population at the editable growth, Henry Hub and wind+solar share at the editable values; the band is the prediction standard error of each model (the 'Other sectors' band adds the four in quadrature assuming independent errors, which understates it if errors are correlated, as common weather misses are). The forecast is a weather-normalised projection of past relationships; it does NOT add new LNG, data-centre or industrial demand, structural change in power supply or price response beyond what is in the fit. The ERCOT baseline excludes the step-up; the 'baseline + unexplained held flat' column is a reference only.",
        f"LOAD BREAKOUT ('Load breakout', 'Load sources' and 'Load breakout values' tabs, chart 'Load breakout'): NO judgement shares. The only observed split of the unexplained load is ERCOT's own monthly figure for large loads >=75 MW (sum of each load's monthly peak, all types together; ERCOT Operational Overviews Apr 2025 - Aug 2026), a PEAK against an AVERAGE and partly inside the 2019-22 baseline, so the remainder can be negative and is never clipped. Latest 12 months: unexplained {bm['U12']:.2f} GW, SOURCED observed large loads {bm['L12']:.2f} GW ({bm['nL12']} of 12 months), JUDGEMENT 0, UNATTRIBUTED (no source) {bm['U12'] - bm['L12']:.2f} GW. ERCOT publishes no observed GW for data centres, oil & gas or industrial separately. By type it gives only FORECASTS (CDR Dec 2025) and the QUEUE of requests (TAC decks, slide 12, an image read by OCR and by eye: data centres 355.8 GW = 87.6% in the 26 Mar 2026 updated deck, crypto ~15 GW; requests, not load), charted separately as 'Queue by type'; its only observed splits are by load zone and connection type ('Observed by zone' / 'Observed by connection', Feb and Mar 2026). Observed crypto alone: the ERCOT IMM puts crypto peak demand at 4.6 GW in 2025 (Load sources, not charted with the ERCOT figure because the definitions differ). Oncor, CenterPoint and AEP Texas queue / contract disclosures by category and named crypto / data-centre sites (SEC filings) are on 'Load sources' with their status. Data-centre GW and gas burn on observed load: no sourced figure, so the scenario-path comparison stays removed. Probes: discovery_archive/ERCOT_LOAD_SOURCES_PROBE.py, ERCOT_LARGE_LOAD_BYTYPE_PROBE*.py, ERCOT_LARGE_LOAD_COMPANIES_PROBE*.py, ERCOT_IMM_LARGE_LOAD_PROBE*.py, ERCOT_APPROVED_BYTYPE_PROBE.py (5 Oct 2026).",
        "Edit the yellow cells on the Assumptions tab and re-run the workflow (texas_demand_regression.yml, 1st/15th) to recompute; a cell left at its auto default follows the new auto value on the next run, an edited cell is kept.",
        f"Latest history: gas {last_gas:%b/%y}, ERCOT {er.index.max():%b/%y}, weather store to {temps.dropna(how='all').index.max():%d %b %Y}. Run time {time.time() - t0:.0f}s.",
        "", "SOURCES",
        "EIA (gas consumption, via texas_gas_monthly.xlsx); EIA-930 (ERCOT, via eia930_fuel_mix_daily.xlsx); Henry Hub (EIA, via henry_hub_daily.xlsx); NASA POWER / Open-Meteo (temperature); US Census Bureau via FRED TXPOP (population); ERCOT Monthly Operational Overviews, TAC Large Load Interconnection Status Updates, Capacity Demand and Reserves report, LTLF report, MORA, Summer 2025 Operational and Market Review; ERCOT IMM (Potomac Economics) 2025 State of the Market Report; Oncor / Sempra, CenterPoint, AEP and crypto / data-centre company SEC filings (large loads; URLs on 'Load sources'); EIA Today in Energy.",
    ]
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes, ["UNITS", "METHOD", "SOURCES"])
    style(out)
    write_breakout(out, bm)
    log("wrote", out)
    return {"chosen": chosen, "S": S, "unexpl": unexpl, "fa": fa, "rmse_b": rmse_b, "mape_b": mape_b}


def style(out):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    wb = load_workbook(out)
    yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    ws = wb["Assumptions"]
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).fill = yellow
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["E"].width = 120
    for name, widths in {"Fit summary": {"A": 34, "B": 52, "C": 28, "G": 100, "O": 70}, "Candidates": {"A": 16, "B": 60, "J": 40}}.items():
        w = wb[name]
        for k, v in widths.items():
            w.column_dimensions[k].width = v
        for c in w[1]:
            c.font = Font(bold=True)
        w.freeze_panes = "B2"
    wb.save(out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args()
    run(a.out)
