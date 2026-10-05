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
Edit the yellow cells of the Assumptions tab (committed workbook is read back on every run) and re-run.

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


def get_population(prev):
    """Annual Texas population (million). FRED TXPOP (Census, July 1 estimates); else the saved store; else IMF/World Bank US x share."""
    try:
        r = requests.get(FRED, headers=UA, timeout=60)
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
    if prev is not None and len(prev):
        return prev["Population_m"], "saved copy from an earlier run (FRED not reachable): " + str(prev["Source"].iloc[-1])
    try:
        m = pd.read_excel(MACRO_XLSX, sheet_name="Population_IMF_m").set_index("year")["USA"].dropna()
        return m * TEXAS_SHARE_OF_US_2020, "IMF WEO US population x Texas 2020 share (8.8%); FRED/Census unreachable"
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
    annual, psrc = get_population(prev_p)
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
              "Texas degree days": dd_out, "Texas population": popdf, "Texas T2M daily": temps}
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
        "Edit the yellow cells on the Assumptions tab and re-run the workflow (texas_demand_regression.yml, 1st/15th) to recompute; a cell left at its auto default follows the new auto value on the next run, an edited cell is kept.",
        f"Latest history: gas {last_gas:%b/%y}, ERCOT {er.index.max():%b/%y}, weather store to {temps.dropna(how='all').index.max():%d %b %Y}. Run time {time.time() - t0:.0f}s.",
        "", "SOURCES",
        "EIA (gas consumption, via texas_gas_monthly.xlsx); EIA-930 (ERCOT, via eia930_fuel_mix_daily.xlsx); Henry Hub (EIA, via henry_hub_daily.xlsx); NASA POWER / Open-Meteo (temperature); US Census Bureau via FRED TXPOP (population).",
    ]
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes, ["UNITS", "METHOD", "SOURCES"])
    style(out)
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
