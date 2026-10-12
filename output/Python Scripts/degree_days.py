"""
Population-weighted heating and cooling degree days per country, one workbook per REGION (config driven).

    python3 weather/degree_days.py --region europe [--budget-min 100] [--out-dir "output/Data and Chart Outputs"]
    -> "<out-dir>/<region>_degree_days.xlsx"  (+ the daily temperature store, see below)

INTERFACE (to add a region, e.g. 'asia': no code change)
  weather/regions.json           one entry per region: title, cities / countries csv names, season_start_month (first
                                 month of the cumulative-HDD season chart: 10 = Oct-Sep for northern regions, 4 for southern
                                 ones whose heating season is Apr-Sep), aggregates {KEY: name}, charts {cumulative_hdd,
                                 monthly_cdd, anomaly_hdd: [country keys or aggregate keys]}, source_line.
  weather/cities_<region>.csv    one row per city: country_code (the key, ISO2 and GB for the UK), country, city, lat, lon,
                                 population (blank = not sourced), population_year, population_source, url, note.  The 4 most
                                 populous cities per country by an OFFICIAL source.  Weights = population share of the cities
                                 of the country; if ANY population of a country is blank, all its cities get equal weight and
                                 the Cities sheet says weight_basis = ASSUMPTION.
  weather/countries_<region>.csv country_code, country, population, population_year, source, url, aggregates (';'-separated
                                 aggregate keys the country belongs to; the aggregate is weighted by these populations).
  weather/BUILD_CITIES_<REGION>.py (optional, run once in Actions) builds the two csv files from the official sources.

STORE: NASA POWER daily point API, parameter T2M (daily mean 2 m air temperature, MERRA-2 grid cell of the city's lat/lon),
  one request per city for the whole span, from 1991-01-01 to the latest day NASA has. The history store is
  "<out-dir>/weather_stores/<region>/city_t2m_<year>.csv.gz" (one file per year, so only the files of the years that change
  are rewritten and the git history stays small). Each run fetches every city from its last stored day minus 10 days (the first
  run backfills 1991 onward); --budget-min stops starting new cities after that many minutes (the next run continues).
  T2M_MAX / T2M_MIN are NOT stored (they would triple the store; degree days use the daily mean).

DEGREE DAYS per city and day, then the population-weighted country mean (a day counts only when every city of the country has
  a temperature; NASA gaps stay gaps, nothing is filled):
  18C       repo convention: HDD = max(0, 18 - Tmean), CDD = max(0, Tmean - 18)
  EUROSTAT  HDD = 18 - Tm when Tm <= 15 C else 0; CDD = Tm - 21 when Tm >= 24 C else 0
  Aggregates (e.g. EUROPE) = mean of the member countries weighted by their official population, on days when every member has
  a value.

SHEETS: Units, Cities, Countries, Daily HDD, Daily CDD (from DAILY_FROM; columns '<key> 18C' / '<key> EUROSTAT'), Monthly
  (sums of complete months, from 1991), Normals month and Normals doy (1991-2020 means; doy = calendar day, 29 Feb included),
  Anomaly (monthly actual minus normal, '... abs', and % of normal, blank where the normal is under 1 degree day), Coverage
  (days present per country and month), City coverage. Charts are drawn by add_charts.py (weather/weather_charts.py).
"""
import argparse
import calendar
import datetime as dt
import gzip
import io
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
NASA = "https://power.larc.nasa.gov/api/temporal/daily/point"
START = dt.date(1991, 1, 1)
REVISION_DAYS = 10
NORMAL_YEARS = (1991, 2020)
DAILY_FROM = "2016-01-01"
DEFS = ("18C", "EUROSTAT")
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def load_regions():
    return json.load(open(os.path.join(HERE, "regions.json"), encoding="utf-8"))


def load_region(region):
    cfg = load_regions()[region]
    cities = pd.read_csv(os.path.join(HERE, cfg["cities"]), dtype={"country_code": str}, keep_default_na=False, na_values=[""])
    countries = pd.read_csv(os.path.join(HERE, cfg["countries"]), dtype={"country_code": str}, keep_default_na=False, na_values=[""])
    cities["city_id"] = cities["country_code"] + "|" + cities["city"]
    # weights: population share of the country's cities; equal + ASSUMPTION if any population is not sourced
    w, basis = [], []
    for cc, g in cities.groupby("country_code", sort=False):
        if g["population"].notna().all() and (g["population"] > 0).all():
            share = g["population"] / g["population"].sum()
            b = "POPULATION"
        else:
            share = pd.Series(1.0 / len(g), index=g.index)
            b = "ASSUMPTION (equal weight: population not sourced)"
        w.append(share)
        basis.append(pd.Series(b, index=g.index))
    cities["weight"] = pd.concat(w).reindex(cities.index)
    cities["weight_basis"] = pd.concat(basis).reindex(cities.index)
    return cfg, cities, countries


# ---------------------------------------------------------------- store
def store_dir(out_dir, region):
    return os.path.join(out_dir, "weather_stores", region)


def load_store(path):
    frames = []
    if os.path.isdir(path):
        for f in sorted(os.listdir(path)):
            if f.startswith("city_t2m_") and f.endswith(".csv.gz"):
                frames.append(pd.read_csv(os.path.join(path, f), index_col=0, parse_dates=True))
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames).sort_index()
    df.index.name = "date"
    return df[~df.index.duplicated(keep="last")]


def save_store(df, path, before):
    os.makedirs(path, exist_ok=True)
    n = 0
    for y, sub in df.groupby(df.index.year):
        old = before[before.index.year == y] if len(before) else before
        sub = sub.dropna(how="all")
        if len(old) and old.shape == sub.shape and old.columns.equals(sub.columns) and np.allclose(
                old.to_numpy(float), sub.to_numpy(float), equal_nan=True):
            continue
        buf = io.StringIO()
        sub.round(2).to_csv(buf)
        tmp = os.path.join(path, f"city_t2m_{y}.csv.gz.tmp")
        with open(tmp, "wb") as raw, gzip.GzipFile(fileobj=raw, mode="wb", mtime=0, compresslevel=9) as gz:
            gz.write(buf.getvalue().encode())
        os.replace(tmp, os.path.join(path, f"city_t2m_{y}.csv.gz"))
        n += 1
    return n


def fetch_nasa(lat, lon, start, end, tries=4):
    import requests   # only the pull needs it (the master imports this module without it)
    params = {"parameters": "T2M", "community": "RE", "longitude": lon, "latitude": lat,
              "start": start.strftime("%Y%m%d"), "end": end.strftime("%Y%m%d"), "format": "JSON"}
    last = None
    for k in range(tries):
        try:
            r = requests.get(NASA, params=params, headers=UA, timeout=300)
            if r.status_code == 200:
                t = r.json()["properties"]["parameter"]["T2M"]
                s = pd.Series({pd.Timestamp(d): (np.nan if v is None or v <= -998 else float(v)) for d, v in t.items()})
                return s.dropna()
            last = f"HTTP {r.status_code}: {r.text[:150]}"
        except Exception as exc:  # noqa: BLE001
            last = repr(exc)
        time.sleep(5 * (k + 1))
    raise RuntimeError(last)


def update_store(cities, saved, budget_min, workers=4):
    end = dt.date.today() - dt.timedelta(days=1)
    t0 = time.time()
    jobs = []
    for r in cities.itertuples():
        cid = r.city_id
        if cid in saved and saved[cid].notna().any():
            start = max(START, saved[cid].dropna().index.max().date() - dt.timedelta(days=REVISION_DAYS))
        else:
            start = START
        if start <= end:
            jobs.append((start, cid, r.lat, r.lon))
    jobs.sort(key=lambda j: j[0])   # oldest first, so a budget-limited run fills the most outdated cities
    print(f"{len(jobs)} city requests ({sum(j[0] == START for j in jobs)} full history) to {end}", flush=True)
    new, failed, skipped = {}, [], 0

    def run(job):
        start, cid, lat, lon = job
        if (time.time() - t0) / 60 > budget_min:
            return cid, None, "budget"
        try:
            return cid, fetch_nasa(lat, lon, start, end), None
        except Exception as exc:  # noqa: BLE001
            return cid, None, str(exc)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        for i, (cid, s, err) in enumerate(ex.map(run, jobs), 1):
            if s is not None and not s.empty:
                new[cid] = s.round(2)
            elif err == "budget":
                skipped += 1
            else:
                failed.append((cid, err))
            if i % 20 == 0:
                print(f"  {i}/{len(jobs)}", flush=True)
    out = pd.DataFrame(new).combine_first(saved) if new else saved.copy()   # new values win (revision window)
    out = out.sort_index()
    out.index.name = "date"
    out = out[[c for c in cities["city_id"] if c in out]]
    return out, failed, skipped


# ---------------------------------------------------------------- degree days
def hdd(t, d):
    v = (18.0 - t).clip(lower=0) if d == "18C" else (18.0 - t).where(t <= 15.0, 0.0)
    return v.where(t.notna())


def cdd(t, d):
    v = (t - 18.0).clip(lower=0) if d == "18C" else (t - 21.0).where(t >= 24.0, 0.0)
    return v.where(t.notna())


def country_degree_days(temps, cities, countries, aggregates):
    """{(kind, def): DataFrame(date x country/aggregate key)} population-weighted."""
    out = {(k, d): {} for k in ("HDD", "CDD") for d in DEFS}
    for cc, g in cities.groupby("country_code", sort=False):
        names = [n for n in g["city_id"]]
        if not all(n in temps for n in names):
            continue
        t = temps[names]
        ok = t.notna().all(axis=1)
        w = g.set_index("city_id")["weight"].reindex(names)
        for d in DEFS:
            for kind, f in (("HDD", hdd), ("CDD", cdd)):
                out[(kind, d)][cc] = f(t, d).mul(w, axis=1).sum(axis=1).where(ok)
    res = {k: pd.DataFrame(v) for k, v in out.items()}
    pop = countries.set_index("country_code")["population"]
    for key in aggregates:
        members = [c for c in countries["country_code"]
                   if key in str(countries.set_index("country_code").loc[c, "aggregates"]).split(";") and pd.notna(pop[c])]
        for k, df in res.items():
            if not all(m in df for m in members):
                continue
            ok = df[members].notna().all(axis=1)
            wm = pop[members] / pop[members].sum()
            df[key] = df[members].mul(wm, axis=1).sum(axis=1).where(ok)
    return res


def monthly_sums(daily):
    n = daily.notna().resample("MS").sum()
    s = daily.resample("MS").sum(min_count=1)
    dim = pd.Series([calendar.monthrange(d.year, d.month)[1] for d in s.index], index=s.index)
    out = s.where(n.eq(dim, axis=0))
    out.index.name = "month"
    return out.dropna(how="all"), n


def doy_normals(daily, years=NORMAL_YEARS):
    d = daily[(daily.index.year >= years[0]) & (daily.index.year <= years[1])]
    g = d.groupby([d.index.month, d.index.day]).mean()
    g.index = [pd.Timestamp(2000, m, dd).dayofyear for m, dd in g.index]   # 2000 is a leap year: 366 days
    g.index.name = "day of year (leap-year numbering)"
    return g


def month_normals(monthly, years=NORMAL_YEARS):
    m = monthly[(monthly.index.year >= years[0]) & (monthly.index.year <= years[1])]
    g = m.groupby(m.index.month).mean()
    g.index.name = "month"
    return g


def season_normal_cumulative(doy_norm_series, start_month):
    """Cumulative normal from the season start, indexed by day of season 1-365 (29 Feb dropped)."""
    days = pd.date_range("2001-%02d-01" % start_month, periods=365, freq="D")
    vals = [doy_norm_series.get(pd.Timestamp(2000, d.month, d.day).dayofyear, np.nan) for d in days]
    return pd.Series(np.cumsum(vals), index=range(1, 366))


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", required=True)
    ap.add_argument("--out-dir", default=DATA_DIR)
    ap.add_argument("--budget-min", type=float, default=100.0)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--no-fetch", action="store_true", help="rebuild the workbook from the store only")
    args = ap.parse_args()
    region = args.region
    cfg, cities, countries = load_region(region)
    sdir = store_dir(args.out_dir, region)
    saved = load_store(sdir)
    print(f"{region}: {len(cities)} cities, {cities['country_code'].nunique()} countries; store {saved.shape}", flush=True)
    failed, skipped = [], 0
    if args.no_fetch:
        temps = saved
    else:
        temps, failed, skipped = update_store(cities, saved, args.budget_min, args.workers)
        n = save_store(temps, sdir, saved)
        print(f"store: {n} yearly files written", flush=True)
    if temps.empty:
        raise SystemExit("no temperatures")
    aggregates = cfg.get("aggregates", {})
    dd = country_degree_days(temps, cities, countries, aggregates)
    keys = list(dd[("HDD", "18C")].columns)
    names = dict(zip(countries["country_code"], countries["country"]))
    names.update(aggregates)

    daily_sheets, monthly, counts = {}, {}, None
    for kind in ("HDD", "CDD"):
        d = pd.DataFrame({f"{k} {df_}": dd[(kind, df_)][k] for k in keys for df_ in DEFS})
        d.index.name = "date"
        daily_sheets[f"Daily {kind}"] = d[d.index >= DAILY_FROM].round(2)
    mon_cols, norm_m_cols, norm_d_cols = {}, {}, {}
    for (kind, d_), df in dd.items():
        m, n = monthly_sums(df)
        for k in df.columns:
            mon_cols[f"{k} {kind} {d_}"] = m[k] if k in m else pd.Series(dtype=float)
        norm_m = month_normals(m)
        norm_d = doy_normals(df)
        for k in df.columns:
            norm_m_cols[f"{k} {kind} {d_}"] = norm_m[k]
            norm_d_cols[f"{k} {kind} {d_}"] = norm_d[k]
        if kind == "HDD" and d_ == "18C":
            counts = n
    monthly_df = pd.DataFrame(mon_cols).round(1)
    norm_m_df = pd.DataFrame(norm_m_cols).round(2)
    norm_d_df = pd.DataFrame(norm_d_cols).round(3)
    anom, anom_cols = {}, {}
    for c in monthly_df.columns:
        nm = monthly_df.index.month.map(norm_m_df[c])
        a = monthly_df[c] - np.asarray(nm, float)
        anom[f"{c} abs"] = a
        anom[f"{c} pct"] = (a / np.asarray(nm, float) * 100.0).where(np.asarray(nm, float) >= 1.0)
    anom_df = pd.DataFrame(anom).round(1)
    counts.columns = [f"{c} days" for c in counts.columns]
    last_day = temps.dropna(how="all").index.max()
    counts = counts[counts.index <= last_day.to_period("M").to_timestamp()]
    city_cov = pd.DataFrame({
        "first day": temps.apply(lambda s: s.first_valid_index()), "last day": temps.apply(lambda s: s.last_valid_index()),
        "days present": temps.notna().sum(),
        "days expected": [(last_day - s.first_valid_index()).days + 1 if s.first_valid_index() is not None else 0
                          for _, s in temps.items()]})
    city_cov["% present"] = (city_cov["days present"] / city_cov["days expected"].replace(0, np.nan) * 100).round(2)
    city_cov.index.name = "city_id"
    cit = cities.drop(columns=["city_id"]).copy()
    cit["weight"] = cit["weight"].round(5)
    cit = cit.set_index("country_code")
    cty = countries.copy()
    cty["cities"] = cty["country_code"].map(cities.groupby("country_code").size())
    cty = cty.set_index("country_code")
    if failed:
        print("FAILED cities:", failed, flush=True)

    sheets = {"Cities": cit, "Countries": cty, **daily_sheets, "Monthly": monthly_df, "Normals month": norm_m_df,
              "Normals doy": norm_d_df, "Anomaly": anom_df, "Coverage": counts, "City coverage": city_cov}
    notes = [
        f"{cfg['title']} degree days - population-weighted countries from NASA POWER city temperatures",
        "",
        "What this is",
        f"Heating (HDD) and cooling (CDD) degree days per country and day, from {cities.groupby('country_code').size().min()}-"
        f"{cities.groupby('country_code').size().max()} cities per country ({len(cities)} cities, {cities['country_code'].nunique()} countries; the 4 most populous cities by an official population source, see 'Cities'), weighted by city population.",
        "Temperature: NASA POWER daily point API, parameter T2M = daily mean 2 m air temperature (MERRA-2 reanalysis, 0.5 x 0.625 degree cell at the city's lat/lon), from 1991-01-01 to "
        f"the latest day NASA has ({last_day:%d %b %Y}). A reanalysis value for a grid cell, not a station reading. T2M_MAX / T2M_MIN are not stored.",
        "",
        "Definitions (both on every sheet, as column suffixes)",
        "18C: repo convention, base 18 C: HDD = max(0, 18 - Tmean); CDD = max(0, Tmean - 18).",
        "EUROSTAT: HDD = 18 - Tm when Tm <= 15 C, else 0; CDD = Tm - 21 when Tm >= 24 C, else 0 (Eurostat's definition of degree days).",
        "Unit: degree days (C x day) per day; monthly sheets are sums over the month.",
        "",
        "Weighting and gaps",
        "City weight = population share of the country's cities (Cities sheet). If a population could not be sourced the country's cities get equal weight, flagged ASSUMPTION in 'weight_basis'.",
        "A day is used only if EVERY city of the country has a temperature; nothing is filled. A month is summed only if every day is present ('Coverage' = days present per month).",
        "Aggregates (" + "; ".join(f"{k} = {v}" for k, v in aggregates.items()) + "): member countries weighted by official population (Countries sheet), on days when every member has a value.",
        "",
        "Sheets",
        f"Daily HDD / Daily CDD: from {DAILY_FROM} (the store keeps 1991 onward). Monthly: complete months from 1991. Normals month / Normals doy: mean over {NORMAL_YEARS[0]}-{NORMAL_YEARS[1]} (doy numbering of a leap year).",
        "Anomaly: monthly actual minus the 1991-2020 normal of that calendar month ('abs') and as % of the normal ('pct', blank where the normal is under 1 degree day).",
        "",
        "Incremental store",
        f"weather_stores/{region}/city_t2m_<year>.csv.gz hold the daily city temperatures; each run refetches every city from its last stored day minus {REVISION_DAYS} days.",
        f"This run: {len(failed)} city requests failed, {skipped} skipped by the time budget." + (" Failed: " + ", ".join(c for c, _ in failed[:20]) if failed else ""),
        "",
        "Sources",
        cfg.get("source_line", ""),
        "NASA POWER: https://power.larc.nasa.gov/",
    ]
    out = os.path.join(args.out_dir, f"{region}_degree_days.xlsx")
    os.makedirs(args.out_dir, exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes, {"What this is", "Definitions (both on every sheet, as column suffixes)",
                                                    "Weighting and gaps", "Sheets", "Incremental store", "Sources"})
    print(f"Saved {out}; last day {last_day:%Y-%m-%d}", flush=True)
    if failed and len(failed) > len(cities) * 0.25:
        raise SystemExit("too many failed cities")


if __name__ == "__main__":
    main()
