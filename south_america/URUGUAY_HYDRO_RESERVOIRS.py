"""
Uruguay hydro reservoirs: daily lake level (m above sea level) of the Rio
Negro reservoirs - Rincon del Bonete (Gabriel Terra, Uruguay's one big
storage lake), Baygorria and Palmar - and of Salto Grande on the Uruguay
river (binational with Argentina). No key needed.

Sources
  1. ADME (Administracion del Mercado Electrico), per-plant SCADA series
     behind its 'Operacion Rio Negro' page (pronos.adme.com.uy/seriesbonete.php):
         GET https://pronos.adme.com.uy/cgi-bin/seriescentralhidro.cgi
             ?idCentral=bon|bay|pal&ts=ods&dtIni=<Excel serial day>&dtFin=<serial>
     an .ods with hourly Pot_MW_, hToma_m_ (level upstream of the dam = the
     lake level), hDescarga_m_, flows. A day's level = the mean of that
     day's hourly hToma readings (Uruguay local time, as ADME stamps them);
     ADME's missing-value codes (-121111, 0) are dropped. Valid data from
     2017 (2016 and earlier are missing codes). ADME times out on long
     ranges, so one request = one month per plant.
  2. INA (Instituto Nacional del Agua, Argentina) a5 database, series 26319
     'Salto Grande Arriba' - the Prefectura Naval gauge on the lake side of
     Salto Grande dam, daily mean level (m):
         https://alerta.ina.gob.ar/a5/obs/puntual/series/26319/observaciones
     ADME publishes only Salto Grande's flows and CTM Salto Grande's own
     'Reporte de Caudales y Niveles' is a one-day PDF with no archive, so
     INA is the daily history for the lake level.
Found via discovery_archive/south_america/HYDRO_PE_EC_UY_DISCOVERY.py (rounds 1-2).

Bonete_pct_of_range = 100 x (Bonete level - 70 m) / (80 m - 70 m): the
level's position between the lake's minimum operating level (70 m) and its
maximum operating level (80 m; ADME's seasonal programming uses this band
and a 2019 court order holds UTE to 80 m). LEVEL-based, not a share of
stored volume or energy - ADME and UTE publish no stored volume / energy
series for the lake. Can exceed 100 in floods (spilling above 80 m).

Incremental: the workbook is the archive. ADME is re-read from the last
saved day minus REFRESH_DAYS; months with no saved reading are filled
newest-first within --budget-min; INA is re-read from 30 days before its
last saved day.

Usage: python3 URUGUAY_HYDRO_RESERVOIRS.py [--out PATH] [--start YYYY-MM-DD] [--budget-min N]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys
import time
import zipfile

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))  # repo root, for xlsx_notes

import URUGUAY_ADME as U  # noqa: E402  (ADME .ods reader and the per-plant series URL)
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/uruguay_hydro_reservoirs.xlsx"
START = dt.date(2017, 1, 1)
REFRESH_DAYS = 5
INA_OBS = "https://alerta.ina.gob.ar/a5/obs/puntual/series/{sid}/observaciones"
INA_SALTO_GRANDE = 26319
# ADME id -> (column stem, name, plant MW, notes)
PLANTS = {
    "bon": ("Bonete", "Rincon del Bonete (Gabriel Terra)", 152, "Rio Negro; the storage lake (~8,800 hm3 at 80 m)"),
    "bay": ("Baygorria", "Baygorria", 108, "Rio Negro, run-of-river below Bonete"),
    "pal": ("Palmar", "Palmar (Constitucion)", 333, "Rio Negro, lowest dam; small seasonal storage"),
}
BONETE_MIN, BONETE_MAX = 70.0, 80.0
TZ = "America/Montevideo"
COLS = ["BoneteLevel_m", "SaltoGrandeLevel_m", "PalmarLevel_m", "BaygorriaLevel_m", "Bonete_pct_of_range"]


def serial(d):
    return (d - dt.date(1899, 12, 30)).days


def month_starts(a, b):
    m = dt.date(a.year, a.month, 1)
    while m <= b:
        yield m
        m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)


def adme_month(session, plant, a, b):
    """Daily mean hToma (m) for one plant, days a..b."""
    params = {"idCentral": plant, "ts": "ods", "dtIni": serial(a), "dtFin": serial(b + dt.timedelta(days=1))}
    for attempt in range(3):
        try:
            r = session.get(U.PLANT_SERIES_URL, params=params, timeout=(15, 180))
            r.raise_for_status()
            rows = next(iter(U.read_ods(r.content).values()))
            break
        except (requests.RequestException, zipfile.BadZipFile, StopIteration) as e:
            if attempt == 2:
                raise
            print(f"    retry {plant} {a}: {type(e).__name__}", flush=True)
            time.sleep(5 * (attempt + 1))
    head = next(i for i, row in enumerate(rows) if len(row) > 2 and str(row[1]).startswith("Pot_MW_"))
    j = rows[head].index("hToma_m_")
    vals = {}
    for row in rows[head + 1:]:
        if not row or not row[0] or len(row) <= j:
            continue
        t = pd.to_datetime(row[0], errors="coerce")
        v = pd.to_numeric(row[j], errors="coerce")
        if pd.isna(t) or pd.isna(v) or v <= 1:   # ADME's missing codes: -121111 and 0
            continue
        vals[t] = float(v)
    s = pd.Series(vals, dtype=float)
    if s.empty:
        return s
    s.index = pd.DatetimeIndex(s.index)
    daily = s.groupby(s.index.normalize()).mean()
    return daily[(daily.index >= pd.Timestamp(a)) & (daily.index <= pd.Timestamp(b))]


def ina_series(session, sid, since, until):
    r = session.get(INA_OBS.format(sid=sid), params={"timestart": f"{since:%Y-%m-%d}",
                                                      "timeend": f"{until + dt.timedelta(days=1):%Y-%m-%d}",
                                                      "format": "json"}, timeout=300)
    r.raise_for_status()
    obs = r.json()
    s = pd.Series({pd.Timestamp(o["timestart"]).tz_convert("America/Argentina/Buenos_Aires").tz_localize(None)
                   .normalize(): pd.to_numeric(o["valor"], errors="coerce") for o in obs}, dtype=float)
    s = s.dropna().sort_index()
    return s[(s > 20) & (s < 45)]   # the lake runs ~30-36 m; drop gauge glitches


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame(columns=COLS)
    df.index = pd.to_datetime(df.index)
    df.index.name = "date"
    return df


def finish(df):
    df = df.reindex(columns=COLS).sort_index()
    df["Bonete_pct_of_range"] = (100 * (df["BoneteLevel_m"] - BONETE_MIN) / (BONETE_MAX - BONETE_MIN)).round(1)
    lv = [c for c in COLS if c.endswith("_m")]
    df[lv] = df[lv].round(3)
    df.index.name = "date"
    return df.dropna(how="all")


def notes(df):
    first, last = df.index.min(), df.index.max()
    b = df["BoneteLevel_m"].dropna()
    sg = df["SaltoGrandeLevel_m"].dropna()
    lines = [
        "UNITS",
        "BoneteLevel_m, BaygorriaLevel_m, PalmarLevel_m: lake level upstream of each Rio Negro dam (ADME 'hToma'), "
        "metres above sea level, daily mean of ADME's hourly SCADA readings.",
        "SaltoGrandeLevel_m: Salto Grande lake level at the dam ('Salto Grande Arriba' gauge, Prefectura Naval "
        "Argentina, via INA), daily mean, metres.",
        f"Bonete_pct_of_range: 100 x (Bonete level - {BONETE_MIN:.0f} m) / ({BONETE_MAX:.0f} m - {BONETE_MIN:.0f} m) - "
        "position of the level between the lake's minimum (70 m) and maximum (80 m) operating levels. LEVEL-based: "
        "not a share of stored volume or energy (ADME and UTE publish neither for the lake). Above 100 = spilling "
        "above 80 m in a flood.",
        "",
        "COVERAGE",
        f"{first:%d-%b-%Y} to {last:%d-%b-%Y} ({len(df):,} days). Latest Bonete {b.iloc[-1]:.2f} m "
        f"({df.loc[b.index[-1], 'Bonete_pct_of_range']:.0f}% of range) on {b.index[-1]:%d-%b-%Y}"
        + (f"; Salto Grande {sg.iloc[-1]:.2f} m on {sg.index[-1]:%d-%b-%Y}." if len(sg) else "."),
        "ADME's Rio Negro level series are valid from 2017 (earlier years hold only missing-value codes); "
        "no public daily stored-energy series exists for Uruguay.",
        "",
        "SOURCE",
        "ADME - https://pronos.adme.com.uy/seriesbonete.php ('Operacion Rio Negro'): "
        "https://pronos.adme.com.uy/cgi-bin/seriescentralhidro.cgi?idCentral=bon|bay|pal&ts=ods (hourly .ods).",
        f"INA a5 - {INA_OBS.format(sid=INA_SALTO_GRANDE)} (series 26319, Salto Grande Arriba).",
        "Script: south_america/URUGUAY_HYDRO_RESERVOIRS.py (scheduled by .github/workflows/uruguay_hydro_reservoirs.yml).",
        "",
        "METHOD",
        "Incremental: ADME is re-read from the last saved day minus 5 days (months never saved are filled "
        "newest-first within the run's time budget); INA from 30 days before its last saved day.",
    ]
    return lines


def save(path, df):
    df = finish(df)
    plants = pd.DataFrame([{"column": f"{v[0]}Level_m", "lake / dam": v[1], "ADME idCentral": k, "plant MW": v[2],
                            "notes": v[3]} for k, v in PLANTS.items()]
                          + [{"column": "SaltoGrandeLevel_m", "lake / dam": "Salto Grande (binational)",
                              "ADME idCentral": None, "plant MW": 1890,
                              "notes": "INA a5 series 26319; Uruguay's share is half the plant"}]).set_index("column")
    nl = notes(df)
    xlsx_notes.write_workbook(path, {"Daily": df, "Reservoirs": plants}, nl, {ln for ln in nl if ln and ln.isupper()})
    return df


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=START)
    ap.add_argument("--budget-min", type=float, default=30, help="stop the ADME backfill after this many minutes")
    args = ap.parse_args()
    t0 = time.time()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    df = load_archive(args.out)
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=-3))).date()
    session = requests.Session()
    session.headers.update(U.HEADERS)

    # INA first (one request)
    col = "SaltoGrandeLevel_m"
    saved = df[col].dropna() if col in df else pd.Series(dtype=float)
    since = (saved.index.max() - pd.Timedelta(days=30)).date() if len(saved) else args.start
    try:
        s = ina_series(session, INA_SALTO_GRANDE, since, today)
        print(f"INA Salto Grande {since}..: {len(s)} days, latest {s.iloc[-1]:.2f} m on {s.index[-1]:%d-%b-%Y}"
              if len(s) else "INA Salto Grande: no data", flush=True)
        df = pd.DataFrame({col: s}).combine_first(df)
    except Exception as e:  # noqa: BLE001 - keep the saved days; retried next run
        print(f"INA FAILED {type(e).__name__}: {e}", flush=True)

    # ADME: recent window, then never-saved months newest first
    jobs = []
    for plant, (stem, *_r) in PLANTS.items():
        col = f"{stem}Level_m"
        saved = df[col].dropna() if col in df else pd.Series(dtype=float)
        if len(saved):
            a = max(args.start, (saved.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date())
            jobs += [(0, a, plant, a, today)]
        have = {(d.year, d.month) for d in saved.index}
        recent_from = (saved.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date().replace(day=1) if len(saved) \
            else None
        for m in month_starts(args.start, today):
            end = min(dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1) - dt.timedelta(days=1), today)
            if (m.year, m.month) not in have and (recent_from is None or m < recent_from):
                jobs.append((1, m, plant, m, end))
    jobs.sort(key=lambda j: (j[0], -j[1].toordinal()))
    print(f"ADME: {len(jobs)} request(s) to make", flush=True)
    done = 0
    for _, _, plant, a, b in jobs:
        if time.time() - t0 > args.budget_min * 60:
            print(f"time budget reached; {len(jobs) - done} ADME request(s) left for later runs", flush=True)
            break
        col = f"{PLANTS[plant][0]}Level_m"
        try:
            s = adme_month(session, plant, a, b) if (b - a).days <= 40 else pd.concat(
                [adme_month(session, plant, max(m, a), min(pd.Timestamp(m) + pd.offsets.MonthEnd(0),
                                                          pd.Timestamp(b)).date())
                 for m in month_starts(a, b)])
        except Exception as e:  # noqa: BLE001
            print(f"  {plant} {a}..{b}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
            s = pd.Series(dtype=float)
        if len(s):
            df = pd.DataFrame({col: s}).combine_first(df)
        done += 1
        if done % 10 == 0:
            print(f"  {done}/{len(jobs)} ADME requests; {plant} {a}..{b}: {len(s)} days", flush=True)
            save(args.out, df)
        time.sleep(0.5)
    if df.dropna(how="all").empty:
        sys.exit("No data fetched and no archive - nothing to write")
    df = save(args.out, df[df.index >= pd.Timestamp(args.start)])
    print(f"Saved {args.out}: {len(df):,} days {df.index.min():%d-%b-%Y} .. {df.index.max():%d-%b-%Y}", flush=True)
    print(df.tail(7).to_string(), flush=True)
    print("Days per column:", df.notna().sum().to_dict(), flush=True)


if __name__ == "__main__":
    main()
