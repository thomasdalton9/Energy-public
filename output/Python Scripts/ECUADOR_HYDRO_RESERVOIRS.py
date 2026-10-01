"""
Ecuador hydro reservoirs: daily lake level (m above sea level) of the Paute
cascade reservoirs - Mazar (the big seasonal store whose drawdown drove
the 2023-24 blackouts) and Amaluza (Paute-Molino's reservoir) - from CELEC
SUR's own SCADA, no key needed.

Source: the public Oracle ORDS service behind CELEC EP Unidad de Negocio
CELEC SUR's dashboard "Graficas de Produccion"
(https://generacioncsr.celec.gob.ec/graficasproduccion/):
    GET https://generacioncsr.celec.gob.ec:8443/ords/csr/sardomcsr/pointValuesMesH24
        ?mrid=<point>&fechaInicio=YYYY-MM-01T00:00:00.000Z&fechaFin=<next month, day 2>T00:00:00.000Z
        &fecha=01/MM/YYYY 00:00:00
returns {"items": [{"loctimestamp": "...Z", "valueedit": <m>}, ...]} - one
reading per day, at local midnight (05:00 UTC). (With fechaFin on the 1st
of the next month the service leaves the month's last day empty, so the
request runs one day further.) Measuring points (mrid): Mazar level 30031,
Amaluza level 24019. Endpoint and points were first documented by the
citizen project github.com/jordanvt18/cotas-embalses-ecuador and checked
in discovery_archive/south_america/HYDRO_PE_EC_UY_DISCOVERY.py (rounds 1-2).
The service has a normal public certificate chain (Sectigo): standard TLS
verification.

Mazar % = the level's position in Mazar's operating range, 100 x (level -
2098) / (2153 - 2098), with 2098 m the minimum operating level and 2153 m
the maximum normal level CELEC uses. It is a LEVEL-based gauge, not a
share of stored volume (CELEC publishes no level-volume curve): Mazar is
a V-shaped canyon reservoir, so a given % of range holds less than that %
of the useful volume near the bottom.

History: the service holds Mazar back to 2010 (the plant's first year),
so the workbook starts 2011-01-01 and covers the 2023-24 drawdowns.
CENACE's own daily page shows no levels and keeps no archive.

Incremental: the workbook is the archive. Each run re-reads the current
and previous month (readings can be filled in late) and any month with a
missing day.

Usage: python3 ECUADOR_HYDRO_RESERVOIRS.py [--out PATH] [--start YYYY-MM-DD]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/ecuador_hydro_reservoirs.xlsx"
HOST = "generacioncsr.celec.gob.ec"
API = f"https://{HOST}:8443/ords/csr/sardomcsr/pointValuesMesH24"
DASHBOARD = f"https://{HOST}/graficasproduccion/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
           "Accept": "application/json", "Referer": DASHBOARD}
START = dt.date(2011, 1, 1)
# column stem -> (mrid, name, min operating level m, max normal level m, operator / plant)
RESERVOIRS = {
    "Mazar": (30031, "Mazar", 2098.0, 2153.0, "CELEC SUR - Mazar plant (170 MW); regulates the Paute cascade"),
    "Amaluza": (24019, "Amaluza", 1975.0, 1991.0, "CELEC SUR - Paute-Molino (1,100 MW); daily-regulation reservoir"),
}
TZ = "America/Guayaquil"


def celec_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def month_starts(a, b):
    m = dt.date(a.year, a.month, 1)
    while m <= b:
        yield m
        m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)


def fetch_month(session, mrid, month):
    nxt = dt.date(month.year + (month.month == 12), month.month % 12 + 1, 1)
    end = nxt + dt.timedelta(days=1)   # one day further, or the month's last day comes back empty
    params = {"mrid": mrid, "fechaInicio": f"{month}T00:00:00.000Z", "fechaFin": f"{end}T00:00:00.000Z",
              "fecha": f"{month:%d/%m/%Y} 00:00:00"}
    for attempt in range(4):
        try:
            r = session.get(API, params=params, timeout=(15, 90))
            r.raise_for_status()
            items = r.json().get("items") or []
            break
        except (requests.RequestException, ValueError) as e:
            if attempt == 3:
                raise
            print(f"    retry {mrid} {month:%Y-%m}: {type(e).__name__}: {str(e)[:150]}", flush=True)
            time.sleep(5 * (attempt + 1))
    vals = {}
    for it in items:
        v = it.get("valueedit")
        if v is None or not it.get("loctimestamp"):
            continue
        day = pd.Timestamp(it["loctimestamp"]).tz_convert(TZ).tz_localize(None).normalize()
        vals[day] = float(v)
    s = pd.Series(list(vals.values()), index=pd.DatetimeIndex(list(vals.keys())), dtype=float).sort_index()
    return s[(s.index >= pd.Timestamp(month)) & (s.index < pd.Timestamp(nxt))]


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    df.index.name = "date"
    return df


def finish(df):
    cols = [f"{k}Level_m" for k in RESERVOIRS]
    df = df.reindex(columns=cols + ["Mazar_pct_of_range"])
    lo, hi = RESERVOIRS["Mazar"][2], RESERVOIRS["Mazar"][3]
    df["Mazar_pct_of_range"] = (100 * (df["MazarLevel_m"] - lo) / (hi - lo)).round(1)
    df[cols] = df[cols].round(2)
    df.index.name = "date"
    return df.sort_index().dropna(how="all")


def notes(df):
    first, last = df.index.min(), df.index.max()
    latest = df.dropna(subset=["MazarLevel_m"]).iloc[-1]
    return [
        "UNITS",
        "MazarLevel_m, AmaluzaLevel_m: reservoir water level (cota), metres above sea level, at local midnight "
        "(CELEC SUR's daily reading, 00:00 Ecuador time = 05:00 UTC).",
        "Mazar_pct_of_range: 100 x (Mazar level - 2098 m) / (2153 m - 2098 m) - the level's position between "
        "Mazar's minimum operating level (2098 m) and maximum normal level (2153 m). LEVEL-based, not a share of "
        "stored volume (no level-volume curve is published); can go slightly above 100 when the lake is over "
        "its normal maximum and below 0 under the minimum. CELEC also watches a 'critical' level of 2115 m "
        f"(= {100 * (2115 - 2098) / 55:.0f}% of range).",
        "",
        "COVERAGE",
        f"{first:%d-%b-%Y} to {last:%d-%b-%Y} ({len(df):,} days). Latest Mazar {latest['MazarLevel_m']:.2f} m "
        f"({latest['Mazar_pct_of_range']:.0f}% of range) on {latest.name:%d-%b-%Y}.",
        "CELEC SUR's service goes back to 2010 (Mazar's first year of operation); this workbook starts 2011-01-01. "
        "CENACE's Informacion Operativa page shows no reservoir levels and keeps no archive.",
        "Not covered: Daule-Peripa (CELEC Hidronacion) and Pisayambo (CELEC Hidroagoyan) - those units publish "
        "no machine-readable level series.",
        "",
        "SOURCE",
        f"CELEC EP - CELEC SUR, 'Graficas de Produccion' dashboard ({DASHBOARD}); data from its ORDS service "
        f"{API} (mrid 30031 Mazar, 24019 Amaluza), one request per reservoir per month.",
        "Script: south_america/ECUADOR_HYDRO_RESERVOIRS.py (scheduled by .github/workflows/ecuador_hydro_reservoirs.yml).",
        "",
        "METHOD",
        "Incremental: the first run reads every month since 2011; later runs re-read the current and previous "
        "month plus any month of the past year with a missing day; earlier months are kept from the workbook.",
        "Daily value = CELEC SUR's own daily (H24) reading at local midnight - not a daily mean.",
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=START)
    args = ap.parse_args()

    arch = load_archive(args.out)
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=-5))).date()
    months = list(month_starts(args.start, today))
    session = celec_session()
    new = {}
    failed = 0
    for k, (mrid, *_rest) in RESERVOIRS.items():
        col = f"{k}Level_m"
        saved = arch[col].dropna() if col in arch else pd.Series(dtype=float)
        todo = []
        for i, m in enumerate(months):
            days = pd.date_range(m, min(pd.Timestamp(m) + pd.offsets.MonthEnd(0), pd.Timestamp(today)), freq="D")
            # re-read: everything on the first run; then the last two months always, and a month of the past
            # year with a missing day. Older gaps are gaps at CELEC SUR and are not asked for again every day.
            if saved.empty or i >= len(months) - 2 or (i >= len(months) - 12 and not days.isin(saved.index).all()):
                todo.append(m)
        print(f"{k}: {len(saved):,} days saved; fetching {len(todo)} month(s) "
              f"{todo[0]:%Y-%m}..{todo[-1]:%Y-%m}", flush=True)
        parts = []
        for m in todo:
            try:
                parts.append(fetch_month(session, mrid, m))
            except Exception as e:  # noqa: BLE001 - one month down shouldn't lose the rest; retried next run
                failed += 1
                print(f"  {k} {m:%Y-%m}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
            time.sleep(0.3)
        s = pd.concat(parts) if parts else pd.Series(dtype=float)
        s = s[~s.index.duplicated(keep="last")]
        new[f"{k}Level_m"] = s
        print(f"  {k} (mrid {mrid}): {len(s)} days" + (f", {s.index.min():%d-%b-%Y}..{s.index.max():%d-%b-%Y}, "
                                                     f"latest {s.iloc[-1]:.2f} m" if len(s) else ""), flush=True)
    new = pd.DataFrame(new)
    df = new.combine_first(arch) if not arch.empty else new
    if df.empty:
        sys.exit("No data fetched and no archive - nothing to write")
    df = finish(df[df.index >= pd.Timestamp(args.start)])
    res = pd.DataFrame([{"column": f"{k}Level_m", "reservoir": v[1], "CELEC SUR mrid": v[0],
                         "min operating level m": v[2], "max normal level m": v[3], "plant / role": v[4]}
                        for k, v in RESERVOIRS.items()]).set_index("column")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    nl = notes(df)
    xlsx_notes.write_workbook(args.out, {"Daily": df, "Reservoirs": res}, nl, {ln for ln in nl if ln and ln.isupper()})
    print(f"Saved {args.out}: {len(df):,} days {df.index.min():%d-%b-%Y} .. {df.index.max():%d-%b-%Y}", flush=True)
    print(df.tail(7).to_string(), flush=True)
    if failed and failed > max(4, len(todo)):
        sys.exit(f"{failed} month requests failed")


if __name__ == "__main__":
    main()
