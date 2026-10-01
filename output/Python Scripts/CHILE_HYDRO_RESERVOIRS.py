"""
Chile hydro reservoirs: stored volume (million m3) of the main hydropower
reservoirs, daily, from DGA (Direccion General de Aguas, Ministry of Public
Works). No API key needed.

Sources
  1. Daily, from 30-Jun-2024: the JSON API behind DGA's Visualizador
     Hidrometrico Nacional (https://vipnet.mop.gob.cl), which serves the
     half-hourly reservoir volume ('instantaneo', Mm3) of each DGA lake/
     reservoir station:  POST https://vipnet.mop.gob.cl/v1/vipnet/estacion/valores
     {codigoEstacion, tipoEstacion: 2 (embalse), fetchDay, fetchHour, hoursRange}.
     Its history starts 2024-06-30 (POST /v1/vipnet/mediciones/fecha-inicio).
     A day's value = the median of that day's readings (Chile local time).
  2. Month-end, Oct-2019 to Jun-2024: DGA's monthly 'Boletin Hidrometrico'
     (dga.mop.gob.cl uploads). Each September issue has a 'Resumen Anual'
     table with the end-of-month volume of every reservoir for the whole
     Oct-Sep water year: the 2020, 2021, 2022 and 2024 issues cover water
     years 2019/20-2021/22 and 2023/24 (the 2023 issue's tables are images
     only, so 2022/23 is a gap).
     These rows sit on the last day of each month.
Found via discovery_archive/south_america/CHILE_POWER_HYDRO_PROBE3.py / _PROBE4.py.
(CEN's own reservoir data needs the SIP API key, and coordinador.cl is behind
a Cloudflare browser check, so neither is used.)

Incremental: the workbook is the archive. Each run re-reads only the daily
API from the last saved day (minus REVISION_DAYS); the bulletin history is
read only while any of its month-ends is missing.

Usage: python3 CHILE_HYDRO_RESERVOIRS.py [--out PATH]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/chile_hydro_reservoirs.xlsx"
API = "https://vipnet.mop.gob.cl/v1/vipnet"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
           "Accept": "application/json", "Origin": "https://vipnet.mop.gob.cl", "Referer": "https://vipnet.mop.gob.cl/"}
TZ = "America/Santiago"
API_START = dt.date(2024, 6, 30)
HISTORY_START = dt.date(2019, 10, 1)   # five water years before the API starts, for the 5-year range
REVISION_DAYS = 3
# Bump when the daily method changes (stations, fallbacks, daily statistic): the next run then re-reads the
# whole daily API history once. Stored in the workbook's 'Method' sheet.
DAILY_METHOD = "2"
WINDOW_HOURS = 8760                    # one API request = up to a year of half-hourly readings

# column -> (DGA station code on vipnet, name in the bulletin 'Resumen Anual' table, capacity Mm3, purpose)
# Capacities: DGA weekly Boletin Hidrometeorologico (Sep-2026).
RESERVOIRS = {
    "Colbun":       ("07321006-2", "Colbún", 1544, "Generación"),
    "Rapel":        ("06054001-2", "Rapel", 695, "Generación"),
    "Ralco":        ("08312002-9", "Ralco", 1174, "Generación"),
    "Pangue":       ("08313002-4", "Pangue", 83, "Generación"),
    "LagoLaja":     ("08370007-6", "Lago Laja", 5582, "Generación y riego (feeds El Toro, Antuco, Abanico)"),
    "LagunaMaule":  ("07300000-9", "Lag. Maule", 1359, "Generación y riego (feeds Cipreses/Isla, Pehuenche)"),
    "Melado":       ("07317004-4", None, None, "Generación (Pehuenche)"),
    "Invernada":    ("07306000-1", None, None, "Generación (Cipreses)"),
}
# backup station used on days the main one has no reading (same lake, gauge at the dam)
FALLBACK = {"LagunaMaule": "07300006-8"}
TOTAL_GAP_DAYS = 10                    # the daily total bridges a reservoir's gaps up to this long
TOTAL = [k for k, v in RESERVOIRS.items() if v[2]]          # in both sources, with a DGA capacity
CAPACITY = sum(RESERVOIRS[k][2] for k in TOTAL)

# September bulletins (water year Oct(Y-1)-Sep(Y)) whose 'Resumen Anual' table gives the month-end history
BULLETINS = {
    2020: "https://dga.mop.gob.cl/uploads/sites/13/2023/07/Boletin_septiembre_2020.pdf",
    2021: "https://dga.mop.gob.cl/uploads/sites/13/2023/07/Boletin_septiembre_2021.pdf",
    2022: "https://dga.mop.gob.cl/uploads/sites/13/2023/07/Informe-Boletin-DGA-septiembre-2022.pdf",
    # 2023 (Boletin-DGA-septiembre-2023.pdf): its reservoir tables are images with no text layer, so water year
    # 2022/23 has no month-end history here.
    2024: "https://dga.mop.gob.cl/uploads/sites/13/2024/06/Boletin-Hidrometrico-DGA-Septiembre-2024.pdf",
}
MONTHS = ["O", "N", "D", "E", "F", "M", "A", "M", "J", "J", "A", "S"]   # Resumen Anual columns, Oct..Sep

COLS = [f"{k}_Mm3" for k in RESERVOIRS] + ["Total_Mm3", "Total_pct", "Source"]

NOTES = [
    "UNITS",
    "Mm3 = million cubic metres of water stored (volume above the reservoir's minimum, as DGA reports it). "
    f"Total_pct = Total_Mm3 / {CAPACITY:,} Mm3, the summed capacity of the six reservoirs in the total.",
    "",
    "COVERAGE",
    "Daily from 30-Jun-2024 (DGA online network via the Visualizador Hidrométrico Nacional API; one value per "
    "day = median of the day's half-hourly readings, Chile time). Oct-2019 to Jun-2024: one row per month on the "
    "month's last day (end-of-month volume from DGA's monthly Boletín Hidrométrico, September issues, 'Resumen "
    "Anual' table); water year Oct-2022 to Sep-2023 is missing (that issue's tables are images only). Column "
    "'Source' says which. The water-year charts join the month-end points with straight "
    "lines so the 5-year range covers 2021/22 onwards.",
    "",
    "RESERVOIRS / METHOD",
    "Total_Mm3 = Colbún + Rapel + Ralco + Pangue + Lago Laja (Laguna de La Laja) + Laguna del Maule - the six "
    "reservoirs DGA classes as generation or generation-and-irrigation, the bulk of the SEN's storable hydro "
    "energy. Days where a reservoir has no reading bridge gaps of up to 10 days by straight-line interpolation "
    "(for the total only; the reservoir columns stay blank); longer gaps leave the total blank. Laguna del Maule "
    "falls back to DGA's dam gauge (07300006-8) on days its main station (07300000-9) has no reading.",
    "Capacities (Mm3, DGA weekly bulletin Sep-2026): Colbún 1,544; Rapel 695; Ralco 1,174; Pangue 83; Lago Laja "
    "5,582; Laguna del Maule 1,359 (DGA's current storage curve; earlier bulletins used 1,420).",
    "Melado and Laguna de la Invernada (Maule basin, generation) are daily only (not in the bulletin table) and "
    "are not in the total. Lago Chapo (Canutillar) has no volume on DGA's network.",
    "No energy-equivalent (GWh) is computed: CEN's energy coefficients per reservoir are not public without "
    "the API key.",
    "",
    "SOURCE",
    "DGA Visualizador Hidrométrico Nacional - https://vipnet.mop.gob.cl (API "
    "https://vipnet.mop.gob.cl/v1/vipnet/estacion/valores, station codes in the script). DGA Boletín "
    "Hidrométrico (monthly) - https://dga.mop.gob.cl/servicios-de-informacion/boletines/ . DGA online data is "
    "provisional and can be revised. Updated daily by GitHub Actions (chile_hydro_reservoirs.yml).",
]


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame(columns=COLS)
    df.index = pd.to_datetime(df.index)
    df.index.name = "date"
    return df


def station_series(session, code, start, end):
    """Half-hourly volume (Mm3) for one station between two dates -> daily median (Chile local days)."""
    vals = []
    day = end
    while day >= start - dt.timedelta(days=1):
        body = {"codigoEstacion": code, "tipoEstacion": 2, "fetchHour": 23, "fetchDay": day.isoformat(),
                "hoursRange": WINDOW_HOURS}
        for attempt in range(3):
            try:
                r = session.post(f"{API}/estacion/valores", json=body, timeout=(10, 180))
                r.raise_for_status()
                rows = r.json().get("data") or []
                break
            except Exception as e:  # noqa: BLE001
                if attempt == 2:
                    raise
                print(f"    retry {code} {day}: {e}", flush=True)
                time.sleep(5 * (attempt + 1))
        vals += [(x["fecha"]["$date"], x.get("instantaneo")) for x in rows]
        day -= dt.timedelta(hours=WINDOW_HOURS - 24)
    if not vals:
        return pd.Series(dtype=float)
    s = pd.Series([v for _, v in vals], index=pd.to_datetime([t for t, _ in vals], utc=True).tz_convert(TZ),
                  dtype=float)
    s = s[~s.index.duplicated()].dropna()
    s = s[s > 0]
    daily = s.groupby(s.index.tz_localize(None).normalize()).median()
    daily = daily[(daily.index >= pd.Timestamp(start)) & (daily.index <= pd.Timestamp(end))]
    return daily.round(2)


def fetch_daily(start, end):
    session = requests.Session()
    session.headers.update(HEADERS)
    out = {}
    for col, (code, *_rest) in RESERVOIRS.items():
        try:
            s = station_series(session, code, start, end)
            if col in FALLBACK:
                b = station_series(session, FALLBACK[col], start, end)
                print(f"  {col}: {int(b.index.difference(s.index).size)} days filled from {FALLBACK[col]}", flush=True)
                s = s.combine_first(b)
        except Exception as e:  # noqa: BLE001 - one station down shouldn't lose the others; retried next run
            print(f"  {col} ({code}): FAILED {type(e).__name__}: {e}", flush=True)
            s = pd.Series(dtype=float)
        print(f"  {col} ({code}): {len(s)} days" + (f", {s.index.min():%d-%b-%Y} to {s.index.max():%d-%b-%Y}, "
                                                  f"latest {s.iloc[-1]:,.1f} Mm3" if len(s) else ""), flush=True)
        out[f"{col}_Mm3"] = s
    df = pd.DataFrame(out)
    df["Source"] = "DGA online (daily median)"
    return df


def _num(tok):
    tok = tok.strip().replace(",", ".")
    return float(tok) if re.fullmatch(r"-?\d+(\.\d+)?", tok) else None


def parse_bulletin(content, wy_end):
    """Month-end volumes (Oct of wy_end-1 .. Sep of wy_end) from a September bulletin's 'Resumen Anual' table."""
    import pdfplumber
    names = {v[1]: k for k, v in RESERVOIRS.items() if v[1]}
    found = {}
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not re.search(r"Resumen\s+Anual", text, re.I):
                continue
            for line in text.splitlines():
                for name, col in names.items():
                    if col in found or not re.match(rf"^{re.escape(name)}\b", line.strip(), re.I):
                        continue
                    nums = [_num(t) for t in re.sub(r"\([^)]*\)", " ", line[len(name):]).split()]
                    nums = [n for n in nums if n is not None]
                    if len(nums) >= 12:
                        found[col] = nums[-12:]
            if len(found) == len(names):
                break
    if not found:
        print(f"  WY {wy_end - 1}/{wy_end}: 'Resumen Anual' table not found (no text layer?)", flush=True)
        return pd.DataFrame()
    idx = [pd.Timestamp(wy_end - 1 if m >= 10 else wy_end, m, 1) + pd.offsets.MonthEnd(0)
           for m in (10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9)]
    df = pd.DataFrame({f"{c}_Mm3": v for c, v in found.items()}, index=idx)
    print(f"  WY {wy_end - 1}/{wy_end}: {sorted(found)}", flush=True)
    return df


def fetch_bulletins(missing_years):
    frames = []
    for wy_end in sorted(missing_years):
        url = BULLETINS.get(wy_end)
        if not url:
            continue
        try:
            r = requests.get(url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=(10, 300))
            r.raise_for_status()
            frames.append(parse_bulletin(r.content, wy_end))
        except Exception as e:  # noqa: BLE001
            print(f"  WY {wy_end - 1}/{wy_end}: {url} failed: {e}", flush=True)
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames)
    df = df[df.index < pd.Timestamp(API_START)]
    df["Source"] = "DGA monthly bulletin (month-end)"
    return df


def finish(df):
    df = df.reindex(columns=COLS)
    tot = df[[f"{k}_Mm3" for k in TOTAL]].copy()
    daily = df["Source"].astype(str).str.startswith("DGA online")
    tot[daily] = tot[daily].interpolate(limit=TOTAL_GAP_DAYS, limit_area="inside")
    df["Total_Mm3"] = tot.sum(axis=1).where(tot.notna().all(axis=1)).round(1)
    df["Total_pct"] = (100 * df["Total_Mm3"] / CAPACITY).round(2)
    df.index.name = "date"
    return df.sort_index()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    arch = load_archive(args.out)
    today = dt.date.today()
    # 1) month-end history from the bulletins, only for water years not yet in the archive
    hist_idx = arch.index[arch["Source"] == "DGA monthly bulletin (month-end)"] if len(arch) else pd.DatetimeIndex([])
    missing = set()
    for y in BULLETINS:
        ends = pd.date_range(pd.Timestamp(y - 1, 10, 31), pd.Timestamp(y, 9, 30), freq="ME")
        ends = ends[ends < pd.Timestamp(API_START)]
        if not ends.isin(hist_idx).all():
            missing.add(y)
    hist = pd.DataFrame()
    if missing:
        print(f"Reading DGA September bulletins for water years ending {sorted(missing)}", flush=True)
        hist = fetch_bulletins(missing)
    # 2) daily API from the last saved day
    daily_arch = arch[arch["Source"] == "DGA online (daily median)"] if len(arch) else arch
    try:
        saved_method = str(pd.read_excel(args.out, sheet_name="Method", index_col=0).loc["daily_method", "value"])
    except Exception:  # noqa: BLE001 - first run / older workbook
        saved_method = None
    if daily_arch.empty or saved_method != DAILY_METHOD:
        start = API_START
    else:
        start = max(API_START, daily_arch.index.max().date() - dt.timedelta(days=REVISION_DAYS))
    print(f"Fetching DGA daily volumes {start} .. {today}", flush=True)
    new = fetch_daily(start, today)

    parts = [p for p in (arch, hist, new) if p is not None and not p.empty]
    if not parts:
        print("No data.", flush=True)
        sys.exit(1)
    base = pd.concat([p for p in (arch, hist) if p is not None and not p.empty] or [pd.DataFrame(columns=COLS)])
    base = base[~base.index.duplicated(keep="last")]
    # new readings win, but a station that returned nothing this run keeps its saved values
    new = new.dropna(how="all", subset=[c for c in new.columns if c.endswith("_Mm3")])
    df = new.combine_first(base) if not new.empty else base
    df = finish(df[df.index >= pd.Timestamp(HISTORY_START)])
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    stations = pd.DataFrame([{"column": f"{k}_Mm3", "reservoir": (v[1] or k), "DGA station code": v[0],
                              "capacity_Mm3": v[2], "purpose": v[3], "in total": k in TOTAL}
                             for k, v in RESERVOIRS.items()]).set_index("column")
    method = pd.DataFrame({"value": [DAILY_METHOD]}, index=pd.Index(["daily_method"], name="key"))
    xlsx_notes.write_workbook(args.out, {"Daily": df, "Reservoirs": stations, "Method": method}, NOTES,
                              {ln for ln in NOTES if ln and ln.isupper()})
    last = df.dropna(subset=["Total_Mm3"]).iloc[-1]
    print(f"Saved {args.out}: {len(df):,} rows ({df.index.min():%d-%b-%Y} to {df.index.max():%d-%b-%Y}); "
          f"latest total {last['Total_Mm3']:,.0f} Mm3 = {last['Total_pct']:.1f}% on {last.name:%d-%b-%Y}", flush=True)
    print(df.tail(5).to_string(), flush=True)


if __name__ == "__main__":
    main()
