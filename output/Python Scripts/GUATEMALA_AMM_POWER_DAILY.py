"""
Guatemala daily power generation by fuel from AMM (Administrador del
Mercado Mayorista), in the standard raw grid-operator layout
(see power_daily_std.py): sheet "Daily" with date, Hydro_MWh, Gas_MWh,
Wind_MWh, Solar_MWh, Coal_MWh, Nuclear_MWh, Oil_MWh, Bioenergy_MWh,
Other_MWh, Total_MWh.

Sources (found via discovery_archive/south_america/NORTH_CENTRAL_AMERICA_POWER_PROBE2..6.py):
  1. From 2025-04-01: AMM "Grafica Generacion" web service
     https://wl12.amm.org.gt/GraficaPW/graficaCombustible?dt=dd/mm/yyyy
     -> hourly MW by fuel (AGUA, CARBON, BUNKER, ...); day MWh = sum of the
     24 hourly values. AMM's own fuel split, so sugar mills burning coal or
     bunker out of harvest season count as Coal / Oil, bagasse as Bioenergy.
     The service returns [] before 2025-04-01.
  2. 2021-01-01 to 2025-03-31: AMM daily Posdespacho workbooks
     https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/<yyyy>/<mm_MES>/PD<yyyymmdd>.zip
     sheet 'Carga Horaria' = hourly MW per generating unit with a daily-energy
     row (MWh). Units are coded <plant>-<type>[n]: -H hydro, -F solar,
     -E wind, ZUN-G/ORT-G/CAL-G geothermal, -I interconnection (excluded);
     every other unit is thermal (coal, bunker, diesel, gas, sugar-mill
     cogeneration, biogas). AMM publishes no daily fuel per unit, so each
     day's thermal energy is split by that month's thermal fuel shares from
     AMM's monthly 'GENERACION POR TIPO DE RECURSO' table (GM<date>.xlsx,
     https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/pubamm/<yyyy>/).
     Daily thermal totals are AMM's; the coal/oil/gas/biomass split within a
     day is the month's average.
Also writes sheet "AMM_Monthly_GWh" (AMM's monthly generation by resource,
GWh) for checking.

Incremental: keeps saved days; fetches missing days plus the last 14
(AMM revises recent days). History from 2021-01-01.

Usage: python3 GUATEMALA_AMM_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD] [--budget-min N]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import os
import re
import sys
import time
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/guatemala_power_generation_daily.xlsx"
GRAFICA = "https://wl12.amm.org.gt/GraficaPW/graficaCombustible"
GRAFICA_START = dt.date(2025, 4, 1)
PD_URL = "https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/{y}/{m:02d}_{mes}/PD{y}{m:02d}{d:02d}.zip"
GM_URL = "https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/pubamm/{y}/GM{stamp}.xlsx"
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE",
         "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
FUELS = std.FUELS
COLS = [f"{f}_MWh" for f in FUELS]
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}

# GraficaPW fuel ('tipo', accents stripped, upper case) -> standard fuel; None = not generation
GRAFICA_MAP = {
    "AGUA": "Hydro",
    "IRRADIACION": "Solar",
    "VIENTO": "Wind",
    "CARBON": "Coal", "CARBON/PETCOKE": "Coal",
    "BUNKER": "Oil", "DIESEL": "Oil",
    "GAS NATURAL": "Gas",
    "BIOMASA": "Bioenergy", "BIOGAS": "Bioenergy",
    "VAPOR": "Other",     # geothermal steam (Orzunil, Ortitlan)
    "SYNGAN": "Other",    # syngas (ESUS)
    "DEM SNI": None, "DEMANDA LOCAL PROG": None, "INTERCONEXION": None, "NULL": None, "NONE": None, "": None,
}
# AMM monthly 'GENERACION POR TIPO DE RECURSO' row -> standard fuel
GM_MAP = {"GEOTERMICA": "Other", "EOLICA": "Wind", "SOLAR": "Solar", "HIDROELECTRICA": "Hydro",
          "BIOGAS": "Bioenergy", "BIOMASA": "Bioenergy", "SYNGAS": "Other", "GAS NATURAL": "Gas",
          "CARBON MINERAL": "Coal", "COQUE DE PETROLEO": "Coal", "BUNKER": "Oil", "DIESEL": "Oil"}
GEOTHERMAL_UNITS = ("ZUN-G", "ORT-G", "CAL-G")

NOTES = [
    "UNITS",
    "MWh per day (energy generated in Guatemala's national grid, SNI). Imports/exports (Mexico, MER) are excluded.",
    "",
    "COVERAGE",
    "Daily from 2021-01-01. Updated daily by GitHub Actions (guatemala_power_generation_daily.yml): missing days plus "
    "the last 14 are fetched each run. Days the source has not published yet are absent.",
    "",
    "SOURCE",
    "AMM (Administrador del Mercado Mayorista), https://www.amm.org.gt/",
    "From 2025-04-01: AMM 'Grafica Generacion' service, https://wl12.amm.org.gt/GraficaPW/graficaCombustible?dt=dd/mm/yyyy "
    "(hourly MW by fuel; a day's MWh = sum of its 24 hourly values; days with fewer than 23 hours are skipped and "
    "retried).",
    "2021-01-01 to 2025-03-31: AMM daily Posdespacho workbooks, https://www.amm.org.gt/pdfs2/post_despacho/"
    "POSDESPACHO_DIARIO/<yyyy>/<mm_MES>/PD<yyyymmdd>.zip, sheet 'Carga Horaria' (daily energy per generating unit). "
    "Unit suffix -H = hydro, -F = solar, -E = wind, ZUN-G/ORT-G/CAL-G = geothermal, -I = interconnection (excluded); "
    "all other units are thermal. AMM gives no daily fuel per unit (sugar mills burn bagasse in harvest and coal or "
    "bunker outside it), so each day's thermal energy is split by that month's thermal fuel shares in AMM's monthly "
    "'GENERACION POR TIPO DE RECURSO' table (GM<date>.xlsx, https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/"
    "pubamm/<yyyy>/). Daily thermal totals are AMM's; the split between coal/oil/gas/biomass within a day is the "
    "month's average.",
    "Sheet 'AMM_Monthly_GWh': AMM's monthly generation by resource (GWh) from the GM workbooks, for checking.",
    "",
    "MAPPING",
    "Hydro_MWh = AGUA (hydro, incl. run-of-river and distributed hydro).",
    "Solar_MWh = IRRADIACION (solar PV). Wind_MWh = VIENTO.",
    "Coal_MWh = CARBON + CARBON/PETCOKE (coal and petroleum coke; includes sugar mills burning coal off-season).",
    "Oil_MWh = BUNKER + DIESEL (bunker engines and steam, diesel turbines).",
    "Gas_MWh = GAS NATURAL (Ocultun, Ooxol, Innova engines; Actun Can turbine).",
    "Bioenergy_MWh = BIOMASA + BIOGAS (sugar-mill bagasse cogeneration, landfill biogas).",
    "Other_MWh = VAPOR (geothermal: Orzunil, Ortitlan - geothermal is reported in Other) + SYNGAN (syngas).",
    "Nuclear_MWh = 0 (none in Guatemala).",
    "Total_MWh = sum of the fuel columns.",
    "Monthly-table rows used for 2021-01..2025-03 thermal shares: Carbon Mineral + Coque de Petroleo -> Coal; "
    "Bunker + Diesel -> Oil; Gas Natural -> Gas; Biomasa + Biogas -> Bioenergy; Syngas -> Other.",
]

S = requests.Session()
S.headers.update(HEADERS)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def get(url, **kw):
    for attempt in range(3):
        try:
            r = S.get(url, timeout=90, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(5 * (attempt + 1))
                continue
            return r
        except requests.RequestException as e:
            print(f"  retry {url}: {type(e).__name__}", flush=True)
            time.sleep(5 * (attempt + 1))
    return None


# ---------- source 1: GraficaPW (2025-04-01 on) ----------
def grafica_day(day):
    r = get(GRAFICA, params={"dt": day.strftime("%d/%m/%Y")})
    if r is None or not r.ok:
        return day, None, set()
    try:
        rows = r.json()
    except ValueError:
        return day, None, set()
    if not rows:
        return day, None, set()
    df = pd.DataFrame(rows)
    df["potencia"] = pd.to_numeric(df["potencia"], errors="coerce")
    df["tipo"] = df["tipo"].map(norm)
    hours = df.loc[df["tipo"] == "AGUA", "hora"].nunique()
    if hours < 23:
        return day, None, set()
    unknown = {t for t in df["tipo"].unique() if t not in GRAFICA_MAP}
    out = dict.fromkeys(FUELS, 0.0)
    for tipo, mwh in df.groupby("tipo")["potencia"].sum().items():
        fuel = GRAFICA_MAP.get(tipo, "Other")
        if fuel:
            out[fuel] += float(mwh)
    return day, out, unknown


# ---------- source 2: daily Posdespacho (2021-01-01 .. 2025-03-31) ----------
def unit_class(code):
    code = str(code).strip()
    if code.startswith(GEOTHERMAL_UNITS):
        return "Geothermal"
    m = re.search(r"-([A-Z])\d*$", code)
    if not m:
        return None
    return {"H": "Hydro", "F": "Solar", "E": "Wind", "I": None}.get(m.group(1), "Thermal")


def posdespacho_day(day):
    url = PD_URL.format(y=day.year, m=day.month, d=day.day, mes=MESES[day.month - 1])
    r = get(url)
    if r is None or not r.ok or r.content[:2] != b"PK":
        return day, None, None
    try:
        z = zipfile.ZipFile(io.BytesIO(r.content))
        name = next(n for n in z.namelist() if n.lower().endswith((".xlsx", ".xls", ".xlsm")))
        ch = pd.read_excel(io.BytesIO(z.read(name)), sheet_name="Carga Horaria", header=None)
    except Exception as e:  # noqa: BLE001
        print(f"  {day}: cannot read Posdespacho ({type(e).__name__}: {e})", flush=True)
        return day, None, None
    hdr = None
    for i in range(min(15, len(ch))):
        vals = [str(v) for v in ch.iloc[i].tolist()]
        if sum(bool(re.search(r"-[A-Z]\d*$", v)) for v in vals) > 20:
            hdr = i
            break
    if hdr is None:
        print(f"  {day}: no unit header row in 'Carga Horaria'", flush=True)
        return day, None, None
    codes = [str(v).strip() for v in ch.iloc[hdr].tolist()]
    energy = pd.to_numeric(ch.iloc[hdr + 1], errors="coerce")
    out = {"Hydro": 0.0, "Solar": 0.0, "Wind": 0.0, "Geothermal": 0.0, "Thermal": 0.0}
    units = {}
    for code, mwh in zip(codes, energy):
        cls = unit_class(code)
        if cls is None or pd.isna(mwh):
            continue
        out[cls] += float(mwh)
        units[code] = float(mwh)
    total = next((float(v) for c, v in zip(codes, energy) if c.lower().startswith("total gen") and pd.notna(v)), None)
    return day, out, {"units": units, "total_generado": total}


# ---------- AMM monthly by resource (GM workbooks) ----------
def gm_year(year, today):
    """AMM monthly generation by resource for one year (GWh), from the latest GM<date>.xlsx of that year."""
    stamps = [f"{year}12{d:02d}" for d in range(31, 14, -1)]
    if year == today.year:
        stamps = [(today - dt.timedelta(days=k)).strftime("%Y%m%d") for k in range(0, 60)]
    for stamp in stamps:
        r = get(GM_URL.format(y=year, stamp=stamp))
        if r is None or not r.ok or r.content[:2] != b"PK":
            continue
        df = pd.read_excel(io.BytesIO(r.content), sheet_name=0, header=None)
        start = next((i for i in range(len(df)) if "TIPO DE RECURSO" in norm(df.iloc[i, 1])), None)
        if start is None:
            continue
        rows = {}
        for i in range(start + 1, len(df)):
            name = str(df.iloc[i, 1]).strip()
            if norm(name) == "TOTAL" or name == "nan":
                break
            rows[name] = pd.to_numeric(df.iloc[i, 2:14], errors="coerce").to_numpy()
        out = pd.DataFrame(rows, index=pd.date_range(f"{year}-01-01", periods=12, freq="MS"))
        out = out[out.sum(axis=1) > 0]
        print(f"  GM {stamp}: {len(out)} months, resources {list(out.columns)}", flush=True)
        return out
    print(f"  GM {year}: no workbook found", flush=True)
    return pd.DataFrame()


def thermal_shares(gm):
    """Month -> {fuel: share of thermal} from AMM's monthly table."""
    shares = {}
    for month, row in gm.iterrows():
        parts = {}
        for name, val in row.items():
            fuel = GM_MAP.get(norm(name))
            if fuel in ("Coal", "Oil", "Gas", "Bioenergy") or norm(name) == "SYNGAS":
                parts[fuel] = parts.get(fuel, 0.0) + (0.0 if pd.isna(val) else float(val))
        tot = sum(parts.values())
        if tot > 0:
            shares[month.date().replace(day=1)] = {k: v / tot for k, v in parts.items()}
    return shares


def to_frame(rows):
    df = pd.DataFrame.from_dict(rows, orient="index")
    df = df.reindex(columns=FUELS).fillna(0.0)
    df.index = pd.to_datetime(df.index)
    df = df.rename(columns={f: f"{f}_MWh" for f in FUELS})
    df["Total_MWh"] = df[COLS].sum(axis=1)
    df.index.name = "date"
    return df.sort_index().round(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=std.HISTORY_START)
    ap.add_argument("--budget-min", type=float, default=45, help="stop fetching after this many minutes")
    ap.add_argument("--workers", type=int, default=4)
    args = ap.parse_args()
    t0 = time.time()
    today = dt.date.today()

    print("MAPPING (GraficaPW fuel -> column):", {k: v for k, v in GRAFICA_MAP.items() if v}, flush=True)
    print("MAPPING (Posdespacho unit suffix): -H Hydro, -F Solar, -E Wind, ZUN-G/ORT-G/CAL-G Other (geothermal), "
          "-I excluded, rest thermal split by AMM monthly shares " + str(GM_MAP), flush=True)

    daily = std.load_sheet(args.out, "Daily")
    if not daily.empty:
        daily = daily.reindex(columns=COLS + ["Total_MWh"]).fillna(0.0)
    end = today - dt.timedelta(days=1)
    days = std.missing_days(daily, args.start, end, refresh_days=14)
    new_g = sorted([d for d in days if d >= GRAFICA_START], reverse=True)
    new_p = sorted([d for d in days if d < GRAFICA_START], reverse=True)
    print(f"{0 if daily.empty else daily['Total_MWh'].notna().sum():,} days saved; to fetch: {len(new_g)} from GraficaPW, "
          f"{len(new_p)} from Posdespacho", flush=True)

    gm = {}
    years = sorted({d.year for d in new_p} | {today.year, today.year - 1})
    for y in years:
        gm[y] = gm_year(y, today)
    gm_all = pd.concat([g for g in gm.values() if not g.empty]) if any(not g.empty for g in gm.values()) else pd.DataFrame()
    shares = thermal_shares(gm_all) if not gm_all.empty else {}

    rows, unknown_all = {}, set()

    def save():
        nonlocal daily
        if rows:
            daily = std.merge(to_frame(rows), daily)
        write(args.out, daily, gm_all)

    with ThreadPoolExecutor(args.workers) as ex:
        for k in range(0, len(new_g), 40):
            if time.time() - t0 > args.budget_min * 60:
                print("time budget reached (GraficaPW)", flush=True)
                break
            for day, out, unknown in ex.map(grafica_day, new_g[k:k + 40]):
                if out:
                    rows[day] = out
                unknown_all |= unknown
            print(f"  GraficaPW: {min(len(new_g), k + 40)}/{len(new_g)} days requested, {len(rows)} rows so far", flush=True)
        if unknown_all:
            print(f"  WARNING: GraficaPW fuels not in the mapping (put in Other): {sorted(unknown_all)}", flush=True)
        save()

        p_done = 0
        for k in range(0, len(new_p), 40):
            if time.time() - t0 > args.budget_min * 60:
                print(f"time budget reached; {len(new_p) - k} Posdespacho days left for later runs", flush=True)
                break
            for day, out, extra in ex.map(posdespacho_day, new_p[k:k + 40]):
                if not out:
                    continue
                share = shares.get(day.replace(day=1))
                row = dict.fromkeys(FUELS, 0.0)
                row.update(Hydro=out["Hydro"], Solar=out["Solar"], Wind=out["Wind"], Other=out["Geothermal"])
                if share:
                    for fuel, s in share.items():
                        row[fuel] += out["Thermal"] * s
                else:
                    print(f"  {day}: no AMM monthly shares; thermal left in Other", flush=True)
                    row["Other"] += out["Thermal"]
                rows[day] = row
                p_done += 1
            print(f"  Posdespacho: {min(len(new_p), k + 40)}/{len(new_p)} days requested, {p_done} parsed", flush=True)
            if (k // 40) % 5 == 4:
                save()
    save()
    if daily.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    check(daily, gm_all)


def write(path, daily, gm_all):
    extra = {}
    if gm_all is not None and not gm_all.empty:
        g = gm_all.copy()
        g.index.name = "month"
        extra["AMM_Monthly_GWh"] = g.round(3)
    std.write(path, daily.reindex(columns=COLS + ["Total_MWh"]), NOTES, extra)


def check(daily, gm_all):
    """Monthly totals from this workbook vs AMM's own monthly table (GWh)."""
    if gm_all is None or gm_all.empty:
        return
    m = daily.resample("MS").sum(min_count=1) / 1000
    n = daily["Total_MWh"].resample("MS").count()
    full = n[n == n.index.days_in_month].index
    ref = pd.DataFrame(index=gm_all.index)
    for name in gm_all.columns:
        fuel = GM_MAP.get(norm(name))
        if fuel:
            ref[fuel] = ref.get(fuel, 0) + gm_all[name].fillna(0)
    ref["Total"] = ref.sum(axis=1)
    idx = [i for i in ref.index if i in full]
    if not idx:
        return
    print("CHECK vs AMM monthly table (GWh, this workbook / AMM):", flush=True)
    for i in idx[-15:]:
        parts = []
        for f in ["Hydro", "Coal", "Oil", "Bioenergy", "Solar", "Total"]:
            a = m.loc[i, f"{f}_MWh"] if f != "Total" else m.loc[i, "Total_MWh"]
            b = ref.loc[i, f] if f in ref else float("nan")
            parts.append(f"{f} {a:,.0f}/{b:,.0f}")
        print(f"  {i:%Y-%m}: " + "; ".join(parts), flush=True)


if __name__ == "__main__":
    main()
