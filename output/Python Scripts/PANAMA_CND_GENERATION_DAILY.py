"""
Panama daily power generation by type from CND (Centro Nacional de Despacho,
ETESA) - raw grid-operator pull in the standard layout.

Source: CND "Reporte Diario" (Informes de Operaciones -> Reporte Diario,
https://www.cnd.com.pa/index.php/informes/categoria/informes-de-operaciones),
one Excel file per day, "REPORTE DIARIO YYYYMMDD.xls". Its first sheet
(REPDIARIO) has the block "ENTREGADO AL SISTEMA - Detalle (MWh)": Hidro, Bunker,
Diesel, Autogeneradores Hidros / Termicos / Solares, Eolicos, Carbon, Solares,
Biogas, Gas Natural and Intercambio. The cnd.com.pa page lists and serves the
files through CND's document service (sitioprivado.cnd.com.pa), using the
public access key embedded in the site's own JavaScript
(sitiopublico.cnd.com.pa/modules/mod_logic_informescnd/assets/js/app.js):
  list:     https://sitioprivado.cnd.com.pa/Informe/GetListOperativosComerciales
            ?page=0&publico=1&key=KEY&categoria=6&tipo=110&anio=YYYY&mes=M&semana=0&dia=0
  download: https://sitioprivado.cnd.com.pa/Informe/Download/{id}?key=KEY
History from 2021-01-01. Found via
discovery_archive/south_america/CENTRAL_AMERICA_POWER_DISCOVERY3.py / 4.py.

Each run fetches only days missing from the workbook plus the last 14 days
(CND re-issues some reports), within a time budget, saving as it goes.

Usage: python3 PANAMA_CND_GENERATION_DAILY.py [--out PATH] [--start YYYY-MM-DD]
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import central_america_power_common as C  # noqa: E402

OUT = "output/Data and Chart Outputs/panama_power_generation_daily.xlsx"
PAGE = "https://www.cnd.com.pa/index.php/informes/categoria/informes-de-operaciones"
API = "https://sitioprivado.cnd.com.pa/"
KEY = "VXd9e23Z9JRA5aIUR21R-P8gocoGOMqdvSo79FduN"  # public key from cnd.com.pa's own page script
CATEGORIA, TIPO = 6, 110  # Informes de Operaciones / Reporte Diario

# 'ENTREGADO AL SISTEMA' line (accents/case ignored) -> standard fuel. None = not generation.
MAPPING = {
    "Hidro": "Hydro",
    "Autogeneradores Hidros": "Hydro",
    "Bunker": "Oil",
    "Diesel": "Oil",
    "Autogeneradores Termicos": "Oil",
    "Eolicos": "Wind",
    "Solares": "Solar",
    "Autogeneradores Solares": "Solar",
    "Carbon": "Coal",
    "Biogas": "Bioenergy",
    "Gas Natural": "Gas",
    "Intercambio": None,
}
KEYS = {C.norm(k): v for k, v in MAPPING.items()}

NOTES = [
    "UNITS",
    "MWh per day, as published by CND in the daily report's 'ENTREGADO AL SISTEMA - Detalle' block "
    "(energy delivered to the grid per type, MWh).",
    "",
    "COVERAGE",
    "@COVERAGE",
    "",
    "SOURCE",
    "CND / ETESA Panama - Informes de Operaciones, 'Reporte Diario' (REPORTE DIARIO YYYYMMDD.xls): " + PAGE,
    "Files listed/served by CND's document service " + API + "Informe/GetListOperativosComerciales "
    "(categoria=6, tipo=110) and " + API + "Informe/Download/{id}, with the public key from the site's page script.",
    "Updated daily by GitHub Actions (panama_power_generation.yml): only missing days plus the last 14.",
    "",
    "MAPPING",
    "CND line -> column in sheet 'Daily':",
    "Hidro + Autogeneradores Hidros -> Hydro_MWh",
    "Gas Natural (LNG-fired: AES Colon combined cycle, Sinolam Gatun) -> Gas_MWh",
    "Eolicos -> Wind_MWh",
    "Solares + Autogeneradores Solares -> Solar_MWh",
    "Carbon (coal) -> Coal_MWh",
    "Bunker + Diesel + Autogeneradores Termicos (self-generators' thermal units, oil-fired) -> Oil_MWh",
    "Biogas -> Bioenergy_MWh",
    "Intercambio (regional interchange; CND marks it Importando/Exportando) -> not generation, left out; "
    "shown signed in sheet 'Detail' (Intercambio_MWh, + import / - export).",
    "Panama has no nuclear or geothermal generation: Nuclear_MWh and Other_MWh are 0.",
    "Self-generator lines are net deliveries to the grid and are occasionally negative (net withdrawal); "
    "negative values are set to 0 in 'Daily' (raw values kept in 'Detail').",
    "Total_MWh = sum of the fuel columns (= CND 'Total Generado' except on days with a negative "
    "self-generator line). Sheet 'Detail' also has CND's 'Total Generado' and 'Total Entregado'.",
]


def session():
    s = requests.Session()
    s.headers.update({"User-Agent": C.USER_AGENT, "Referer": PAGE})
    return s


def list_month(s, year, month):
    """{report date: download id} for reports published in year-month (latest id wins)."""
    r = s.get(API + "Informe/GetListOperativosComerciales",
              params={"page": 0, "publico": 1, "key": KEY, "categoria": CATEGORIA, "tipo": TIPO,
                      "anio": year, "mes": month, "semana": 0, "dia": 0}, timeout=(15, 90))
    r.raise_for_status()
    out = {}
    for x in r.json():
        name = ((x.get("adjunto") or {}).get("path") or "").split("\\")[-1]
        m = re.search(r"(20\d{2})(\d{2})(\d{2})", name)
        if not m or not name.lower().endswith((".xls", ".xlsx")):
            continue
        try:
            day = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        if day not in out or x["id"] > out[day]:
            out[day] = x["id"]
    return out


def parse_report(content):
    """CND daily report bytes -> {line label: MWh} from 'ENTREGADO AL SISTEMA', plus totals."""
    engine = "openpyxl" if content[:2] == b"PK" else "xlrd"
    raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None, engine=engine)
    out = {}
    inside = False
    for _, row in raw.iterrows():
        label = C.norm(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
        if label in ("total generado", "total entregado"):
            out[label.title()] = pd.to_numeric(row.iloc[3], errors="coerce")
        if label == "entregado al sistema":
            inside = True
            continue
        if not inside or label in ("", "detalle"):
            continue
        if label.startswith("maximos"):
            break
        val = pd.to_numeric(row.iloc[2], errors="coerce")
        if label == "intercambio" and C.norm(row.iloc[1]).startswith("export"):
            val = -val
        out[label] = val
    return out


def to_frames(per_day):
    detail = pd.DataFrame(per_day).T
    detail.index = pd.to_datetime(detail.index)
    detail.index.name = "date"
    fuels = pd.DataFrame(index=detail.index)
    unknown = []
    for col in detail.columns:
        if col in ("Total Generado", "Total Entregado"):
            continue
        if col not in KEYS:
            unknown.append(col)
            fuel = "Other"
        else:
            fuel = KEYS[col]
        if fuel is None:
            continue
        fuels[fuel] = fuels.get(fuel, 0) + detail[col].fillna(0).clip(lower=0)
    if unknown:
        print(f"  WARNING: CND lines not in MAPPING, counted as Other: {unknown}", flush=True)
    labels = {C.norm(k): k for k in MAPPING}
    detail = detail.rename(columns=lambda c: (labels.get(c, c).replace(" ", "_") + "_MWh")
                           if c not in ("Total Generado", "Total Entregado") else c.replace(" ", "_") + "_MWh")
    return C.standardise(fuels), detail.round(2)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=C.HISTORY_START)
    ap.add_argument("--refresh-days", type=int, default=14)
    ap.add_argument("--budget-min", type=float, default=50, help="stop fetching after this many minutes")
    ap.add_argument("--save-every", type=int, default=60, help="write the workbook every N reports")
    args = ap.parse_args()
    t0 = time.time()
    C.print_mapping(MAPPING, "CND 'ENTREGADO AL SISTEMA' line -> standard fuel; None = excluded")

    daily = C.load_sheet(args.out, "Daily")
    detail = C.load_sheet(args.out, "Detail")
    end = dt.date.today() - dt.timedelta(days=1)
    days = C.days_to_fetch(daily, args.start, end, args.refresh_days)
    print(f"{len(daily):,} days saved; {len(days):,} to fetch", flush=True)
    if not days:
        return
    s = session()

    # A day's report is published that day or within the next few days: list those months.
    months = sorted({(d.year, d.month) for d in days} | {((d + dt.timedelta(days=7)).year,
                                                         (d + dt.timedelta(days=7)).month) for d in days})
    ids = {}
    for y, m in months:
        if (y, m) > (dt.date.today().year, dt.date.today().month):
            continue
        try:
            for day, i in list_month(s, y, m).items():
                if day not in ids or i > ids[day]:
                    ids[day] = i
        except (requests.RequestException, ValueError) as e:
            print(f"  list {y}-{m:02d}: {type(e).__name__}: {e}", flush=True)
    todo = sorted((d for d in days if d in ids), reverse=True)  # newest first
    print(f"{len(ids):,} reports listed; {len(todo):,} of the {len(days):,} days have a report", flush=True)

    per_day, failed = {}, []

    def flush():
        nonlocal daily, detail, per_day
        if per_day:
            new, new_detail = to_frames(per_day)
            daily, detail = C.merge(new, daily), C.merge(new_detail, detail)
            per_day = {}
        if not daily.empty:
            notes = [x for line in NOTES for x in (C.coverage_lines(daily) if line == "@COVERAGE" else [line])]
            C.write(args.out, daily, notes, detail)

    for k, day in enumerate(todo, 1):
        if time.time() - t0 > args.budget_min * 60:
            print(f"time budget reached; {len(todo) - k + 1} reports left for later runs", flush=True)
            break
        try:
            r = s.get(API + f"Informe/Download/{ids[day]}", params={"key": KEY}, timeout=(15, 120))
            r.raise_for_status()
            got = parse_report(r.content)
            gen = sum(v for kk, v in got.items() if kk not in ("intercambio", "Total Generado", "Total Entregado")
                      and pd.notna(v))
            tg = got.get("Total Generado")
            if pd.notna(tg) and abs(gen - tg) > max(5, 0.005 * tg):
                print(f"  {day}: lines sum to {gen:,.1f} MWh vs Total Generado {tg:,.1f}", flush=True)
            if gen <= 0:
                raise ValueError("no ENTREGADO AL SISTEMA lines found")
            per_day[pd.Timestamp(day)] = got
        except Exception as e:  # noqa: BLE001 - one bad file must not stop the run
            print(f"  {day} (id {ids[day]}): {type(e).__name__}: {e}", flush=True)
            failed.append(day)
            time.sleep(2)
        if k % 50 == 0:
            print(f"  {k}/{len(todo)} ({day})", flush=True)
        if k % args.save_every == 0:
            flush()
    flush()
    if failed:
        print(f"{len(failed)} reports failed: {[str(d) for d in failed[:20]]}", flush=True)
    if daily.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    C.print_monthly(daily)


if __name__ == "__main__":
    main()
