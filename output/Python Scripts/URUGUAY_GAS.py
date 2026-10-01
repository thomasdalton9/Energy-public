"""
Uruguay natural gas demand by sector, monthly from 2021, from the Ministry
of Industry, Energy and Mining's statistics visualiser (MIEM / DNE, area
Planificacion, Estadistica y Balance):
  https://visualpeb.miem.gub.uy/visualPEB/gas_natural
  report "Flujos de gas natural segun Tarifas" (gnFlujosTarifasMes),
  requested one month at a time, in m3.

Uruguay produces no gas: everything is imported from Argentina through
two pipelines - Gasoducto del Litoral (Paysandu, north-west) and Gasoducto
Cruz del Sur (to the south and Montevideo). Each monthly report is a flow
diagram from pipeline -> zone -> distributor -> tariff:
  Norte (Litoral):  Conecta (Norte) residential + general service, plus "Otros"
  Sur (Cruz del Sur): Conecta (Sur) residential + general service + large users,
                    plus "Otros" (MIEM's sector view calls it "Insumo del sector
                    Energia" = gas fed to power generation, i.e. Punta del Tigre)
  Montevideo:       Montevideo Gas residential + general service + large users,
                    plus "Otros" ("Consumo del sector Energia" = energy-sector own use)
"Otros" is what is left of a zone's pipeline import after the distributors'
billing, so it carries timing differences and can be small or noisy.

Sectors written here follow MIEM's own sector view (gnFlujosSectores):
  Residential  = Residencial (all three distributors)
  Commercial   = Servicio General (commerce, services, small industry)
  Industrial   = Grandes usuarios (Sur, Montevideo) + Otros (Norte)
  Power        = Otros (Sur)  - insumo del sector Energia
  Energy_own_use = Otros (Montevideo) - consumo del sector Energia
  Total        = sum of the above (= pipeline imports, within rounding)
Imports (Litoral + Cruz del Sur) are written alongside as a check.

Incremental: keeps the months already in the workbook, fetches missing
months up to the latest complete one, and re-fetches the last few months
(billing revisions). A month counts as complete once all three
distributors report residential billing and imports are present.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import datetime
import json
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts
import xlsx_notes

BASE = "https://visualpeb.miem.gub.uy/visualPEB/"
PAGE = BASE + "gas_natural"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)
DATA_START = pd.Timestamp(2021, 1, 1)
REFRESH_MONTHS = 3
MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]
RAW_SHEET, SECTOR_SHEET = "Flows by tariff (m3)", "Demand by sector"
DISTRIBUTORS = {"Norte": "Conecta (Norte)", "Sur": "Conecta (Sur)", "Mvdo.": "Montevideo Gas"}
SECTORS = ["Residential", "Commercial", "Industrial", "Power", "Energy_own_use"]


def out(*a):
    print(*a, flush=True)


def last_selectable_year():
    """The page's hidden 'ultimoAnioSeleccionable' (the report rejects requests without it)."""
    try:
        r = requests.get(PAGE, headers=H, timeout=T)
        m = re.search(r'name="ultimoAnioSeleccionable"\s+value="(\d{4})"', r.text)
        if m:
            return int(m.group(1))
    except requests.RequestException as e:
        out(f"gas page unavailable ({e})")
    return datetime.date.today().year


def fetch_month(month, last_year):
    """{leaf name: m3} for one month, e.g. 'Residencial (Mvdo.)', 'Gasoducto Cruz del Sur -> (Sur)'."""
    params = {"ultimoAnioSeleccionable": last_year, "mesDesde": MES[month.month - 1], "anioDesde": month.year,
              "mesHasta": MES[month.month - 1], "anioHasta": month.year, "unidades": "m³"}
    for attempt in range(3):
        try:
            r = requests.get(BASE + "gnFlujosTarifasMes", params=params, headers=H, timeout=T)
            if r.status_code == 200:
                break
            out(f"  {month:%Y-%m}: HTTP {r.status_code}")
        except requests.RequestException as e:
            out(f"  {month:%Y-%m}: {e}")
        time.sleep(3 * (attempt + 1))
    else:
        return None
    m = re.search(r"datos\s*=\s*(\[\[.*?\]\])", r.text, re.S)
    if not m:
        out(f"  {month:%Y-%m}: no flow table in the response")
        return None
    flows = {}
    for src, dst, val in json.loads(m.group(1)):
        key = f"{src} -> {dst}" if src.startswith("Gasoducto") else dst
        flows[key] = flows.get(key, 0.0) + float(val or 0)
    return flows


def complete(flows):
    if not flows:
        return False
    res = all(flows.get(f"Residencial ({z})", 0) > 0 for z in DISTRIBUTORS)
    imp = sum(v for k, v in flows.items() if k.startswith("Gasoducto")) > 0
    return res and imp


def load_existing(path):
    try:
        df = pd.read_excel(path, sheet_name=RAW_SHEET, index_col=0)
        df.index = pd.to_datetime(df.index.astype(str))
        out(f"existing archive: {len(df)} months {df.index.min():%Y-%m}..{df.index.max():%Y-%m}")
        return df
    except (FileNotFoundError, ValueError) as e:
        out(f"no existing archive ({e}); full pull from {DATA_START:%Y-%m}")
        return pd.DataFrame()


def sectors(raw):
    g = lambda *c: raw.reindex(columns=list(c)).sum(axis=1, min_count=1)  # noqa: E731
    s = pd.DataFrame({
        "Residential": g("Residencial (Norte)", "Residencial (Sur)", "Residencial (Mvdo.)"),
        "Commercial": g("Servicio General (Norte)", "Servicio General (Sur)", "Servicio General (Mvdo.)"),
        "Industrial": g("Grandes usuarios (Sur)", "Grandes usuarios (Mvdo.)", "Otros (Norte)"),
        "Power": g("Otros (Sur)"),
        "Energy_own_use": g("Otros (Mvdo.)"),
    })
    s["Total"] = s[SECTORS].sum(axis=1, min_count=1)
    s["Imports"] = raw[[c for c in raw.columns if c.startswith("Gasoducto")]].sum(axis=1, min_count=1)
    return s / 1e6  # million m3 per month


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/uruguay_gas_demand_by_sector.xlsx")
    ap.add_argument("--full", action="store_true", help="re-pull every month from 2021")
    args = ap.parse_args()

    raw = pd.DataFrame() if args.full else load_existing(args.out)
    last_year = last_selectable_year()
    today = pd.Timestamp(datetime.date.today()).to_period("M").to_timestamp()
    have = set(raw.index) if not raw.empty else set()
    refresh = set(sorted(have)[-REFRESH_MONTHS:])
    months = [m for m in pd.date_range(DATA_START, today, freq="MS") if m not in have or m in refresh]
    out(f"last selectable year {last_year}; months to fetch: {len(months)} "
        f"({months[0]:%Y-%m}..{months[-1]:%Y-%m})" if months else "nothing to fetch")

    rows = {}
    for m in months:
        if m.year > last_year:
            break
        flows = fetch_month(m, last_year)
        if not complete(flows):
            out(f"  {m:%Y-%m}: incomplete or not yet published ({flows and len(flows)} flows) - stopping here")
            if m >= today - pd.DateOffset(months=6):
                break  # recent months are published in order; later ones won't be complete either
            continue
        rows[m] = flows
        out(f"  {m:%Y-%m}: {len(flows)} flows, imports {sum(v for k, v in flows.items() if k.startswith('Gasoducto')) / 1e6:.2f} mcm")
    if rows:
        new = pd.DataFrame.from_dict(rows, orient="index")
        raw = new.combine_first(raw) if not raw.empty else new
    if raw.empty:
        raise SystemExit("no data fetched and no existing archive")
    raw = raw.sort_index()
    raw = raw[[c for c in sorted(raw.columns, key=lambda c: (c.startswith("Gasoducto"), c))]]

    monthly = sectors(raw)
    days = monthly.index.days_in_month
    perday = monthly.div(days, axis=0)
    table = pd.concat([monthly.round(3).add_suffix("_mcm_month"), perday.round(4).add_suffix("_mcm_per_day")], axis=1)
    gap = ((monthly["Total"] - monthly["Imports"]) / monthly["Imports"]).abs().max()
    out(f"{len(monthly)} months {monthly.index.min():%Y-%m}..{monthly.index.max():%Y-%m}; "
        f"max |sectors - imports| / imports = {gap:.2%}")
    out(perday.tail(6).round(3).to_string())

    notes = [
        "UNITS",
        "Demand by sector: *_mcm_month = million cubic metres in the month; *_mcm_per_day = the same divided by days "
        "in the month.",
        f"{RAW_SHEET}: every flow in MIEM's monthly report, in cubic metres, as published.",
        "",
        "SECTORS (following MIEM's own sector view of the same flows)",
        "Residential: residential tariff, all distributors (Montevideo Gas, Conecta Sur, Conecta Norte/Paysandu).",
        "Commercial: 'Servicio General' tariff - commerce, services and small industry on the distribution networks.",
        "Industrial: 'Grandes usuarios' (large users, Sur + Montevideo) plus 'Otros (Norte)' (large industrial use on "
        "the Litoral pipeline, which MIEM's sector view counts as Industrial).",
        "Power: 'Otros (Sur)' - MIEM's sector view calls it 'Insumo del sector Energia' (gas input to power "
        "generation on the Cruz del Sur southern branch).",
        "Energy_own_use: 'Otros (Mvdo.)' - 'Consumo del sector Energia' (energy-sector own use in Montevideo).",
        "The 'Otros' flows are what is left of each zone's pipeline import after distributor billing, so they absorb "
        "billing-cycle timing and can be noisy month to month.",
        "Total: sum of sectors. Imports: Gasoducto del Litoral + Gasoducto Cruz del Sur (should match Total).",
        "",
        "COVERAGE",
        f"Monthly from {DATA_START:%Y-%m}. Uruguay has no domestic production; all gas is imported from Argentina. "
        f"The last {REFRESH_MONTHS} months are re-fetched on every run (billing revisions); a month is added once all "
        "three distributors report.",
        "",
        "SOURCE",
        f"MIEM / DNE (Planificacion, Estadistica y Balance), VisualPEB gas natural - 'Flujos de gas natural segun "
        f"Tarifas' (monthly): {PAGE}",
        "Underlying series (downloadable zip): https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-"
        "estadisticas/datos/series-estadisticas-gas-natural",
    ]
    table.index = table.index.strftime("%Y-%m")
    table.index.name = "Month"
    raw_out = raw.copy()
    raw_out.index = raw_out.index.strftime("%Y-%m")
    raw_out.index.name = "Month"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {SECTOR_SHEET: table, RAW_SHEET: raw_out}, notes,
                              {"UNITS", "SECTORS (following MIEM's own sector view of the same flows)", "COVERAGE",
                               "SOURCE"})
    chart_df = perday[SECTORS].rename(columns={"Energy_own_use": "Energy own use"})
    xlsx_charts.add_chart_sheet(args.out, chart_df, "Uruguay gas demand by sector", "million m3/day",
                                kind="stacked_bar")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
