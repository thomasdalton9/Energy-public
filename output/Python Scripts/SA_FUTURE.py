"""
South America "future" workbook (one-off, run on demand): generation projects in the pipeline.

  Brazil     ANEEL SIGA plant register (the one brazil_power_capacity uses): plants with DscFaseUsina
             "Construção" (under construction) and "Construção não iniciada" (granted, construction not started),
             with granted capacity (MdaPotenciaOutorgadaKw) and any expected-operation date in the register
  Argentina  Secretaría de Energía open data "Obras de generación de electricidad" (generation works):
             http://datos.energia.gob.ar/ ... obras-generacion.csv

Writes south_america_future.xlsx. Other countries: no machine-readable pipeline found yet (Chile's Coordinador
blocks automated access; Colombia UPME and Peru COES publish PDF / web lists).

Usage: python3 future/SA_FUTURE.py [--out "output/Data and Chart Outputs/south_america_future.xlsx"]
"""
import argparse
import io
import os
import re
import sys
from datetime import date

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
from future_common import FUELS, by_year_fuel, col, fuel_from_text, get  # noqa: E402

AR_OBRAS = ("http://datos.energia.gob.ar/dataset/84bf14b4-47bf-4efb-ad7c-0d70684a50d3/resource/"
            "535ed5fc-ca6d-463e-9f88-2fe33b7b3790/download/obras-generacion.csv")
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "south_america_future.xlsx")
PHASES = {"Construção": "Under construction", "Construção não iniciada": "Granted, construction not started"}


def brazil():
    import BRAZIL_ANEEL_CAPACITY as siga
    from BRAZIL_ANEEL_CAPACITY_MONTHLY import siga_fuel

    raw = siga.fetch_raw()
    print(f"SIGA: {len(raw):,} rows; phases {raw['DscFaseUsina'].value_counts().to_dict()}", flush=True)
    print(f"  date columns: {[c for c in raw.columns if c.startswith('Dat')]}", flush=True)
    p = raw[raw["DscFaseUsina"].isin(PHASES)].copy()
    p["Phase"] = p["DscFaseUsina"].map(PHASES)
    p["MW"] = siga.to_float_br(p["MdaPotenciaOutorgadaKw"]) / 1000.0
    p["Fuel"] = [siga_fuel(t, f) for t, f in zip(p["SigTipoGeracao"], p["DscFonteCombustivel"])]
    sheets = {}
    ph = p.pivot_table(index="Phase", columns="Fuel", values="MW", aggfunc="sum").fillna(0)
    ph = ph.reindex(columns=[f for f in FUELS if f in ph.columns]).round(0)
    ph.index.name = "Phase"
    sheets["BR pipeline by phase"] = ph
    dcol = next((c for c in p.columns if re.search(r"Previs|Prevista|EntradaOperacao", c)), None)
    if dcol:
        p["Year"] = pd.to_datetime(p[dcol], errors="coerce").dt.year
        print(f"  expected-date column {dcol}: {p['Year'].notna().sum()} of {len(p)} projects dated", flush=True)
        if p["Year"].notna().any():
            sheets["BR pipeline by year"] = by_year_fuel(p, "Year", "Fuel", "MW", start=date.today().year - 1)
    keep = [c for c in ("NomEmpreendimento", "SigUFPrincipal", "SigTipoGeracao", "DscFonteCombustivel", "Fuel",
                        "Phase", "MW", dcol) if c and c in p.columns]
    sheets["BR pipeline projects"] = p[keep].sort_values("MW", ascending=False).reset_index(drop=True)
    print(f"  pipeline {p['MW'].sum() / 1000:,.1f} GW: {p.groupby('Fuel')['MW'].sum().round(0).to_dict()}", flush=True)
    return sheets


def argentina():
    d = pd.read_csv(io.BytesIO(get(AR_OBRAS).content), low_memory=False, encoding="utf-8", sep=None, engine="python")
    print(f"AR obras: {d.shape}; columns {list(d.columns)}\n{d.head(3).to_string()[:1500]}", flush=True)
    tech = col(d, r"tecnolog|fuente|tipo")
    mw = col(d, r"potencia|mw")
    d["Fuel"] = d[tech].map(fuel_from_text)
    d["MW"] = pd.to_numeric(d[mw].astype(str).str.replace(",", "."), errors="coerce")
    sheets = {}
    st = col(d, r"estado|situaci", required=False)
    if st:
        t = d.pivot_table(index=st, columns="Fuel", values="MW", aggfunc="sum").fillna(0).round(0)
        t.index.name = "Status"
        sheets["AR works by status"] = t
    yc = col(d, r"a[ñn]o|fecha.*(habilit|operaci|fin)|habilitaci", required=False)
    if yc:
        yr = pd.to_numeric(d[yc], errors="coerce")
        if yr.isna().all():
            yr = pd.to_datetime(d[yc], errors="coerce", dayfirst=True).dt.year
        if yr.notna().any():
            sheets["AR works by year"] = by_year_fuel(d.assign(_y=yr), "_y", "Fuel", "MW", start=2021)
    sheets["AR works list"] = d.reset_index(drop=True)
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    sheets, src = {}, []
    for name, fn, s in (("Brazil", brazil, "ANEEL SIGA plant register (phase: under construction / not started)"),
                        ("Argentina", argentina, f"Secretaría de Energía, Obras de generación: {AR_OBRAS}")):
        try:
            sheets.update(fn())
            src.append(s)
        except Exception as e:  # noqa: BLE001
            print(f"{name} failed: {type(e).__name__}: {e}", flush=True)
    if not sheets:
        raise SystemExit("nothing fetched")
    notes = [
        "UNITS",
        "MW by fuel. Brazil: ANEEL granted capacity of plants under construction or granted but not started "
        "(many granted wind/solar projects never get built - read 'not started' as an upper bound). Argentina: "
        "generation works listed by the Secretaría de Energía, MW as published.",
        "",
        "COVERAGE",
        "Brazil, Argentina. Chile, Colombia, Peru and the rest: no machine-readable project pipeline found yet.",
        "",
        "SOURCE",
        *src,
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: {list(sheets)}", flush=True)


if __name__ == "__main__":
    main()
