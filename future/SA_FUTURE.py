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

RALIE_CKAN = "https://dadosabertos.aneel.gov.br/api/3/action/package_search"
AR_CKAN = "https://datos.energia.gob.ar/api/3/action/package_search"
AR_OBRAS = ("https://datos.energia.gob.ar/dataset/84bf14b4-47bf-4efb-ad7c-0d70684a50d3/resource/"
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
    # SIGA's DatEntradaOperacao is blank (read as day 0) for plants not yet operating: expected dates come from RALIE
    keep = [c for c in ("NomEmpreendimento", "SigUFPrincipal", "SigTipoGeracao", "DscFonteCombustivel", "Fuel",
                        "Phase", "MW") if c in p.columns]
    sheets["BR pipeline projects"] = p[keep].sort_values("MW", ascending=False).reset_index(drop=True)
    print(f"  pipeline {p['MW'].sum() / 1000:,.1f} GW: {p.groupby('Fuel')['MW'].sum().round(0).to_dict()}", flush=True)
    return sheets


def _ckan_csv(api, query, pattern):
    """URL of the first CSV resource matching `pattern` in a CKAN search."""
    for pkg in get(api, params={"q": query, "rows": 10}).json()["result"]["results"]:
        for res in pkg.get("resources", []):
            u = res.get("url", "")
            if re.search(pattern, u, re.I) and u.lower().endswith(".csv"):
                return u
    raise RuntimeError(f"no CSV matching {pattern} for '{query}'")


def _csv(content):
    for enc in ("utf-8", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(content), low_memory=False, encoding=enc, sep=None, engine="python")
        except UnicodeDecodeError:
            continue
    raise ValueError("unreadable CSV")


def brazil_ralie():
    """ANEEL RALIE (generation expansion monitoring): projects under construction / granted with their forecast
    commercial-operation date -> MW by forecast year and fuel."""
    url = _ckan_csv(RALIE_CKAN, "ralie", r"ralie.*usina|usina.*ralie|ralie")
    d = _csv(get(url).content)
    print(f"RALIE {url}: {d.shape}; columns {list(d.columns)}\n{d.head(3).to_string()[:1500]}", flush=True)
    mw = col(d, r"MdaPotenciaOutorgada|Potencia.*Outorgada|Potencia|MW")
    dates = [c for c in d.columns if re.search(r"Prev.*Oper.*Comerc|OperacaoComercial|Prev.*Comerc|Prev", c, re.I)]
    print(f"  date columns: {dates}", flush=True)
    if not dates:
        raise KeyError("no forecast-date column")
    when = pd.to_datetime(d[dates[0]], errors="coerce", dayfirst=True)
    for c in dates[1:]:   # first forecast date given among the candidates
        when = when.fillna(pd.to_datetime(d[c], errors="coerce", dayfirst=True))
    tipo = col(d, r"SigTipoGeracao|TipoGera", required=False)
    fonte = col(d, r"DscFonte|Combust|Fonte", required=False)
    text = (d[tipo].fillna("").astype(str) if tipo else "") + " " + (d[fonte].fillna("").astype(str) if fonte else "")
    from BRAZIL_ANEEL_CAPACITY_MONTHLY import siga_fuel
    fuel = [siga_fuel(t, f) for t, f in zip(d[tipo].fillna(""), d[fonte].fillna(""))] if tipo and fonte else \
        text.map(fuel_from_text)
    kw = pd.to_numeric(d[mw].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                       errors="coerce")
    t = d.assign(Fuel=fuel, MW=kw / 1000.0 if "kw" in mw.lower() else kw, Year=when.dt.year)
    t = t[t["Year"] >= date.today().year - 1]
    print(f"  {len(t)} dated projects, {t['MW'].sum() / 1000:,.1f} GW: {t.groupby('Year')['MW'].sum().round(0).to_dict()}",
          flush=True)
    return {"BR expected additions": by_year_fuel(t, "Year", "Fuel", "MW", end=2040)}, url


def argentina():
    try:
        content = get(AR_OBRAS).content
    except Exception as e:  # noqa: BLE001 - the direct link moves; look it up in the open-data catalogue
        print(f"  AR direct link failed ({type(e).__name__}); searching the catalogue", flush=True)
        content = get(_ckan_csv(AR_CKAN, "obras generacion", r"obra")).content
    d = _csv(content)
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
                        ("Brazil RALIE", brazil_ralie, "ANEEL RALIE generation expansion monitoring"),
                        ("Argentina", argentina, f"Secretaría de Energía, Obras de generación: {AR_OBRAS}")):
        try:
            got = fn()
            if isinstance(got, tuple):
                got, url = got
                s = f"{s}: {url}"
            sheets.update(got)
            src.append(s)
        except Exception as e:  # noqa: BLE001
            print(f"{name} failed: {type(e).__name__}: {e}", flush=True)
    if not sheets:
        raise SystemExit("nothing fetched")
    notes = [
        "UNITS",
        "MW by fuel. Brazil: ANEEL granted capacity of plants under construction or granted but not started "
        "(many granted wind/solar projects never get built - read 'not started' as an upper bound). Expected "
        "additions: ANEEL RALIE projects by forecast commercial-operation year. Argentina: "
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
