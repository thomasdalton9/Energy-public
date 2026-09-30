"""
Bolivia generation and demand from CNDC (Comite Nacional de Despacho de
Carga) - free, public, no key. Found via BOLIVIA_CNDC_PLAYWRIGHT_DISCOVERY.py
+ BOLIVIA_CNDC_API_INSPECT.py + BOLIVIA_XLSX_INSPECT.py: the dashboard's own
WordPress REST API, and a monthly XLSX statistics archive.

Sources:
  https://www.cndc.bo/wp-json/cndc/v1/historico/generacion?desde=Y1&hasta=Y2
    Annual national generation by source (Hidraulica/Eolica/Solar/
    Termoelectrica), MWh, back to 1990 (Eolica/Solar null before each
    existed).
  https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos?categoria_id=155&agrupado=true
    Lists every monthly statistics XLSX CNDC has published, grouped by
    month, with a direct archivo_url per document. This script uses two
    of those document types:
      "Generacion Bruta y Demanda Max. Instantanea" (gen_dia_*.xlsx):
        daily gross generation per PLANT for that month, MWh.
      "Retiros de Energia en Nodos del STI" (ret_dia_*.xlsx):
        daily energy withdrawals per DISTRIBUTION COMPANY / large
        consumer for that month, MWh - i.e. demand by company.

Outputs:
  bo_gen_by_plant.csv      cache: one row per day, one column per plant (MWh)
  bo_demand_by_company.csv cache: one row per day, one column per "company [node]" (MWh)
  bolivia_generation.xlsx  'Annual by source', 'Daily by plant',
                           'Daily generation by region' (long: date, plant,
                           region, country, mwh), 'Daily by company',
                           'Daily demand by region' (long: date, company,
                           region, country, mwh)

Region columns: CNDC's files carry no department/region field. PLANT_REGION
and COMPANY_REGION below are this script's own mapping, looked up by hand
from public knowledge of where each plant/utility actually is. Only
entries we're reasonably confident about are mapped - anything else is
left region=None with a printed warning rather than guessed, the same
rule used for Argentina's fuel codes and Ecuador's plant table.

Usage: python3 BOLIVIA_CNDC.py [--full]
  --full: re-download every month found in the archive listing, ignoring
  the local cache (slow - only needed once, or to force a full refresh).
"""

print("STARTING", flush=True)

import argparse
import io
import os
import sys
from datetime import date

import openpyxl
import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
COUNTRY = "Bolivia"

HISTORICO_URL = "https://www.cndc.bo/wp-json/cndc/v1/historico/generacion"
DOCUMENTOS_URL = "https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos"
DOCUMENTOS_CATEGORIA = 155  # "Mensual" parent category

GEN_TIPO_DOC = "Generación Bruta y Demanda Máx. Instantánea"
RET_TIPO_DOC = "Retiros de Energía en Nodos del STI"

GEN_CACHE = "bo_gen_by_plant.csv"
DEMAND_CACHE = "bo_demand_by_company.csv"
OUT_FILE = "bolivia_generation.xlsx"

# Bolivia's 9 departments. Mapped only where we have reasonable public-
# knowledge confidence in the plant's actual location; CNDC's own plant
# names sometimes carry a unit/node suffix (e.g. "ARANJUEZ" has DF/MG/TG
# units) which this script folds back to the base plant name for mapping.
PLANT_REGION = {
    "ZONGO": "La Paz", "MIGUILLAS": "La Paz", "TAQUESI": "La Paz", "KANATA": "La Paz",
    "EL ALTO": "La Paz", "EASBA": "La Paz",
    "CORANI": "Cochabamba", "S. ISABEL": "Cochabamba", "MISICUNI": "Cochabamba",
    "QOLL I": "Cochabamba", "QOLL II": "Cochabamba", "BULO BULO": "Cochabamba",
    "CARRASCO": "Cochabamba", "ENTRE RIOS": "Cochabamba",
    "GUARACACHI": "Santa Cruz", "SCZ": "Santa Cruz", "WARNES": "Santa Cruz",
    "EWARNES": "Santa Cruz", "GUABIRA": "Santa Cruz", "UNAGRO": "Santa Cruz", "SUR": "Santa Cruz",
    "ORURO I": "Oruro", "ORURO II": "Oruro",
    "UYUNI": "Potosí", "UYUNI II": "Potosí",
    "SAN JACINTO": "Tarija", "YUNCHARA": "Tarija",
    "SAN JOSE I": "Chuquisaca", "SAN JOSE II": "Chuquisaca",
    "MOXOS": "Beni",
}

# Distribution companies/large consumers from RET DIA - each is tied to a
# single department by its service area. Smaller industrial consumers
# we couldn't confidently place are left unmapped.
COMPANY_REGION = {
    "CRE": "Santa Cruz", "DELAPAZ": "La Paz", "ELFEC": "Cochabamba",
    "ENDE DEORURO": "Oruro", "SEPSA": "Potosí", "CESSA": "Chuquisaca",
    "SETAR (TARIJA)": "Tarija", "SETAR (VILLAMONTES)": "Tarija",
    "SETAR (YACUIBA)": "Tarija", "SETAR (BERMEJO)": "Tarija",
    "ENDE DELBENI": "Beni", "EMVINTO": "Oruro", "COBOCE": "Cochabamba",
    "YLB": "Potosí", "PIL ANDINA S.A.": "Cochabamba", "EMDEECRUZ": "Santa Cruz",
}


def fetch_annual_generation():
    r = requests.get(HISTORICO_URL, headers=HEADERS, timeout=TIMEOUT,
                      params={"desde": 1990, "hasta": date.today().year})
    r.raise_for_status()
    data = r.json()
    rows = []
    for series in data.get("series", []):
        name = series["name"]
        for year_str, value in zip(data["categorias"], series["data"]):
            if value is not None:
                rows.append({"year": int(year_str), "source": name, "country": COUNTRY, "mwh": value})
    return pd.DataFrame(rows)


def fetch_documentos():
    r = requests.get(DOCUMENTOS_URL, headers=HEADERS, timeout=TIMEOUT,
                      params={"categoria_id": DOCUMENTOS_CATEGORIA, "agrupado": "true"})
    r.raise_for_status()
    return r.json().get("grupos", [])


def _base_name(label):
    """'ENTRE RIOS' stays as-is; strip-a trailing unit code isn't needed
    since CNDC's own CENTRAL row already repeats the plant name per unit -
    this just uppercases/strips for matching against PLANT_REGION/COMPANY_REGION."""
    return (label or "").strip()


def parse_gen_dia(xlsx_bytes):
    """GEN DIA sheet: header row starts with 'CENTRAL', columns are plant
    names (some repeated with a unit sub-code on the next row), then
    TOTAL, then P. MAXIMA (power, not energy - dropped). Returns a wide
    DataFrame indexed by date, one column per plant (unit codes appended
    where CNDC repeats a plant name)."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = next((i for i, row in enumerate(rows) if row and row[0] == "CENTRAL"), None)
    if header_idx is None:
        raise RuntimeError("GEN DIA: 'CENTRAL' header row not found")
    header = list(rows[header_idx])
    subheader = list(rows[header_idx + 1]) if header_idx + 1 < len(rows) else [None] * len(header)
    seen = {}
    columns = []
    for name, sub in zip(header[1:-2], subheader[1:-2]):  # drop label col, TOTAL, P.MAXIMA
        name = _base_name(name)
        seen[name] = seen.get(name, 0) + 1
        columns.append(f"{name} {sub}" if sub and seen[name] > 1 else name)
    records = {}
    for row in rows[header_idx + 2:]:
        if not row or row[0] is None or not hasattr(row[0], "year"):
            continue
        dt = pd.Timestamp(row[0])
        records[dt] = {col: val for col, val in zip(columns, row[1:-2]) if val is not None}
    return pd.DataFrame.from_dict(records, orient="index").sort_index()


def parse_ret_dia(xlsx_bytes):
    """RET DIA sheet: header row starts with 'CENTRAL' (company names,
    repeated per node), next row has NODO codes. Columns are named
    'COMPANY [NODE]' for uniqueness since a company can span many nodes."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = next((i for i, row in enumerate(rows) if row and row[0] == "CENTRAL"), None)
    if header_idx is None:
        raise RuntimeError("RET DIA: 'CENTRAL' header row not found")
    header = list(rows[header_idx])
    subheader = list(rows[header_idx + 1]) if header_idx + 1 < len(rows) else [None] * len(header)
    columns = [f"{_base_name(name)} [{node}]" for name, node in zip(header[1:-1], subheader[1:-1])]
    records = {}
    for row in rows[header_idx + 2:]:
        if not row or row[0] is None or not hasattr(row[0], "year"):
            continue
        dt = pd.Timestamp(row[0])
        records[dt] = {col: val for col, val in zip(columns, row[1:-1]) if val is not None}
    return pd.DataFrame.from_dict(records, orient="index").sort_index()


def upsert(path, frame):
    if frame.empty:
        return pd.read_csv(path, index_col=0, parse_dates=True) if os.path.exists(path) else frame
    if os.path.exists(path):
        old = pd.read_csv(path, index_col=0, parse_dates=True)
        frame = frame.combine_first(old)
    frame = frame.sort_index()
    frame.to_csv(path)
    return frame


def company_of(column):
    return column.rsplit(" [", 1)[0]


def build_region_long(wide, region_map, entity_col):
    rows = []
    unmapped = set()
    for col in wide.columns:
        entity = company_of(col) if entity_col == "company" else col
        region = region_map.get(entity)
        if region is None:
            unmapped.add(entity)
        for dt, val in wide[col].items():
            if pd.notna(val):
                rows.append({"date": dt, entity_col: entity, "region": region, "country": COUNTRY, "mwh": val})
    if unmapped:
        print(f"  {entity_col}: {len(unmapped)} unmapped (region=None): {sorted(unmapped)}", flush=True)
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--full", action="store_true", help="ignore cache, re-download every month")
    parser.add_argument("--limit", type=int, default=None, help="only process the N most recent months (testing)")
    args = parser.parse_args()

    print("Annual generation by source (historico/generacion)...", flush=True)
    annual = fetch_annual_generation()
    print(f"  {len(annual)} year x source rows, {annual['year'].min()}-{annual['year'].max()}", flush=True)

    print("Monthly document listing...", flush=True)
    grupos = fetch_documentos()
    print(f"  {len(grupos)} months listed", flush=True)
    if args.limit:
        grupos = grupos[:args.limit]
        print(f"  --limit {args.limit}: only processing {len(grupos)} most recent months", flush=True)

    existing_gen = pd.read_csv(GEN_CACHE, index_col=0, parse_dates=True) if (not args.full and os.path.exists(GEN_CACHE)) else pd.DataFrame()
    existing_demand = pd.read_csv(DEMAND_CACHE, index_col=0, parse_dates=True) if (not args.full and os.path.exists(DEMAND_CACHE)) else pd.DataFrame()
    known_months = set()
    if not existing_gen.empty:
        known_months = {(d.year, d.month) for d in existing_gen.index}

    gen_frames, demand_frames = [], []
    session = requests.Session()
    session.headers.update(HEADERS)
    for i, grupo in enumerate(grupos):
        year, month = grupo.get("año"), int(grupo["periodo"].split("-")[1])
        is_recent = i < 2  # always re-fetch the most recent couple of months (late corrections)
        if (year, month) in known_months and not is_recent and not args.full:
            continue
        docs = {d["tipo_doc"]: d["archivo_url"] for d in grupo.get("docs", [])}
        if GEN_TIPO_DOC in docs:
            try:
                r = session.get(docs[GEN_TIPO_DOC], timeout=TIMEOUT)
                r.raise_for_status()
                gen_frames.append(parse_gen_dia(r.content))
            except Exception as e:
                print(f"  {grupo['periodo']} gen: FAILED ({type(e).__name__}: {e})", flush=True)
        if RET_TIPO_DOC in docs:
            try:
                r = session.get(docs[RET_TIPO_DOC], timeout=TIMEOUT)
                r.raise_for_status()
                demand_frames.append(parse_ret_dia(r.content))
            except Exception as e:
                print(f"  {grupo['periodo']} demand: FAILED ({type(e).__name__}: {e})", flush=True)
        print(f"  {grupo['periodo_fmt']}: fetched", flush=True)

    gen_wide = upsert(GEN_CACHE, pd.concat(gen_frames) if gen_frames else pd.DataFrame())
    demand_wide = upsert(DEMAND_CACHE, pd.concat(demand_frames) if demand_frames else pd.DataFrame())
    print(f"Generation by plant: {len(gen_wide)} days, {len(gen_wide.columns)} plants", flush=True)
    print(f"Demand by company: {len(demand_wide)} days, {len(demand_wide.columns)} company/node columns", flush=True)

    print("Mapping to regions...", flush=True)
    gen_by_region = build_region_long(gen_wide, PLANT_REGION, "plant")
    demand_company_totals = pd.DataFrame({company_of(c): demand_wide[c] for c in demand_wide.columns}).T.groupby(level=0).sum().T \
        if not demand_wide.empty else pd.DataFrame()
    demand_by_region = build_region_long(demand_company_totals, COMPANY_REGION, "company")

    notes = [
        "UNITS",
        "All MWh. 'Annual by source': national annual generation by source, from CNDC's historico/generacion "
        "series. 'Daily by plant'/'Daily by company': raw daily values as CNDC publishes them (one plant or "
        "company/node column each). 'Daily generation by region'/'Daily demand by region': the same data reshaped "
        "long with a region column added.",
        "",
        "SOURCES",
        "CNDC (Comite Nacional de Despacho de Carga, www.cndc.bo) - Bolivia's grid operator. Generation: "
        "'Generacion Bruta y Demanda Max. Instantanea' monthly statistics (gross generation per plant). Demand: "
        "'Retiros de Energia en Nodos del STI' monthly statistics (energy withdrawals per distribution company / "
        "large consumer, i.e. demand). Annual by-source history: CNDC's own historico/generacion API.",
        "",
        "REGIONS",
        "CNDC's files carry no department field. PLANT_REGION and COMPANY_REGION in this script are its own "
        "mapping from plant/company name to Bolivian department, looked up by hand from public knowledge. Only "
        "entries mapped with reasonable confidence are included; anything else is left region=None (printed as "
        "'unmapped' when the script runs) rather than guessed.",
        "",
        "SOURCE",
        "cndc.bo - wp-json/cndc/v1/historico/generacion and wp-json/cndc/v1/estadisticas/documentos (monthly "
        "XLSX archive), both discovered via the public dashboard's own API (no key required).",
    ]
    sheets = {"Annual by source": annual, "Daily by plant": gen_wide, "Daily by company": demand_wide}
    if not gen_by_region.empty:
        sheets["Daily generation by region"] = gen_by_region
    if not demand_by_region.empty:
        sheets["Daily demand by region"] = demand_by_region
    xlsx_notes.write_workbook(OUT_FILE, sheets, notes, {"UNITS", "SOURCES", "REGIONS", "SOURCE"})
    print(f"Saved {OUT_FILE}", flush=True)


if __name__ == "__main__":
    main()
