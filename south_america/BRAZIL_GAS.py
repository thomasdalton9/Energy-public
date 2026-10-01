"""
Brazil monthly natural gas balance - demand by segment and supply by
source - from the Ministry of Mines and Energy's (MME) monthly gas
bulletin annex "historico-balanco-boletim.xlsx" (one sheet, "DADOS
BRASIL", a row per balance line and a column per month). Free, public,
no key. Found via SA_GAS_DEMAND_DISCOVERY2.py / BRAZIL_GAS_BALANCE_INSPECT.py.

Units: million m3/day (monthly average), as published.
Demand segments (as labelled by MME): Industrial (includes refineries
and fertiliser plants, per the sheet's own footnote), Automotivo
(vehicle CNG), Residencial, Comercial, Geracao Eletrica (power),
Cogeracao, Outros (inclui GNC) and DEMANDA TOTAL. The supply block
above it (production, imports from Bolivia, LNG imports, E&P own use,
OFERTA TOTAL) is kept on its own sheet.

The month header row is located by content (the row with the most
date-like cells) rather than a fixed index, and every label is printed
in the run log so a layout change in MME's file is visible.

Outputs (brazil_gas_monthly.xlsx):
  Demand by segment   date x segment columns, from --start-date (2021-01)
  Balance (all rows)  date x every labelled row in the sheet
"""
import argparse
import io
import os
import re
import sys
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
       "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural/anexos/historico-balanco-boletim.xlsx")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 90)
DATA_START = "2021-01-01"
OUT_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "brazil_gas_monthly.xlsx")

SEGMENTS = {  # normalised label prefix -> output column
    "industrial": "Industrial",
    "automotivo": "Automotive",
    "residencial": "Residential",
    "comercial": "Commercial",
    "geracao eletrica": "Power_Generation",
    "cogeracao": "Cogeneration",
    "outros": "Other_incl_CNG",
    "demanda total": "Total_Demand",
}

PT_MONTHS = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10,
             "nov": 11, "dez": 12}


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9 ]+", " ", s).strip()


def to_month(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 30000 < v < 60000:  # Excel serial date
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))).to_period("M").to_timestamp()
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
        try:
            return pd.Timestamp(v).to_period("M").to_timestamp()
        except Exception:
            return None
    m = re.match(r"^\s*([A-Za-zçÇ]{3})[a-zç]*[\s/\-.]*(\d{2,4})\s*$", str(v))
    if m and norm(m.group(1))[:3] in PT_MONTHS:
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        return pd.Timestamp(y, PT_MONTHS[norm(m.group(1))[:3]], 1)
    return None


def fetch():
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    if r.content[:2] != b"PK":
        raise RuntimeError(f"Not an xlsx (content-type {r.headers.get('content-type')}) - MME may have moved the file")
    xl = pd.ExcelFile(io.BytesIO(r.content))
    raw = xl.parse(xl.sheet_names[0], header=None)
    print(f"sheet {xl.sheet_names[0]!r} shape {raw.shape}", flush=True)

    best_row, best_n = None, 0
    for i in range(min(12, len(raw))):
        n = sum(to_month(v) is not None for v in raw.iloc[i, 1:])
        if n > best_n:
            best_row, best_n = i, n
    if best_row is None or best_n < 12:
        raise RuntimeError("Could not find the month header row")
    months = [to_month(v) for v in raw.iloc[best_row, 1:]]
    print(f"header row {best_row}: {best_n} month columns, {min(m for m in months if m)} to {max(m for m in months if m)}",
          flush=True)

    rows = {}
    for i in range(best_row + 1, len(raw)):
        label = raw.iloc[i, 0]
        if pd.isna(label) or not str(label).strip():
            continue
        vals = pd.to_numeric(raw.iloc[i, 1:], errors="coerce")
        if vals.notna().sum() == 0:
            continue
        name = str(label).strip()
        rows[name] = pd.Series(vals.to_numpy(), index=months)
        print(f"  row {i:2d}: {name}", flush=True)
    bal = pd.DataFrame(rows)
    bal = bal[bal.index.notna()]
    bal.index = pd.DatetimeIndex(bal.index, name="date")
    return bal.sort_index()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--start-date", default=DATA_START)
    args = parser.parse_args()

    bal = fetch()
    bal = bal[bal.index >= args.start_date]

    demand = pd.DataFrame(index=bal.index)
    for col in bal.columns:
        n = norm(col)
        for prefix, out in SEGMENTS.items():
            if n.startswith(prefix) and out not in demand.columns:
                demand[out] = bal[col]
    missing = set(SEGMENTS.values()) - set(demand.columns)
    if missing:
        raise RuntimeError(f"segments not found in MME sheet: {missing} - labels were {list(bal.columns)}")
    parts = [c for c in demand.columns if c != "Total_Demand"]
    gap = (demand[parts].sum(axis=1) - demand["Total_Demand"]).abs().max()
    print(f"\nsegments sum vs DEMANDA TOTAL: max abs diff {gap:.3f} million m3/d", flush=True)
    demand = demand.dropna(how="all")
    print(f"demand: {len(demand)} months, {demand.index.min().date()} to {demand.index.max().date()}", flush=True)
    print(demand.tail(6).round(2).to_string(), flush=True)

    notes = [
        "UNITS",
        "Million m3 per day (monthly average), as published by MME.",
        "",
        "SEGMENTS",
        "Industrial (includes refineries and fertiliser plants, per MME's footnote), Automotive (vehicle CNG "
        "sold at the pump), Residential, Commercial, Power_Generation (gas-fired power plants), Cogeneration, "
        "Other_incl_CNG, Total_Demand (MME's own total).",
        "",
        "BALANCE",
        "'Balance (all rows)' keeps every labelled line of MME's sheet, including the supply side (domestic "
        "production, imports from Bolivia, LNG imports, E&P own use, pipeline consumption/imbalance, total "
        "supply), under MME's own Portuguese labels.",
        "",
        "COVERAGE",
        f"Monthly from {args.start_date}; MME's file goes further back - rerun with --start-date to pull more.",
        "",
        "SOURCE",
        f"MME, Boletim Mensal de Acompanhamento da Industria de Gas Natural, annex: {URL}",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Demand by segment": demand, "Balance (all rows)": bal}, notes,
                              {"UNITS", "SEGMENTS", "BALANCE", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
