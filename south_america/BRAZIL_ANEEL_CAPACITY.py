"""
Brazil installed generation capacity by plant, from ANEEL's SIGA
(Sistema de Informacoes de Geracao) registry - free, public, no key.
Found via CAPACITY_DISCOVERY.py + BRAZIL_ANEEL_CAPACITY_INSPECT.py: a
single CSV listing every registered generation enterprise in Brazil
(large + small hydro, thermal, wind, solar, nuclear, biomass ...), with
state, municipality, fuel, operational status and capacity in kW.

Source: https://dadosabertos.aneel.gov.br - dataset
"siga-sistema-de-informacoes-de-geracao-da-aneel", resource
siga-empreendimentos-geracao.csv. ';'-delimited, UTF-8 encoded,
comma-decimal numbers (Brazilian format). (The inspection script that
first looked at this file picked latin-1 because that decodes ANY byte
sequence without erroring - it just silently mojibakes "Operação" into
"OperaÃ§Ã£o" rather than failing loudly, so the encoding must be
confirmed by checking the actual text, not just "did read_csv raise".)

Capacity fields (kW, as published):
  MdaPotenciaOutorgadaKw    licensed/granted capacity (may include
                            capacity not yet built)
  MdaPotenciaFiscalizadaKw  "fiscalized" (verified/audited) capacity -
                            used here as the main capacity figure, since
                            it reflects what's actually in service
  MdaGarantiaFisicaKw       firm "physical guarantee" capacity commitment

Only plants with DscFaseUsina == "Operação" (in commercial operation)
are kept - the registry also lists planned/under-construction/retired
entries, which aren't real installed capacity today.

Outputs (brazil_capacity.xlsx):
  By plant           one row per operating plant: name, state,
                     municipality, type, fuel, capacity_mw, country
  By state and fuel  capacity_mw summed by (state, fuel_category)
  By fuel (national) capacity_mw summed by fuel_category nationally

Usage: python3 BRAZIL_ANEEL_CAPACITY.py [--out brazil_capacity.xlsx]
"""
print("STARTING", flush=True)

import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = ("https://dadosabertos.aneel.gov.br/dataset/6d90b77c-c5f5-4d81-bdec-7bc619494bb9/"
       "resource/11ec447d-698d-4ab8-977f-b424d5deee6a/download/siga-empreendimentos-geracao.csv")
HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 180)
COUNTRY = "Brazil"
OUT_DEFAULT = "brazil_capacity.xlsx"

# SigTipoGeracao (plant-size/technology class) -> broad fuel category.
# Hydro is split into 3 size classes in ANEEL's own scheme (UHE = large,
# PCH = small, CGH = mini) - all rolled up to "hydro" here. Remaining
# codes map straight to the obvious category; DscFonteCombustivel is
# used as a fallback for anything not in this dict (e.g. "UTE" thermal
# plants, where the fuel - gas/coal/oil/biomass - actually varies).
TIPO_TO_CATEGORY = {
    "UHE": "hydro", "PCH": "hydro", "CGH": "hydro",
    "EOL": "wind", "UFV": "solar", "UTN": "nuclear",
}

# DscFonteCombustivel (fuel source description) -> category, used for
# UTE (thermal) plants where the technology code alone doesn't say what
# fuel is burned. Exact strings confirmed from a live pull (ANEEL's own
# capitalization, e.g. "Gás natural" not "Gás Natural") - a first pass
# using guessed capitalization silently fell through to "other (...)"
# for every thermal fuel, since Python string matching is case-exact.
FONTE_TO_CATEGORY = {
    "Gás natural": "gas_thermal", "Gás de Processo": "gas_thermal", "Gás de processo": "gas_thermal",
    "Petróleo": "oil_thermal", "Óleo Combustível": "oil_thermal", "Óleo Diesel": "oil_thermal",
    "Carvão mineral": "coal",
    # Biomass sub-types, all rolled up to one "biomass" category
    "Agroindustriais": "biomass", "Floresta": "biomass", "Resíduos sólidos urbanos": "biomass",
    "Resíduos animais": "biomass", "Biocombustíveis líquidos": "biomass", "Biomassa": "biomass",
    "Potencial hidráulico": "hydro", "Cinética do vento": "wind",
    "Radiação solar": "solar", "Urânio": "nuclear",
}


def category_of(tipo, fonte):
    if tipo in TIPO_TO_CATEGORY:
        return TIPO_TO_CATEGORY[tipo]
    return FONTE_TO_CATEGORY.get(fonte, f"other ({fonte})")


def to_float_br(series):
    """Brazilian number format: comma as decimal separator."""
    return pd.to_numeric(series.astype(str).str.replace(",", ".", regex=False), errors="coerce")


def fetch_raw():
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return pd.read_csv(io.BytesIO(r.content), sep=";", encoding="utf-8", low_memory=False)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=OUT_DEFAULT)
    args = parser.parse_args()

    print("Downloading ANEEL SIGA registry (full national CSV)...", flush=True)
    raw = fetch_raw()
    print(f"  {len(raw):,} total registry rows (all statuses)", flush=True)

    operating = raw[raw["DscFaseUsina"] == "Operação"].copy()
    print(f"  {len(operating):,} rows with DscFaseUsina == 'Operação'", flush=True)

    operating["capacity_mw"] = to_float_br(operating["MdaPotenciaFiscalizadaKw"]) / 1000.0
    operating["category"] = [
        category_of(tipo, fonte) for tipo, fonte in zip(operating["SigTipoGeracao"], operating["DscFonteCombustivel"])
    ]
    operating["country"] = COUNTRY

    by_plant = operating[[
        "NomEmpreendimento", "SigUFPrincipal", "DscMuninicpios", "SigTipoGeracao", "category",
        "NomFonteCombustivel", "DatEntradaOperacao", "capacity_mw", "country",
    ]].rename(columns={
        "NomEmpreendimento": "plant", "SigUFPrincipal": "state", "DscMuninicpios": "municipality",
        "SigTipoGeracao": "plant_type", "NomFonteCombustivel": "fuel", "DatEntradaOperacao": "operation_start",
    }).sort_values(["state", "category", "plant"])

    by_state_fuel = (by_plant.groupby(["state", "category"])["capacity_mw"].sum()
                      .reset_index().sort_values(["state", "category"]))
    by_fuel = (by_plant.groupby("category")["capacity_mw"].sum()
               .reset_index().sort_values("capacity_mw", ascending=False))

    print(f"National total: {by_plant['capacity_mw'].sum():,.0f} MW across {len(by_plant):,} operating plants",
          flush=True)
    print(by_fuel.to_string(index=False), flush=True)

    notes = [
        "UNITS",
        "capacity_mw is MdaPotenciaFiscalizadaKw (ANEEL's 'fiscalized'/verified capacity) / 1000, i.e. the "
        "capacity actually in commercial service - not the licensed ('outorgada') capacity, which can include "
        "capacity not yet built.",
        "",
        "SCOPE",
        "Only rows with DscFaseUsina == 'Operacao' (in commercial operation) are included - ANEEL's registry "
        "also lists planned, under-construction and retired entries, which aren't real installed capacity today.",
        "",
        "CATEGORIES",
        "Hydro (UHE/PCH/CGH - large/small/mini, by ANEEL's own size classes), wind (EOL), solar (UFV), nuclear "
        "(UTN) come straight from the plant-type code. Thermal plants (UTE) are split by DscFonteCombustivel into "
        "gas_thermal, oil_thermal, coal and biomass. Anything not matching either mapping keeps its raw "
        "DscFonteCombustivel value as 'other (...)' rather than being dropped or guessed.",
        "",
        "SOURCE",
        "dadosabertos.aneel.gov.br - dataset 'siga-sistema-de-informacoes-de-geracao-da-aneel', resource "
        "siga-empreendimentos-geracao.csv. Snapshot as of whenever this script last ran - ANEEL republishes the "
        "full file (not incremental), so this isn't a time series of capacity additions.",
    ]
    sheets = {"By plant": by_plant, "By state and fuel": by_state_fuel, "By fuel (national)": by_fuel}
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SCOPE", "CATEGORIES", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
