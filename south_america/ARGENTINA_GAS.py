"""
Argentina national gas production and consumption by sector and by
distributor, from the Ministry of Economy's Series de Tiempo API
(apis.datos.gob.ar/series/api/series/) - free, public, no key.

Found via ARGENTINA_GAS_DISCOVERY.py + ARGENTINA_GAS_DISCOVERY2.py: all
these series belong to one dataset ("Producción y consumo de gas
natural", Secretaría de Energía / Ministerio de Economía), monthly
(R/P1M), 1996-01 to the latest available month. Units as published:
millones de metros cúbicos (million m3) per month.

Outputs (argentina_gas.xlsx):
  National              date, produccion_gas_natural, + 6 use-sector
                         columns (residencial, comercial, entes_oficiales,
                         industria, centrales_electricas [gas burned for
                         power generation], gnc [compressed gas for
                         vehicles]), country
  By sector (long)       the 6 use-sector columns reshaped long
  By distributor         date, one column per gas distribution company,
                         country
  By region (long)       the distributor columns reshaped long with a
                         region column added (REGION_MAP below)

REGION_MAP covers the 9 classic distribution licensees from Argentina's
1992 gas-distribution privatization, each with a single well-defined
service territory. "sdb" and "redengas" are smaller/newer entities this
script couldn't confidently place in one province - left region=None
(printed as unmapped) rather than guessed. "gnc" (compressed gas for
vehicles) is a use-category, not a regional distributor, so it's not
in the distributor table at all.

Usage: python3 ARGENTINA_GAS.py [--out argentina_gas.xlsx]
"""
print("STARTING", flush=True)

import argparse
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)
SERIES_API = "https://apis.datos.gob.ar/series/api/series/"
COUNTRY = "Argentina"
OUT_DEFAULT = "argentina_gas.xlsx"

# column name -> confirmed monthly (R/P1M) series id, from
# ARGENTINA_GAS_DISCOVERY2.py's live output
PRODUCTION_ID = "364.3_PRODUCCIoNRAL__25"

SECTOR_IDS = {
    "residencial": "364.3_RESIDENCIAIAL__11",
    "comercial": "364.3_COMERCIALIAL__9",
    "entes_oficiales": "364.3_ENTES_OFICLES__15",
    "industria": "364.3_INDUSTRIARIA__9",
    "centrales_electricas": "364.3_CENTRALES_CAS__20",
    "gnc": "364.3_GNCGNC__3",
}

DISTRIBUTOR_IDS = {
    "metrogas": "364.3_METROGASGAS__8",
    "gas_natural_fenosa": "364.3_GAS_NATURAOSA__18",
    "distrib_gas_del_centro_ecogas": "364.3_DISTRIB._GGAS__30",
    "distrib_gas_cuyana_ecogas": "364.3_DISTRIB._GGAS__26",
    "litoral_gas": "364.3_LITORAL_GAGAS__11",
    "gasnea": "364.3_GASNEANEA__6",
    "gasnor": "364.3_GASNORNOR__6",
    "camuzzi_gas_pampeana": "364.3_CAMUZZI_GAANA__20",
    "camuzzi_gas_del_sur": "364.3_CAMUZZI_GASUR__19",
    "sdb": "364.3_SDBSDB__3",
    "redengas": "364.3_REDENGASGAS__8",
}

# The 9 classic 1992-privatization distribution licensees, each with one
# well-defined service territory. sdb/redengas deliberately left out -
# not confidently placeable in a single province.
REGION_MAP = {
    "metrogas": "CABA + Norte GBA",
    "gas_natural_fenosa": "Oeste/Sur GBA",
    "camuzzi_gas_pampeana": "Buenos Aires (provincia) + La Pampa",
    "camuzzi_gas_del_sur": "Patagonia",
    "litoral_gas": "Santa Fe + Entre Ríos",
    "distrib_gas_del_centro_ecogas": "Córdoba",
    "distrib_gas_cuyana_ecogas": "Cuyo (Mendoza/San Juan/San Luis)",
    "gasnor": "NOA",
    "gasnea": "NEA",
}


def fetch_series(ids):
    """ids: {column_name: series_id}. Returns a wide DataFrame indexed by
    date, one column per name, values in the series' native units
    (million m3/month)."""
    id_list = list(ids.values())
    r = requests.get(SERIES_API, headers=HEADERS, timeout=TIMEOUT,
                      params={"ids": ",".join(id_list), "format": "json", "limit": 5000})
    r.raise_for_status()
    payload = r.json()
    dates = [row[0] for row in payload["data"]]
    df = pd.DataFrame(payload["data"], columns=["date"] + id_list).set_index("date")
    df.index = pd.to_datetime(df.index)
    id_to_name = {v: k for k, v in ids.items()}
    return df.rename(columns=id_to_name).sort_index()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=OUT_DEFAULT)
    args = parser.parse_args()

    print("Fetching production + sector consumption...", flush=True)
    national = fetch_series({"produccion_gas_natural": PRODUCTION_ID, **SECTOR_IDS})
    print(f"  {len(national)} months, {national.index.min().date()} to {national.index.max().date()}", flush=True)

    print("Fetching distributor consumption...", flush=True)
    distributors = fetch_series(DISTRIBUTOR_IDS)
    print(f"  {len(distributors)} months, {distributors.index.min().date()} to {distributors.index.max().date()}",
          flush=True)

    national_out = national.copy()
    national_out.insert(0, "country", COUNTRY)

    sector_long = national[list(SECTOR_IDS)].reset_index().melt(
        id_vars="date", var_name="sector", value_name="mmc")
    sector_long["country"] = COUNTRY
    sector_long = sector_long.dropna(subset=["mmc"])

    distributors_out = distributors.copy()
    distributors_out.insert(0, "country", COUNTRY)

    region_long = distributors.reset_index().melt(id_vars="date", var_name="distributor", value_name="mmc")
    region_long["region"] = region_long["distributor"].map(REGION_MAP)
    region_long["country"] = COUNTRY
    region_long = region_long.dropna(subset=["mmc"])
    unmapped = sorted(set(region_long.loc[region_long["region"].isna(), "distributor"]))
    if unmapped:
        print(f"  distributors: {len(unmapped)} unmapped (region=None): {unmapped}", flush=True)

    notes = [
        "UNITS",
        "All values in millones de metros cubicos (million m3) per month, as published.",
        "",
        "SECTORS",
        "'National': produccion_gas_natural (national gas production) plus 6 use-sector consumption columns - "
        "residencial, comercial, entes_oficiales (official/government bodies), industria, centrales_electricas "
        "(gas burned for power generation), gnc (compressed gas for vehicles). 'By sector (long)' reshapes the "
        "6 sector columns long.",
        "",
        "REGIONS",
        "'By distributor': consumption by each of Argentina's licensed gas distribution companies. 'By region "
        "(long)' adds a region column via REGION_MAP - the 9 classic 1992-privatization distribution licensees, "
        "each with one well-defined service territory. 'sdb' and 'redengas' are smaller/newer entities not "
        "confidently placeable in a single province - left region=None (printed as unmapped) rather than guessed.",
        "",
        "SOURCE",
        "apis.datos.gob.ar/series/api/series - Secretaria de Energia, Ministerio de Economia, dataset "
        "'Produccion y consumo de gas natural'. Monthly, 1996-01 onward.",
    ]
    sheets = {
        "National": national_out,
        "By sector (long)": sector_long,
        "By distributor": distributors_out,
        "By region (long)": region_long,
    }
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "REGIONS", "SOURCE"})
    print(f"Saved {args.out}", flush=True)
    print(national_out.tail().to_string(), flush=True)


if __name__ == "__main__":
    main()
