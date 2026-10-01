"""
Dominican Republic natural gas burned for power, monthly, from the
Superintendencia de Electricidad (SIE).

Source: SIE "Consumo Combustible, 2015 - 2026" - monthly fuel consumed to
produce electricity in the Mercado Electrico Mayorista (SENI), by fuel: Gas
Natural, Fuel Oil No.2, Fuel Oil No.6, Carbon. Published on the government
open-data portal:
  https://datos.gob.do/dataset/energia-y-potencia-facturadas-ede
  (CKAN API: https://datos.gob.do/api/3/action/package_show?id=energia-y-potencia-facturadas-ede
   -> the XLSX resource on sie.gob.do)
The gas column is headed "GAS NATURAL (MMBTU)" but its values are THOUSAND
MMBtu: May-2026 = 7,673 against 888 GWh of gas-fired generation (MEM
bulletin) is an 8.6 MMBtu/MWh heat rate; read as plain MMBtu it would be
0.0086. The script converts with that reading and checks it.

Natural gas in the Dominican Republic is all imported LNG (AES Andres terminal
at Punta Caucedo, plus the Energia Natural del Caribe / Manzanillo terminal);
power is most of its use. Gas sold to industry and vehicles (GNV) is not in
this series and is not published monthly.

Incremental: the file is one small sheet; new months are appended and the
last six saved months refreshed (SIE revises). Data from 2021.

Usage: python3 DOMINICAN_REPUBLIC_GAS_SIE.py [--out PATH]
"""

print("STARTING", flush=True)

import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/dominican_republic_gas.xlsx"
CKAN = "https://datos.gob.do/api/3/action/package_show"
DATASET = "energia-y-potencia-facturadas-ede"
DATASET_PAGE = f"https://datos.gob.do/dataset/{DATASET}"
FALLBACK = "https://sie.gob.do/wp-content/uploads/2025/07/Copia-de-Consumo-Combustible_Mensual_2026-xls-1.xlsx"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
START = pd.Timestamp("2021-01-01")
REFRESH_MONTHS = 6
BTU_PER_CF = 1037                       # EIA average heat content of natural gas
M3_PER_MMBTU = 1e6 / BTU_PER_CF * 0.0283168466   # ~27.3 m3 per MMBtu


def resource_url():
    try:
        r = requests.get(CKAN, params={"id": DATASET}, headers=UA, timeout=60)
        r.raise_for_status()
        res = r.json()["result"]["resources"]
        xl = [x["url"] for x in res if str(x.get("format", "")).upper() == "XLSX"]
        if xl:
            return xl[0]
    except (requests.RequestException, ValueError, KeyError) as e:
        print(f"CKAN lookup failed ({e}); using the last known file", flush=True)
    return FALLBACK


def fetch():
    url = resource_url()
    print(f"Downloading {url}", flush=True)
    r = requests.get(url, headers=UA, timeout=120)
    r.raise_for_status()
    d = pd.read_excel(io.BytesIO(r.content), header=0)
    d = d.rename(columns={d.columns[0]: "month"})
    d["month"] = pd.to_datetime(d["month"], errors="coerce").dt.to_period("M").dt.to_timestamp()
    d = d.dropna(subset=["month"]).drop_duplicates("month", keep="last").set_index("month").sort_index()
    gas = next(c for c in d.columns if "GAS" in str(c).upper())
    other = {c: c for c in d.columns if c != gas}
    return d[gas].astype(float), d[list(other)].astype(float), url


def build(gas_k_mmbtu, others):
    days = gas_k_mmbtu.index.days_in_month
    out = pd.DataFrame(index=gas_k_mmbtu.index)
    out["Power_mcm_per_day"] = (gas_k_mmbtu * 1000 * M3_PER_MMBTU / 1e6 / days).round(4)
    out["Power_thousand_MMBtu_per_month"] = gas_k_mmbtu.round(3)
    out["Power_MMBtu_per_day"] = (gas_k_mmbtu * 1000 / days).round(0)
    for c in others.columns:   # the other generation fuels, as published (context)
        out["SIE " + str(c)] = others[c].round(3)
    out.index.name = "Month"
    return out


def load(path):
    try:
        d = pd.read_excel(path, sheet_name="Gas use", index_col=0)
        d.index = pd.to_datetime(d.index)
        return d
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def notes(d, url):
    return [
        "UNITS",
        "Month: first day of the month (monthly data).",
        "Power_thousand_MMBtu_per_month: SIE's figure (headed 'GAS NATURAL (MMBTU)' in the source, but the values are "
        "thousand MMBtu - see SOURCE).",
        f"Power_mcm_per_day: million m3 per day = thousand MMBtu x 1000 x {M3_PER_MMBTU:.2f} m3/MMBtu (EIA average "
        f"{BTU_PER_CF} Btu per cubic foot) / days in month. Power_MMBtu_per_day: MMBtu per day.",
        "SIE <fuel> columns: the other generation fuels exactly as SIE publishes them (context only).",
        "",
        "COVERAGE",
        f"{d.index.min():%b-%Y} to {d.index.max():%b-%Y} ({len(d)} months). Natural gas burned to generate electricity "
        "for the SENI wholesale market (AES Andres, Los Mina, Estrella del Mar, Energia Natural del Caribe / Manzanillo, "
        "converted engines and steam units). All Dominican gas is imported LNG; gas sold to industry and vehicles (GNV) "
        "is not included and is not published monthly. SIE updates the file every few months.",
        "",
        "SOURCE",
        "Superintendencia de Electricidad (SIE), 'Consumo Combustible, 2015 - 2026' (monthly fuel consumption for "
        f"electricity generation), on the open-data portal {DATASET_PAGE}",
        f"File read this run: {url}",
        "Unit check: May-2026 gas 7,673 vs 888 GWh of gas-fired generation (MEM 'Boletin de Generacion y Gestion de "
        "Energia', May 2026) = 8.6 MMBtu/MWh, a normal heat rate, so the values are thousand MMBtu.",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    old = load(args.out)
    gas, others, url = fetch()
    if not old.empty:   # backfill: months not saved yet, plus the last few SIE may have revised
        keep_from = old.index.max() - pd.DateOffset(months=REFRESH_MONTHS)
        gas = gas[(gas.index > keep_from) | ~gas.index.isin(old.index)]
        others = others.reindex(gas.index)
    new = build(gas[gas.index >= START], others[others.index >= START])
    d = new if old.empty else pd.concat([old[~old.index.isin(new.index)], new]).sort_index()
    d = d[d.index >= START]
    last = d.index.max()
    print(f"{len(new)} months fetched/refreshed; latest {last:%b-%Y}: {d.loc[last, 'Power_mcm_per_day']:.2f} mcm/d",
          flush=True)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Gas use": d}, notes(d, url), {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)
    print(d.iloc[-8:, :3].to_string(), flush=True)


if __name__ == "__main__":
    main()
