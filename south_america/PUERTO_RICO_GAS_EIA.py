"""
Puerto Rico natural gas use, monthly, by plant - all Puerto Rico gas is
imported LNG burned for power, so gas burned per plant is the island's gas
use (and, grouped by terminal, its LNG send-out).

Source: EIA (US Energy Information Administration) API v2,
electricity/facility-fuel (EIA-923 plant-level fuel receipts and
consumption), state PR, fuel NG, monthly, total-consumption in Mcf
(thousand cubic feet) and MMBtu:
  https://api.eia.gov/v2/electricity/facility-fuel/data/
Needs the EIA_API_KEY secret.

EIA does not publish Puerto Rico LNG imports (the island is outside the US
customs/import series) and there is no other official monthly import series;
gas burned at the plants is the closest raw figure. Plants are grouped by
the LNG terminal that supplies them:
  Penuelas (EcoEléctrica LNG terminal): EcoEléctrica CCGT, Costa Sur 5&6
  San Juan (New Fortress Energy terminal): San Juan 5&6, the Palo Seco and
            other temporary TM2500 units, and any other plant burning gas
Industrial / commercial gas use outside power plants is small and is not
published monthly.

Units: original unit Mcf per month is kept; mcm/d = Mcf x 0.0283168 / 1000 /
days in month. Incremental: fetches from six months before the last saved
month (EIA revises recent months) and merges; the first run starts 2021-01.

Usage: python3 PUERTO_RICO_GAS_EIA.py [--out PATH] [--full]
"""

print("STARTING", flush=True)

import argparse
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/puerto_rico_gas.xlsx"
URL = "https://api.eia.gov/v2/electricity/facility-fuel/data/"
API_KEY = os.environ.get("EIA_API_KEY")
START = "2021-01"
REFRESH_MONTHS = 6
M3_PER_MCF = 28.3168466
PENUELAS = re.compile(r"eco\s*el|costa\s*sur", re.I)
# Gas burned by use: the four PREPA / Genera PR and IPP gas power stations, other utility units
# (temporary TM2500s etc.), and everything else = industrial combined heat & power (pharma plants)
GROUPS = ["Power_EcoElectrica", "Power_Costa_Sur", "Power_San_Juan", "Power_Palo_Seco_and_other_utility",
          "Industrial_CHP"]
UTILITY = re.compile(r"palo seco|aguirre|mayag|cambalache|daguao|vega baja|jobos|yabucoa|culebra|vieques|"
                     r"tm ?2500|temporary|genera|prepa|peaker", re.I)


def group_of(plant):
    if re.search(r"eco\s*el", plant, re.I):
        return "Power_EcoElectrica"
    if re.search(r"costa\s*sur", plant, re.I):
        return "Power_Costa_Sur"
    if re.search(r"san\s*juan", plant, re.I):
        return "Power_San_Juan"
    if UTILITY.search(plant):
        return "Power_Palo_Seco_and_other_utility"
    return "Industrial_CHP"


def fetch(start):
    rows, offset = [], 0
    while True:
        r = requests.get(URL, params={
            "api_key": API_KEY, "frequency": "monthly", "data[]": ["total-consumption", "total-consumption-btu",
                                                                   "generation"],
            "facets[state][]": "PR", "facets[fuel2002][]": "NG", "facets[primeMover][]": "ALL", "start": start,
            "sort[0][column]": "period", "sort[0][direction]": "asc", "offset": offset, "length": 5000},
            timeout=(10, 90))
        r.raise_for_status()
        resp = r.json()["response"]
        rows += resp["data"]
        offset += len(resp["data"])
        if not resp["data"] or offset >= int(resp["total"]):
            return pd.DataFrame(rows)


def plant_label(name):
    n = re.sub(r"\s+Plant$", "", str(name).strip())
    return n.replace("EcoElectrica", "EcoEléctrica")


def build(raw):
    raw = raw.copy()
    for c in ("total-consumption", "total-consumption-btu", "generation"):
        raw[c] = pd.to_numeric(raw.get(c), errors="coerce")
    raw["plant"] = raw["plantName"].map(plant_label)
    raw["date"] = pd.to_datetime(raw["period"])
    mcf = raw.pivot_table(index="date", columns="plant", values="total-consumption", aggfunc="sum")
    mmbtu = raw.pivot_table(index="date", columns="plant", values="total-consumption-btu", aggfunc="sum")
    mcf = mcf.loc[:, mcf.fillna(0).abs().sum() > 0]
    groups = {g: [p for p in mcf.columns if group_of(p) == g] for g in GROUPS}
    days = mcf.index.days_in_month
    to_mcmd = lambda s: s * M3_PER_MCF / 1e6 / days  # noqa: E731
    use = pd.DataFrame(index=mcf.index)
    for g in GROUPS:
        use[f"{g}_mcm_per_day"] = to_mcmd(mcf[groups[g]].sum(axis=1, min_count=1)) if groups[g] else 0.0
    pen = [p for p in mcf.columns if PENUELAS.search(p)]
    use["Penuelas_terminal_mcm_per_day"] = to_mcmd(mcf[pen].sum(axis=1, min_count=1))
    use["San_Juan_terminal_mcm_per_day"] = to_mcmd(mcf[[p for p in mcf.columns if p not in pen]].sum(axis=1, min_count=1))
    use["Total_mcm_per_day"] = to_mcmd(mcf.sum(axis=1, min_count=1))
    use["Total_Mcf_per_month"] = mcf.sum(axis=1, min_count=1)
    use["Total_MMBtu_per_month"] = mmbtu.reindex(index=mcf.index).sum(axis=1, min_count=1)
    for g in GROUPS:
        use[f"{g}_Mcf_per_month"] = mcf[groups[g]].sum(axis=1, min_count=1) if groups[g] else 0.0
    for p in mcf.columns:
        use[f"Plant {p}_Mcf_per_month"] = mcf[p]
    use.index.name = "Month"
    return use.round(4)


def load(path):
    try:
        d = pd.read_excel(path, sheet_name="Gas use", index_col=0)
        d.index = pd.to_datetime(d.index)
        return d
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def notes(d):
    return [
        "UNITS",
        "*_mcm_per_day: million cubic metres per day (monthly average) = Mcf x 0.0283168 / 1000 / days in month.",
        "*_Mcf_per_month: thousand cubic feet in the month, EIA's original unit. Total_MMBtu_per_month: heat content.",
        "Month: first day of the month. MONTHLY data - the finest granularity EIA publishes for Puerto Rico.",
        "",
        "COVERAGE",
        f"{d.index.min():%b-%Y} to {d.index.max():%b-%Y} ({len(d)} months). EIA-923 is published about two months "
        "after month end.",
        "Covers gas burned in every EIA-923 plant: the power stations (almost all Puerto Rico gas use) and the "
        "industrial combined heat & power plants (pharmaceutical sites) that report to EIA. Puerto Rico LNG imports "
        "are not published by EIA (or anyone official, monthly); the terminal groups below are the send-out implied "
        "by plant burn.",
        "",
        "SOURCE",
        "EIA API v2 electricity/facility-fuel (EIA-923 plant fuel consumption), state PR, fuel NG: " + URL,
        "Browse: https://www.eia.gov/electricity/data/browser/ (Plant level data, Puerto Rico)",
        "",
        "MAPPING",
        "Power_EcoElectrica: EcoEléctrica CCGT (independent producer, Penuelas). Power_Costa_Sur: Costa Sur 5&6. "
        "Power_San_Juan: Central San Juan 5&6. Power_Palo_Seco_and_other_utility: Palo Seco temporary TM2500 units "
        "and any other PREPA / Genera PR site. Industrial_CHP: every other plant (AbbVie, Baxter, Lilly, Stryker, "
        "HP, Neolpharma, ...). 'Plant <name>' columns: each plant's own Mcf per month.",
        "Penuelas_terminal = EcoEléctrica + Costa Sur (both fed from the EcoEléctrica LNG terminal at Penuelas). "
        "San_Juan_terminal = every other plant (New Fortress Energy's San Juan terminal, by pipeline or trucked "
        "ISO containers).",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--full", action="store_true")
    args = ap.parse_args()
    if not API_KEY:
        sys.exit("EIA_API_KEY is not set")
    old = load(args.out)
    if not old.empty and "Industrial_CHP_mcm_per_day" not in old.columns:
        old = pd.DataFrame()   # saved with an older column layout: rebuild from 2021
    start = START
    if not args.full and not old.empty:
        start = max(pd.Timestamp(START), old.index.max() - pd.DateOffset(months=REFRESH_MONTHS)).strftime("%Y-%m")
    print(f"Fetching EIA-923 Puerto Rico plant gas use from {start}", flush=True)
    raw = fetch(start)
    if raw.empty:
        print("EIA returned nothing", flush=True)
        if old.empty:
            sys.exit(1)
        return
    new = build(raw)
    d = new if old.empty else pd.concat([old[~old.index.isin(new.index)], new]).sort_index()
    d = d[d.index >= pd.Timestamp(START)]
    lead = [c for c in d.columns if c.endswith("_mcm_per_day")]
    plants = sorted(c for c in d.columns if c.startswith("Plant "))
    d = d[[*lead, *[c for c in d.columns if c not in lead and c not in plants], *plants]]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Gas use": d}, notes(d), {"UNITS", "COVERAGE", "SOURCE", "MAPPING"})
    print(f"Saved {args.out}: {len(d)} months")
    print(d[lead].tail(6).round(2).to_string())


if __name__ == "__main__":
    main()
