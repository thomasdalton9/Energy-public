"""
Ireland daily gas demand and supply, continuous to the latest day:
GNI open data (data.gov.ie, 2018 .. 2026-03-31) spliced with GNI's
data-transparency pages (2026-03-31 on), which the open data stops at.

Reads (doesn't refetch) the three workbooks the scheduled pulls write:
  ireland_gas_demand_daily.xlsx       (open data: NDM, DM, LDM ex power, power, total ROI)
  ireland_gas_supply_daily.xlsx       (open data: Corrib, Inch, Moffat imports, total)
  ireland_gni_transparency_daily.xlsx (Entry flows + Consumption by market sector)

Reconciliation of the transparency data to the open-data definitions
(checked on the overlap day, 2026-03-31):
  - Moffat in the transparency entry flows is the whole interconnector,
    including gas transiting to Northern Ireland. NI offtake =
    Total_LDM - ROI_LDM; ROI total supply = Aggregate - NI offtake.
  - GNI's Moffat column has blank/partial days while Aggregate is
    complete, so ROI imports = ROI total supply - Corrib (Bellanaboy)
    - Inch - Gormanston.
  - ROI_LDM includes power generation; LDM ex power = ROI_LDM - ROI_Power_Gen.
  - Total ROI demand = NDM + DM + LDM ex power + power.
Each row's Source column says which dataset it came from.

Usage: python3 ireland_gas_combined.py [--out "output/Data and Chart Outputs/ireland_gas_combined_daily.xlsx"]
"""
import argparse
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
SPLICE = pd.Timestamp("2026-03-31")


def load(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    dc = next(c for c in df.columns if "date" in str(c).lower())
    df[dc] = pd.to_datetime(df[dc])
    return df.set_index(dc).sort_index()


def pick(df, *keywords):
    for c in df.columns:
        if all(k in str(c).lower() for k in keywords):
            return c
    raise KeyError(f"no column matching {keywords} in {list(df.columns)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "ireland_gas_combined_daily.xlsx"))
    args = ap.parse_args()
    d = args.data_dir

    open_demand = load(os.path.join(d, "ireland_gas_demand_daily.xlsx"), "Data")
    open_supply = load(os.path.join(d, "ireland_gas_supply_daily.xlsx"), "Data")
    tr = os.path.join(d, "ireland_gni_transparency_daily.xlsx")
    entry, cons = load(tr, "Entry flows"), load(tr, "Consumption by sector")

    # ---- supply
    ni_offtake = (cons[pick(cons, "total_ldm")] - cons[pick(cons, "roi_ldm")]).reindex(entry.index)
    inch = [c for c in entry.columns if "inch" in str(c).lower()]
    gorm = [c for c in entry.columns if "gormanston" in str(c).lower()]
    s_new = pd.DataFrame({
        "Corrib_Production_GWh": entry[pick(entry, "bellanaboy")],
        "Inch_Production_GWh": entry[inch[0]] if inch else 0.0,
        "Total_Supply_GWh": entry[pick(entry, "aggregate")] - ni_offtake,
    })
    s_new["Moffat_Imports_GWh"] = (s_new["Total_Supply_GWh"] - s_new["Corrib_Production_GWh"]
                                   - s_new["Inch_Production_GWh"] - (entry[gorm[0]].fillna(0) if gorm else 0.0))
    s_new = s_new.dropna(subset=["Total_Supply_GWh"])
    supply = pd.concat([open_supply[open_supply.index <= SPLICE].assign(Source="GNI open data"),
                        s_new[s_new.index > SPLICE].assign(Source="GNI transparency")]).sort_index()
    supply = supply[["Corrib_Production_GWh", "Inch_Production_GWh", "Moffat_Imports_GWh", "Total_Supply_GWh", "Source"]]

    # ---- demand
    pg = cons[pick(cons, "power")]
    d_new = pd.DataFrame({
        "NDM_GWh": cons[pick(cons, "ndm")],
        "DM_GWh": cons[[c for c in cons.columns if str(c).lower().startswith("dm")][0]],
        "LDM_ex_PowerGen_GWh": cons[pick(cons, "roi_ldm")] - pg,
        "PowerGen_GWh": pg,
    })
    d_new["Total_ROI_GWh"] = d_new.sum(axis=1, min_count=4)
    d_new = d_new.dropna(subset=["Total_ROI_GWh"])
    demand = pd.concat([open_demand[open_demand.index <= SPLICE].assign(Source="GNI open data"),
                        d_new[d_new.index > SPLICE].assign(Source="GNI transparency")]).sort_index()
    demand = demand[["NDM_GWh", "DM_GWh", "LDM_ex_PowerGen_GWh", "PowerGen_GWh", "Total_ROI_GWh", "Source"]]

    # overlap check on the splice day
    for name, a, b, col in (("supply", open_supply, s_new, "Total_Supply_GWh"),
                            ("demand", open_demand, d_new, "Total_ROI_GWh")):
        if SPLICE in a.index and SPLICE in b.index:
            print(f"{name} on {SPLICE.date()}: open data {a.loc[SPLICE, col]:.1f} vs transparency {b.loc[SPLICE, col]:.1f} GWh")
    for name, df in (("demand", demand), ("supply", supply)):
        full = pd.date_range(df.index.min(), df.index.max(), freq="D")
        print(f"{name}: {df.index.min():%Y-%m-%d}..{df.index.max():%Y-%m-%d}, missing days {len(full.difference(df.index))}")

    notes = [
        "UNITS",
        "GWh per gas day (energy), as published by Gas Networks Ireland (GNI). Republic of Ireland (ROI) only.",
        "",
        "SOURCES",
        f"Up to {SPLICE.date()}: GNI open data on data.gov.ie (ireland_gas_demand_daily.xlsx / ireland_gas_supply_daily.xlsx). "
        f"After {SPLICE.date()}: GNI data-transparency pages (ireland_gni_transparency_daily.xlsx: entry flows, "
        "consumption by market sector). The Source column marks each row.",
        "",
        "RECONCILIATION",
        "Transparency Moffat includes gas in transit to Northern Ireland; NI offtake = Total_LDM - ROI_LDM and ROI total "
        "supply = Aggregate - NI offtake. Moffat imports are derived as ROI total supply minus Corrib (Bellanaboy), Inch "
        "and Gormanston, because GNI's own Moffat column has blank days. LDM ex power = ROI_LDM - ROI_Power_Gen. "
        "These match the open-data definitions on the overlap day.",
        "",
        "SECTORS",
        "NDM: non-daily metered (homes, small business). DM: daily metered (mid-size industry/commercial). "
        "LDM_ex_PowerGen: large daily metered sites excluding power stations. PowerGen: gas-fired power stations.",
    ]
    out_d = demand.copy()
    out_d.index.name = "date"
    out_s = supply.copy()
    out_s.index.name = "date"
    xlsx_notes.write_workbook(args.out, {"Demand": out_d, "Supply": out_s}, notes,
                              {"UNITS", "SOURCES", "RECONCILIATION", "SECTORS"})
    print(f"Saved {args.out}")


if __name__ == "__main__":
    main()
