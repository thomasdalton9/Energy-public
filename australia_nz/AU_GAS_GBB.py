"""
Australia east coast gas (plus Northern Territory) from AEMO's Gas Bulletin
Board (GBB) - daily facility flows, public, no key:

  https://nemweb.com.au/Reports/Current/GBB/GasBBActualFlowStorage.zip      full history (from Sep 2018)
  https://nemweb.com.au/Reports/Current/GBB/GasBBActualFlowStorageLast31.CSV last 31 gas days
  https://nemweb.com.au/Reports/Current/GBB/GasBBLNGShipments.CSV            LNG export cargoes

Writes au_gas.xlsx:
  Demand by sector  TJ/day: gas-powered generation (BBGPG), large industrial users (BBLARGE), LNG export plants
                    (LNGEXPORT). From 15 Mar 2023, when AEMO extended the Bulletin Board to these facility types.
                    The Bulletin Board has no distribution-network (residential/commercial) facility type, so mass-
                    market demand is not in this table.
  Production        TJ/day by state, production facilities (PROD) supply
  Storage           TJ held in storage by facility (STOR: Iona, Roma, Silver Springs, Newcastle, Dandenong LNG,
                    Moomba) and Total
  LNG shipments     PJ per month by LNG export plant (Curtis Island: APLNG, GLNG, QCLNG)

Incremental: a new file reads the full-history zip once; later runs read only the last-31-days file and
the shipments list, and merge them over the saved rows (recent gas days are revised).

Usage: python3 AU_GAS_GBB.py [--out "output/Data and Chart Outputs/au_gas.xlsx"]
"""
import argparse
import io
import os
import sys
import time
import zipfile

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

GBB = "https://nemweb.com.au/Reports/Current/GBB/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_gas.xlsx")
SECTORS = {"BBGPG": "Gas_power_generation", "BBLARGE": "Large_industrial", "LNGEXPORT": "LNG_export_plants"}
SECTOR_START = "2023-03-15"


def get(url):
    for attempt in range(4):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 300))
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    attempt {attempt + 1}/4 failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(f"failed: {url}")


def flows(full):
    if full:
        z = zipfile.ZipFile(io.BytesIO(get(GBB + "GasBBActualFlowStorage.zip").content))
        d = pd.read_csv(z.open(z.namelist()[0]), low_memory=False)
    else:
        d = pd.read_csv(io.BytesIO(get(GBB + "GasBBActualFlowStorageLast31.CSV").content), low_memory=False)
    d["date"] = pd.to_datetime(d["GasDate"], errors="coerce")
    for c in ("Demand", "Supply", "HeldInStorage"):
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.dropna(subset=["date"])
    print(f"  {'full history' if full else 'last 31 days'}: {len(d):,} rows {d['date'].min():%Y-%m-%d}.."
          f"{d['date'].max():%Y-%m-%d}; types {sorted(d['FacilityType'].dropna().unique())}", flush=True)
    return d


def tables(d):
    s = d[d["FacilityType"].isin(SECTORS)]
    demand = s.pivot_table(index="date", columns="FacilityType", values="Demand", aggfunc="sum").rename(columns=SECTORS)
    demand = demand[demand.index >= SECTOR_START].reindex(columns=list(SECTORS.values()))
    demand["Total"] = demand.sum(axis=1, min_count=1)
    p = d[d["FacilityType"].eq("PROD")]
    prod = p.pivot_table(index="date", columns="State", values="Supply", aggfunc="sum")
    prod["Total"] = prod.sum(axis=1, min_count=1)
    st = d[d["FacilityType"].eq("STOR")]
    # one HeldInStorage per facility and day (a facility can report several locations with the same figure)
    stor = st.groupby(["date", "FacilityName"])["HeldInStorage"].max().unstack()
    stor["Total"] = stor.sum(axis=1, min_count=1)
    return {"Demand by sector": demand, "Production": prod, "Storage": stor}


def shipments():
    d = pd.read_csv(io.BytesIO(get(GBB + "GasBBLNGShipments.CSV").content))
    d["date"] = pd.to_datetime(d["ShipmentDate"], format="%d %b %Y %H:%M:%S", errors="coerce")
    d["VolumePJ"] = pd.to_numeric(d["VolumePJ"], errors="coerce")
    d = d.sort_values("VersionDateTime").drop_duplicates("TransactionId", keep="last")   # latest version of each cargo
    m = d.pivot_table(index=d["date"].dt.to_period("M").dt.to_timestamp(), columns="FacilityName",
                      values="VolumePJ", aggfunc="sum")
    m["Total"] = m.sum(axis=1, min_count=1)
    m["Cargoes"] = d.groupby(d["date"].dt.to_period("M").dt.to_timestamp())["TransactionId"].count()
    m.index.name = "Month"
    return m


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--full", action="store_true", help="re-read the full-history zip")
    args = ap.parse_args()

    saved = {k: load(args.out, k) for k in ("Demand by sector", "Production", "Storage")}
    full = args.full or any(v.empty for v in saved.values())
    new = tables(flows(full))
    out = {}
    for k, v in new.items():
        old = saved[k]
        if not old.empty and not full:
            v = v.combine_first(old)   # new values win over the saved ones on the days both have
            v = v[[c for c in v.columns if c != "Total"] + ["Total"]]
        out[k] = v.sort_index()
    ship = shipments()
    out["LNG shipments"] = ship
    for k, v in out.items():
        if k != "LNG shipments":
            v.index.name = "date"
        print(f"  {k}: {len(v)} rows to {v.index.max():%Y-%m-%d}", flush=True)

    def cover(d):
        return f"{d.index.min():%d %b %Y} to {d.index.max():%d %b %Y}"

    notes = [
        "UNITS",
        "Demand by sector, Production: terajoules per gas day (TJ/day). Gas day 06:00-06:00 AEST, as published.",
        "Storage: TJ held in storage at the end of the gas day (HeldInStorage), by facility; Total = sum of the "
        "facilities reporting that day. 1 PJ = 1,000 TJ.",
        "LNG shipments: PJ loaded per month by plant (latest version of each cargo record); Cargoes = number "
        "of cargoes.",
        "Demand by sector: end-use facility types - gas-powered generation (BBGPG), large industrial users "
        "(BBLARGE), LNG export plants (LNGEXPORT: feed gas to the Curtis Island plants). Residential and "
        "commercial (distribution network) demand is not a Bulletin Board facility type and is not included.",
        "",
        "COVERAGE",
        f"East coast market and Northern Territory (no WA). Demand by sector {cover(out['Demand by sector'])} - AEMO "
        "extended the Bulletin Board to these facility types on 15 Mar 2023, nothing is reported before. "
        f"Production {cover(out['Production'])}; Storage {cover(out['Storage'])}; LNG shipments "
        f"{cover(out['LNG shipments'])}.",
        "",
        "SOURCE",
        f"AEMO Gas Bulletin Board: {GBB}GasBBActualFlowStorage.zip (full history), "
        "GasBBActualFlowStorageLast31.CSV (incremental updates), GasBBLNGShipments.CSV",
        "https://aemo.com.au/energy-systems/gas/gas-bulletin-board-gbb",
    ]
    xlsx_notes.write_workbook(args.out, {k: v.round(3) for k, v in out.items()}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
