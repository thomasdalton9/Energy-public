"""
Monthly installed generation capacity for Denmark and Norway from the official registers (nothing interpolated).

  output/Data and Chart Outputs/denmark_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/norway_power_capacity_monthly.xlsx

Date convention (same as the ENTSO-E annual files, which are dated 1 January): the row dated 1 May holds the capacity
installed at 00:00 on 1 May, i.e. at the end of April.

DENMARK - Energinet, Energi Data Service dataset 'CapacityPerMunicipality' (api.energidataservice.dk, TSO electricity;
  Danish Energy Agency master data summed over the 98 municipalities). RAW, MONTHLY from Jan 2016, capacity calculated at
  midnight on the first day of each month, published during the first two weeks of the next month. PARTIAL: offshore wind,
  onshore wind and solar PV by fuel; thermal plant is only split by size (>= 100 MW / < 100 MW), NOT by fuel, so it is
  kept as 'Thermal_unsplit' (Energinet says a version by main fuel is under development). Hydro (9 MW) is not in the
  dataset. Incremental: months stored are kept, the last 4 months are re-read (revisions) plus any missing month.

NORWAY - NVE (Norges vassdrags- og energidirektorat) hydropower database and wind power database (api.nve.no, keyless):
  month-end capacity REBUILT from each plant's / turbine group's commissioning date and, for decommissioned plant, its
  removal date. PARTIAL: hydro, pumped-storage plants and wind only (no other fuel has a dated register: Norway has
  ~0.5 GW gas, ~0.04 GW solar, almost nothing else). Limits: the hydro database carries only the CURRENT installed power
  of each plant, so later upgrades are back-dated to the plant's first commissioning date; a commissioning date of
  1 January is often a year-only placeholder; the removal date of a hydro plant is a YEAR (taken as 1 July of that year);
  wind has exact turbine-group dates (commissioning and removal). Offshore wind (Hywind Tampen, 88 MW, supplies oil
  platforms) is not in the wind database. NVE's 'Pumpekraftverk' (pumped-storage plants with a generating turbine) are
  kept in PumpedStorage_MW and also summed with conventional hydro in Hydro_incl_pumped_MW, because ENTSO-E's 'Hydro'
  matches that sum (34,029 MW on 1 Jan 2026 against 34,084). Plants listed 'under construction' are not counted.
  The registers are whole-file sources: each run downloads them, merges them into the stored 'Register_*' sheets
  (a plant that disappears from the source keeps its stored dates) and rebuilds the months; the hash of each file is
  recorded on 'Releases'.

Cross-check: sheet 'Check_ENTSOE' sets every 1 January against ENTSO-E A68 (europe/ENTSOE_CAPACITY.py workbooks).

Usage: python3 NORDIC_CAPACITY_MONTHLY.py [--out-dir DIR] [--only dk|no]
"""
import argparse
import hashlib
import json
import os
import sys
import time
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
DK_FILE = "denmark_power_capacity_monthly.xlsx"
NO_FILE = "norway_power_capacity_monthly.xlsx"
SES = requests.Session()
SES.headers["User-Agent"] = "energy-data-pulls/1.0"
NOW = pd.Timestamp(datetime.now(timezone.utc).date()).to_period("M").to_timestamp()

DK_COLS = ["Wind_Offshore_MW", "Wind_Onshore_MW", "Wind_MW", "Solar_MW", "Thermal_ge100MW_MW", "Thermal_lt100MW_MW",
           "Thermal_unsplit_MW"]
NO_COLS = ["Hydro_MW", "PumpedStorage_MW", "Hydro_incl_pumped_MW", "Wind_MW", "Total_MW"]


def get_json(url, params=None, tries=6):
    for i in range(tries):
        r = SES.get(url, params=params, timeout=(20, 180))
        if r.status_code == 429:
            wait = 60 * (i + 1)
            print(f"  429 from {url[:60]}, waiting {wait}s")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json(), r.content
    raise RuntimeError(f"giving up on {url}")


def read_sheet(path, sheet, index_col=None):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {sheet} from {os.path.basename(path)}: {type(e).__name__}")
        return None


def releases_frame(rows):
    return pd.DataFrame(rows, columns=["source", "checked_utc", "detail", "sha256"]).set_index("source")


def entsoe_check(out_dir, country, monthly, mapping):
    """1 January values against ENTSO-E A68 ({our column: ENTSO-E column})."""
    p = os.path.join(out_dir, f"{country}_power_capacity.xlsx")
    if not os.path.exists(p):
        return pd.DataFrame()
    e = pd.read_excel(p, sheet_name="Monthly")
    e["date"] = pd.to_datetime(e["date"])
    e = e.set_index("date")
    rows = []
    for d in e.index:
        if d not in monthly.index:
            continue
        for ours, theirs in mapping.items():
            if ours not in monthly or theirs not in e:
                continue
            a, b = monthly.at[d, ours], e.at[d, theirs]
            if pd.isna(a):
                continue
            rows.append({"date": d, "series": ours, "this_file_MW": round(float(a), 1), "ENTSOE_A68_MW": float(b),
                         "difference_MW": round(float(a) - float(b), 1),
                         "difference_pct": round((float(a) - float(b)) / float(b) * 100, 1) if b else np.nan})
    return pd.DataFrame(rows).set_index("date") if rows else pd.DataFrame()


# ----------------------------------------------------------------------------------------------------- Denmark
def dk_pull(path):
    old = read_sheet(path, "Monthly", 0)
    start = pd.Timestamp("2016-01-01")
    if old is not None and len(old):
        old.index = pd.to_datetime(old.index)
        start = max(start, old.index.max() - pd.DateOffset(months=4))
    params = {"start": start.strftime("%Y-%m-%d"), "end": (NOW + pd.DateOffset(months=1)).strftime("%Y-%m-%d"),
              "limit": 0, "sort": "Month asc"}
    js, raw = get_json("https://api.energidataservice.dk/dataset/CapacityPerMunicipality", params)
    d = pd.DataFrame(js["records"])
    if d.empty:
        raise RuntimeError("Energi Data Service returned no capacity rows")
    d["Month"] = pd.to_datetime(d["Month"])
    num = ["OffshoreWindCapacity", "OnshoreWindCapacity", "SolarPowerCapacity", "CapacityGe100MW", "CapacityLt100MW"]
    for c in num:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    g = d.groupby("Month")[num].sum(min_count=1)
    new = pd.DataFrame(index=g.index)
    new["Wind_Offshore_MW"] = g["OffshoreWindCapacity"]
    new["Wind_Onshore_MW"] = g["OnshoreWindCapacity"]
    new["Wind_MW"] = g["OffshoreWindCapacity"].fillna(0) + g["OnshoreWindCapacity"].fillna(0)
    new["Solar_MW"] = g["SolarPowerCapacity"]
    new["Thermal_ge100MW_MW"] = g["CapacityGe100MW"]
    new["Thermal_lt100MW_MW"] = g["CapacityLt100MW"]
    new["Thermal_unsplit_MW"] = g[["CapacityGe100MW", "CapacityLt100MW"]].sum(axis=1, min_count=1)
    new = new.round(3)
    monthly = new if old is None else pd.concat([old[~old.index.isin(new.index)][DK_COLS], new]).sort_index()
    monthly.index.name = "date"
    rel = [("Energi Data Service CapacityPerMunicipality", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
            f"{len(d)} rows from {start.date()}; latest month {g.index.max().date()}", hashlib.sha256(raw).hexdigest()[:16])]
    return monthly, rel


def dk_notes():
    lines = ["Denmark - monthly installed generation capacity (MW)", "",
             "Source", "Energinet, Energi Data Service dataset CapacityPerMunicipality (Danish Energy Agency master data, summed over the 98 municipalities): https://api.energidataservice.dk/dataset/CapacityPerMunicipality",
             "Raw official source, MONTHLY from Jan 2016. Capacity is calculated at midnight on the first day of the month, so the row dated 1 May is the capacity at the end of April. Published in the first two weeks of the next month.",
             "", "Columns",
             "Wind_Offshore_MW, Wind_Onshore_MW, Wind_MW (sum), Solar_MW (all registered PV incl. small roof systems): by fuel.",
             "Thermal_ge100MW_MW / Thermal_lt100MW_MW / Thermal_unsplit_MW (sum): thermal plant is split by SIZE only, not by fuel. Energinet notes a version by main fuel is under development; until then gas, coal, oil and biomass cannot be separated.",
             "Hydro (about 9 MW) and batteries are not in the dataset. PARTIAL series: wind, solar and thermal in total.",
             "", "Method", "Incremental: stored months are kept; the last 4 months and any new month are re-read each run. Nothing is interpolated or filled.",
             "", "Sheets", "Monthly, Check_ENTSOE (1 January against ENTSO-E A68), Releases"]
    titles = ["Source", "Columns", "Method", "Sheets"]
    return lines, titles


def run_dk(out_dir):
    path = os.path.join(out_dir, DK_FILE)
    monthly, rel = dk_pull(path)
    chk = entsoe_check(out_dir, "denmark", monthly, {"Wind_MW": "Wind_MW", "Solar_MW": "Solar_MW", "Thermal_unsplit_MW": "Total_MW"})
    if len(chk):   # ENTSO-E total includes wind and solar: compare thermal with Total less wind less solar instead
        e = pd.read_excel(os.path.join(out_dir, "denmark_power_capacity.xlsx"), sheet_name="Monthly")
        e["date"] = pd.to_datetime(e["date"])
        e = e.set_index("date")
        chk = chk[chk["series"] != "Thermal_unsplit_MW"]
        th = (e["Total_MW"] - e["Wind_MW"] - e["Solar_MW"]).dropna()
        rows = []
        for d, v in th.items():
            if d in monthly.index and pd.notna(monthly.at[d, "Thermal_unsplit_MW"]):
                a = monthly.at[d, "Thermal_unsplit_MW"]
                rows.append({"date": d, "series": "Thermal_unsplit_MW (ENTSO-E: total less wind less solar)", "this_file_MW": round(a, 1),
                             "ENTSOE_A68_MW": float(v), "difference_MW": round(a - v, 1), "difference_pct": round((a - v) / v * 100, 1)})
        chk = pd.concat([chk, pd.DataFrame(rows).set_index("date")]).sort_index() if rows else chk
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Releases"] = releases_frame(rel)
    notes, titles = dk_notes()
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"Denmark: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}; last row\n{monthly.tail(1).T}")


# ----------------------------------------------------------------------------------------------------- Norway
HYDRO_COLS = ["id", "name", "type", "MW", "commissioned", "removed_year", "status", "last_seen"]
WIND_COLS = ["key", "plant", "turbines", "kW_each", "commissioned", "removed", "last_seen"]


def merge_store(old, new, key):
    """Stored rows win only for keys the source no longer lists."""
    if old is None or old.empty:
        return new
    old = old.copy()
    old[key] = old[key].astype(str)
    new[key] = new[key].astype(str)
    keep = old[~old[key].isin(new[key])]
    return pd.concat([keep, new], ignore_index=True)


def no_pull(path):
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    hj, hraw = get_json("https://api.nve.no/web/Powerplant/GetHydroPowerPlants")
    wj, wraw = get_json("https://api.nve.no/web/WindPowerplant/GetWindPowerplants")
    hyd = pd.DataFrame([{"id": x["VannKraftverkID"], "name": x["Navn"], "type": x["VannKVType"], "MW": x.get("MaksYtelse"),
                         "commissioned": pd.to_datetime(x.get("IDriftDato"), errors="coerce"),
                         "removed_year": x.get("UteAvDrift") if x.get("Kraftverkstatus") == "Ute av drift" else None,
                         "status": x.get("Kraftverkstatus"), "last_seen": today} for x in hj])
    rows = []
    for p in wj:
        for t in p.get("Turbiner") or []:
            rows.append({"key": f"{p['VindkraftAnleggId']}-{t.get('TurbinID')}-{str(t.get('DatoIdriftsatt'))[:10]}-{str(t.get('DatoUtavdrift'))[:10]}",
                         "plant": p["Navn"], "turbines": t.get("AntallTurbiner"), "kW_each": t.get("TurbinStorrelse_kW"),
                         "commissioned": pd.to_datetime(t.get("DatoIdriftsatt"), errors="coerce"),
                         "removed": pd.to_datetime(t.get("DatoUtavdrift"), errors="coerce"), "last_seen": today})
    win = pd.DataFrame(rows)
    oh = read_sheet(path, "Register_hydro")
    ow = read_sheet(path, "Register_wind")
    hyd = merge_store(oh, hyd, "id")
    win = merge_store(ow, win, "key")
    for c in ("commissioned",):
        hyd[c] = pd.to_datetime(hyd[c], errors="coerce")
        win[c] = pd.to_datetime(win[c], errors="coerce")
    win["removed"] = pd.to_datetime(win["removed"], errors="coerce")
    rel = [("NVE GetHydroPowerPlants", today, f"{len(hj)} plants", hashlib.sha256(hraw).hexdigest()[:16]),
           ("NVE GetWindPowerplants", today, f"{len(wj)} plants, {len(rows)} turbine groups", hashlib.sha256(wraw).hexdigest()[:16])]
    return hyd[HYDRO_COLS], win[WIND_COLS], rel


def no_rebuild(hyd, win):
    months = pd.date_range("2010-01-01", NOW, freq="MS")
    h = hyd[(hyd["status"] != "Under bygging") & hyd["commissioned"].notna() & (pd.to_numeric(hyd["MW"], errors="coerce") > 0)].copy()
    h["MW"] = pd.to_numeric(h["MW"])
    h["end"] = pd.to_datetime(h["removed_year"].apply(lambda y: f"{int(y)}-07-01" if pd.notna(y) and str(y) not in ("", "nan") else None), errors="coerce")
    out = pd.DataFrame(index=months)
    for label, typ in (("Hydro_MW", "Kraftverk"), ("PumpedStorage_MW", "Pumpekraftverk")):
        s = h[h["type"] == typ]
        out[label] = [float(s.loc[(s["commissioned"] < d) & (s["end"].isna() | (s["end"] >= d)), "MW"].sum()) for d in months]
    out["Hydro_incl_pumped_MW"] = out["Hydro_MW"] + out["PumpedStorage_MW"]
    w = win[win["commissioned"].notna()].copy()
    w["MW"] = pd.to_numeric(w["turbines"]) * pd.to_numeric(w["kW_each"]) / 1000.0
    out["Wind_MW"] = [float(w.loc[(w["commissioned"] < d) & (w["removed"].isna() | (w["removed"] >= d)), "MW"].sum()) for d in months]
    out["Total_MW"] = out["Hydro_MW"] + out["Wind_MW"]   # excludes pumped storage, as in the other capacity workbooks
    out = out.round(2)
    out.index.name = "date"
    return out[NO_COLS]


def no_notes():
    lines = ["Norway - monthly installed generation capacity (MW), rebuilt from NVE's registers", "",
             "Source", "NVE (Norwegian Water Resources and Energy Directorate) hydropower database https://api.nve.no/web/Powerplant/GetHydroPowerPlants and wind power database https://api.nve.no/web/WindPowerplant/GetWindPowerplants (keyless).",
             "Raw official registers; month-end capacity is REBUILT from commissioning and removal dates, so it is a reconstruction, not a published monthly series. The row dated 1 May is the capacity at the end of April.",
             "", "What is covered", "Hydro (conventional plants), pumped-storage plants (NVE 'Pumpekraftverk') and onshore wind. PARTIAL: gas (~0.5 GW), solar, bioenergy and other fuels have no dated register here. Total_MW = hydro + wind (excludes pumped storage).",
             "Hydro_incl_pumped_MW = conventional hydro + pumped-storage plants; ENTSO-E's Hydro matches this sum (see Check_ENTSOE).",
             "", "Limits",
             "The hydro database holds only each plant's CURRENT installed power: later upgrades are back-dated to first commissioning, so earlier years can be too high for upgraded plants, while plants removed from the database altogether are missing.",
             "A hydro commissioning date of 1 January is often a year-only placeholder. A removed hydro plant has a removal YEAR only (taken as 1 July of that year). Wind has exact turbine-group commissioning and removal dates.",
             "Offshore wind (Hywind Tampen, 88 MW, feeds oil platforms) and plants under construction are not counted.",
             "", "Method", "Each run downloads both registers, merges them into the stored Register_hydro / Register_wind sheets (a plant that disappears from the source keeps its stored dates) and rebuilds every month from Jan 2010. The file hashes are on Releases.",
             "", "Sheets", "Monthly, Check_ENTSOE, Register_hydro, Register_wind, Releases"]
    return lines, ["What is covered", "Limits", "Method", "Sheets", "Source"]


def run_no(out_dir):
    path = os.path.join(out_dir, NO_FILE)
    hyd, win, rel = no_pull(path)
    monthly = no_rebuild(hyd, win)
    chk = entsoe_check(out_dir, "norway", monthly, {"Hydro_MW": "Hydro_MW", "PumpedStorage_MW": "PumpedStorage_MW", "Wind_MW": "Wind_MW"})
    e = read_sheet(os.path.join(out_dir, "norway_power_capacity.xlsx"), "Monthly")
    if e is not None:
        e["date"] = pd.to_datetime(e["date"])
        e = e.set_index("date")
        rows = []
        for d in e.index:
            if d in monthly.index:
                a, b = monthly.at[d, "Hydro_incl_pumped_MW"], e.at[d, "Hydro_MW"] + e.at[d, "PumpedStorage_MW"]
                rows.append({"date": d, "series": "Hydro_incl_pumped_MW (ENTSO-E: Hydro + PumpedStorage)", "this_file_MW": round(a, 1),
                             "ENTSOE_A68_MW": float(b), "difference_MW": round(a - b, 1), "difference_pct": round((a - b) / b * 100, 1)})
        if rows:
            chk = pd.concat([chk, pd.DataFrame(rows).set_index("date")]).sort_index() if len(chk) else pd.DataFrame(rows).set_index("date")
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Register_hydro"] = hyd
    sheets["Register_wind"] = win
    sheets["Releases"] = releases_frame(rel)
    notes, titles = no_notes()
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"Norway: {len(monthly)} months; 1 Jan values\n{monthly[monthly.index.month == 1].tail(6)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--only", choices=["dk", "no"])
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    fails = []
    for key, fn in (("dk", run_dk), ("no", run_no)):
        if a.only in (None, key):
            try:
                fn(a.out_dir)
            except Exception as e:  # noqa: BLE001
                import traceback
                traceback.print_exc()
                fails.append(f"{key}: {type(e).__name__}: {e}")
    if fails:
        print("FAILED:", fails)
        sys.exit(1)


if __name__ == "__main__":
    main()
