"""
Australia installed generating capacity by fuel, from AEMO's registration
lists (public, no key):

  NEM  https://www.aemo.com.au/-/media/Files/Electricity/NEM/Participant_Information/NEM-Registration-and-Exemption-List.xls
       sheet "PU and Scheduled Loads": Region, Fuel Source / Technology Type descriptors, Reg Cap generation (MW)
  WEM  https://data.wa.aemo.com.au/public/public-data/datafiles/facilities/facilities.csv
       Facility Code, Facility Type, Maximum Capacity (MW) - fuel from the facility code (as in AU_WEM_GENERATION)

Both lists are snapshots of what is registered now, so each run writes the
current month's row and keeps the rows saved by earlier runs. NEM months before
the first snapshot (from 2021) are rebuilt once from AEMO's MMS unit history
(DUDETAIL registered capacity by effective date, DUDETAILSUMMARY region /
dispatch type / start and end dates, via nemosis); WEM is held at its first
snapshot for those months (no public WEM capacity history).

Writes au_power_capacity.xlsx: "Monthly" (standard layout: date, <Fuel>_MW, Total_MW, plus
Battery_storage_MW and Pumped_storage_MW kept out of Total_MW) and "By region" (MW by NEM region and WEM,
same month rows).

Usage: python3 AU_POWER_CAPACITY.py [--out "output/Data and Chart Outputs/au_power_capacity.xlsx"]
"""
import argparse
import io
import os
import sys
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
import aemo_registration  # noqa: E402
from AU_WEM_GENERATION import fuel_of  # noqa: E402

WEM_FACILITIES = "https://data.wa.aemo.com.au/public/public-data/datafiles/facilities/facilities.csv"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_power_capacity.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Nuclear", "Other"]
STORAGE = ["Battery_storage", "Pumped_storage"]


def nem_units():
    return aemo_registration.units()


HISTORY_START = "2021-01-01"


def nem_history(months):
    """NEM generating capacity by fuel and region at each month-end in `months`, rebuilt from MMS DUDETAIL
    (REGISTEREDCAPACITY, latest version effective by the month-end) for units registered at that month-end
    (DUDETAILSUMMARY START_DATE..END_DATE). Fuel from the registration list, plus units that have left it."""
    from AU_NEM_GENERATION import _table
    summ = aemo_registration.latest_mms(_table, "DUDETAILSUMMARY", pd.Timestamp(HISTORY_START))
    det = aemo_registration.latest_mms(_table, "DUDETAIL", pd.Timestamp(HISTORY_START))
    print(f"DUDETAILSUMMARY {summ.shape}, DUDETAIL {det.shape}; DUDETAIL effective dates "
          f"{det['EFFECTIVEDATE'].min()} .. {det['EFFECTIVEDATE'].max()}", flush=True)
    fuel = aemo_registration.with_history(aemo_registration.units(include_loads=True), None, None,
                                          lambda *a: summ)
    fuel = fuel[fuel["role"].eq("gen")].drop_duplicates("duid").set_index("duid")["fuel"]
    for c in ("START_DATE", "END_DATE"):
        summ[c] = pd.to_datetime(summ[c], errors="coerce")
    det["EFFECTIVEDATE"] = pd.to_datetime(det["EFFECTIVEDATE"], errors="coerce")
    det["MW"] = pd.to_numeric(det["REGISTEREDCAPACITY"], errors="coerce")
    det = det.sort_values(["EFFECTIVEDATE", "VERSIONNO"])
    gen = ~summ["DISPATCHTYPE"].astype(str).str.upper().str.contains("LOAD")
    summ = summ[gen & summ["REGIONID"].isin(["NSW1", "QLD1", "SA1", "TAS1", "VIC1"])]
    rows, regs = {}, {}
    for m in months:
        me = m + pd.offsets.MonthEnd(0)
        live = summ[(summ["START_DATE"] <= me) & (summ["END_DATE"].isna() | (summ["END_DATE"] > me))]
        live = live.sort_values("START_DATE").drop_duplicates("DUID", keep="last")
        cap = det[det["EFFECTIVEDATE"] <= me].drop_duplicates("DUID", keep="last").set_index("DUID")["MW"]
        # DUIDs with no fuel (demand response, ancillary-service and dummy units, Basslink) are not generators
        u = pd.DataFrame({"region": live["REGIONID"].str[:-1].values, "mw": live["DUID"].map(cap).values,
                          "fuel": live["DUID"].map(fuel).values}).dropna(subset=["fuel"])
        rows[m] = u.groupby("fuel")["mw"].sum()
        regs[m] = u[~u["fuel"].isin(STORAGE)].groupby("region")["mw"].sum()
    by_fuel, by_region = pd.DataFrame(rows).T, pd.DataFrame(regs).T
    print(f"NEM history rebuilt for {len(months)} months; last:\n{by_fuel.tail(2).round(0).to_string()}", flush=True)
    return by_fuel, by_region


def wem_units():
    r = requests.get(WEM_FACILITIES, headers={"User-Agent": "Mozilla/5.0"}, timeout=(10, 120))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    tcol = next((c for c in d.columns if "type" in c.lower()), None)
    if tcol:
        d = d[d[tcol].astype(str).str.contains("Gen|Storage", case=False)]   # generators and storage, not loads
    capcol = next((c for c in d.columns if "maximum capacity" in c.lower()),
                  next(c for c in d.columns if "capacity" in c.lower() and "credit" not in c.lower()))
    fuel = d["Facility Code"].map(fuel_of).replace({"Battery_discharge": "Battery_storage"})
    out = pd.DataFrame({"region": "WA (WEM)", "fuel": fuel, "mw": pd.to_numeric(d[capcol], errors="coerce")})
    print(f"WEM: {len(out)} facilities, {out['mw'].sum() / 1000:.1f} GW ({capcol}); types "
          f"{d[tcol].value_counts().to_dict() if tcol else 'n/a'}", flush=True)
    return out


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
    args = ap.parse_args()

    units = pd.concat([nem_units(), wem_units()], ignore_index=True)
    month = pd.Timestamp(date.today()).to_period("M").to_timestamp()
    by_fuel = units.groupby("fuel")["mw"].sum()
    row = pd.DataFrame([{f"{f}_MW": by_fuel.get(f) for f in FUELS if f in by_fuel}], index=[month])
    row["Total_MW"] = row.sum(axis=1)
    for s in STORAGE:
        if s in by_fuel:
            row[f"{s}_MW"] = by_fuel[s]
    reg = units[~units["fuel"].isin(STORAGE)].groupby("region")["mw"].sum().to_frame(month).T.add_suffix("_MW")

    monthly, region = load(args.out, "Monthly"), load(args.out, "By region")
    # this run's snapshot REPLACES any row saved earlier for the same month (no stale columns carried over)
    monthly = pd.concat([monthly.drop(index=month, errors="ignore"), row]).sort_index().dropna(axis=1, how="all")
    region = pd.concat([region.drop(index=month, errors="ignore"), reg]).sort_index().dropna(axis=1, how="all")
    # one-off: NEM months before the first snapshot, rebuilt from the MMS unit history (WEM at its first snapshot)
    first = monthly.index.min()
    gap = [m for m in pd.date_range(HISTORY_START, first, freq="MS") if m < first]
    if gap:
        try:
            hf, hr = nem_history(gap + [month])
            check = hf.loc[month]
            nem_now = nem_units().groupby("fuel")["mw"].sum()
            print("rebuilt vs registration list, this month (MW):\n" + pd.DataFrame(
                {"rebuilt": check, "registration list": nem_now}).round(0).to_string(), flush=True)
            wem = units[units["region"].eq("WA (WEM)")].groupby("fuel")["mw"].sum()
            wem_first = region.loc[first, "WA (WEM)_MW"] if "WA (WEM)_MW" in region else wem[~wem.index.isin(STORAGE)].sum()
            hf = hf.drop(index=month)
            for f, v in wem.items():   # WEM held at its current snapshot
                hf[f] = hf[f].fillna(0) + v if f in hf else v
            hist = pd.DataFrame({f"{f}_MW": hf[f] for f in FUELS if f in hf}, index=hf.index)
            hist["Total_MW"] = hist.sum(axis=1)
            for st in STORAGE:
                if st in hf:
                    hist[f"{st}_MW"] = hf[st]
            hr = hr.drop(index=month).add_suffix("_MW")
            hr["WA (WEM)_MW"] = wem_first
            monthly = pd.concat([hist, monthly]).sort_index()
            region = pd.concat([hr, region]).sort_index()
        except Exception as e:  # noqa: BLE001 - keep the snapshots
            print(f"NEM history rebuild failed: {type(e).__name__}: {str(e)[:300]}", flush=True)
    monthly = monthly[[c for c in row.columns] + [c for c in monthly.columns if c not in row.columns]]
    for x in (monthly, region):
        x.index.name = "date"
    print(monthly.tail(3).to_string(), flush=True)
    notes = [
        "UNITS",
        "Registered generating capacity, MW, as registered on the date of each run (one row per month, the "
        "latest run in the month wins). NEM: 'Reg Cap generation (MW)' per unit; WEM: 'Maximum Capacity (MW)'.",
        "Monthly: Hydro (excl. pumped storage), Gas, Wind, Solar (utility scale), Coal, Oil, Bioenergy; "
        "Battery_storage_MW and Pumped_storage_MW are storage, kept out of Total_MW.",
        "NEM fuel = AEMO's Fuel Source / Technology Type descriptors; WEM fuel = from the facility code (AEMO's "
        "WEM list has no fuel field; see au_wem_power_generation_daily.xlsx Units for the rules).",
        "By region: generating capacity (excl. storage) by NEM region and WEM.",
        "",
        "COVERAGE",
        "NEM (QLD, NSW, VIC, SA, TAS) + WA's WEM. Rooftop solar, NT and off-grid plant not included. Months from "
        "Jan 2021 to Sep 2026: NEM rebuilt from AEMO MMS unit history (DUDETAIL registered capacity, "
        "DUDETAILSUMMARY registration dates); WEM held at its Oct 2026 snapshot. From Oct 2026: monthly snapshots "
        "of the registration lists.",
        "",
        "SOURCE",
        f"AEMO NEM Registration and Exemption List: {aemo_registration.URL}; AEMO WA facilities list: "
        f"{WEM_FACILITIES}",
        "https://aemo.com.au/en/energy-systems/electricity/national-electricity-market-nem/participate-in-the-market/registration",
    ]
    xlsx_notes.write_workbook(args.out, {"Monthly": monthly.round(1), "By region": region.round(1)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
