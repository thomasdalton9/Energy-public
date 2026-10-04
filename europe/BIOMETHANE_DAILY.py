"""
Biomethane (upgraded biogas) injected into the gas grids, from operator / statistical sources, for the Europe gas balances.
Biomethane is injected mostly into distribution grids, so ENTSOG's transmission-level 'production' misses most of it.

  output/Data and Chart Outputs/europe_biomethane_operators.xlsx
    sheet "Daily"    date (gas day), GWh per day:   FR_biomethane  (ODRE: GRTgaz/NaTran + Teréga + GRDF/other DSOs, definitive data)
                                                    DK_biomethane  (Energinet Gasflow KWhFromBiogas)
    sheet "Monthly"  month start, GWh per month:    NL_biomethane  (CBS StatLine 86103NED 'productie uit andere bronnen' = green gas)
    sheet "Annual"   year, GWh per year:            DE_biomethane (when a source was found, see Units)
    sheet "Units"    source, unit and definition of every column

Incremental: the committed workbook is the history store; each run re-fetches the last REVISION_DAYS days (operators restate
recent days) plus any gap; the CBS table is small and is re-read whole each run (its figures are revised).

Usage: python3 BIOMETHANE_DAILY.py [--out-dir DIR] [--start 2020-01-01] [--only FR,DK,NL,DE]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_biomethane_operators.xlsx"
REVISION_DAYS = 21
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
DAILY_COLS = ["FR_biomethane", "DK_biomethane"]
MONTHLY_COLS = ["NL_biomethane"]
ANNUAL_COLS = ["DE_biomethane"]
# Dutch gas is quoted in m3 of 35.17 MJ (Groningen-equivalent standard m3): 35.17 / 3.6 = 9.769 kWh per m3
NL_KWH_PER_M3 = 35.17 / 3.6


def get(url, tries=3, **kw):
    kw.setdefault("headers", {"User-Agent": UA})
    for i in range(tries):
        try:
            r = requests.get(url, timeout=(15, 90), **kw)
            if r.ok:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code} {r.text[:200]} for {r.url}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError(f"gave up on {url}")


# ---- France: ODRE ------------------------------------------------------------------------------------------------------
ODRE = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{ds}/records"
# Daily definitive regionalised biomethane production on the transport AND distribution grids (MWh per day per region);
# 'Meilleur Statut' = definitive value where it exists, else the latest provisional value.
FR_DS = "prod-def-reg-jour-biom-reseau-grtgrd"
FR_FIELD = "production_biomethane"


def france(d0, d1):
    rec, offset = {}, 0
    while True:
        r = get(ODRE.format(ds=FR_DS), params={
            "select": f"date,sum({FR_FIELD}) as v", "group_by": "date", "order_by": "date", "limit": 100, "offset": offset,
            "where": f"date>=date'{d0.isoformat()}' and date<=date'{d1.isoformat()}'"})
        res = r.json().get("results", [])
        for row in res:
            if row.get("v") is not None:
                rec[pd.Timestamp(row["date"][:10])] = float(row["v"]) / 1000.0        # MWh -> GWh
        if len(res) < 100:
            break
        offset += 100
    return pd.DataFrame({"FR_biomethane": pd.Series(rec, dtype=float)}).sort_index()


# ---- Denmark: Energinet --------------------------------------------------------------------------------------------------
ENERGINET = "https://api.energidataservice.dk/dataset/Gasflow"


def denmark(d0, d1):
    r = get(ENERGINET, params={"start": d0.isoformat(), "end": (d1 + timedelta(days=1)).isoformat(), "limit": 100000, "sort": "GasDay ASC"})
    d = pd.DataFrame(r.json().get("records", []))
    if d.empty:
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["GasDay"]).dt.normalize()
    d = d.drop_duplicates("date", keep="last").set_index("date")
    return pd.DataFrame({"DK_biomethane": pd.to_numeric(d["KWhFromBiogas"], errors="coerce") / 1e6})


# ---- Netherlands: CBS ----------------------------------------------------------------------------------------------------
CBS = "https://opendata.cbs.nl/ODataApi/odata/86103NED/TypedDataSet"


def netherlands():
    r = get(CBS, params={"$select": "Perioden,ProductieUitAndereBronnen_3", "$filter": "substringof('MM',Perioden)", "$top": 1000})
    d = pd.DataFrame(r.json()["value"])
    d["month"] = pd.to_datetime(d["Perioden"].str.replace("MM", "-"), format="%Y-%m")
    d = d.dropna(subset=["ProductieUitAndereBronnen_3"]).set_index("month").sort_index()
    return pd.DataFrame({"NL_biomethane": d["ProductieUitAndereBronnen_3"].astype(float) * NL_KWH_PER_M3 / 1000.0})   # mln m3 -> GWh


# ---- storage -------------------------------------------------------------------------------------------------------------
def read_sheet(path, sheet, cols, idx):
    if not os.path.exists(path):
        return pd.DataFrame(columns=cols)
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception as e:  # noqa: BLE001
        print(f"could not read {sheet} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=cols)
    d[idx] = pd.to_datetime(d[idx], errors="coerce")
    d = d.dropna(subset=[idx]).set_index(idx).sort_index()
    for c in cols:
        if c not in d:
            d[c] = float("nan")
    return d[cols]


def merge(combined, new, cols):
    combined = combined.reindex(combined.index.union(new.index))
    for c in cols:
        if c in new:
            s = new[c].dropna()
            combined.loc[s.index, c] = s
    return combined


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2020-01-01")
    ap.add_argument("--only", default="", help="comma-separated subset of FR,DK,NL (the workbook keeps the other countries' history)")
    args = ap.parse_args()
    only = {x.strip().upper() for x in args.only.split(",") if x.strip()}
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    daily = read_sheet(path, "Daily", DAILY_COLS, "date")
    monthly = read_sheet(path, "Monthly", MONTHLY_COLS, "month")
    annual = read_sheet(path, "Annual", ANNUAL_COLS, "year")
    old_daily = daily.copy()
    for code, label, fn in (("FR", "France (ODRE)", france), ("DK", "Denmark (Energinet)", denmark)):
        if only and code not in only:
            continue
        col = f"{code}_biomethane"
        have = old_daily[col].dropna()
        fs = start if have.empty else max(start, have.index.max().date() - timedelta(days=REVISION_DAYS))
        print(f"{label}: start from {fs}", flush=True)
        try:
            new = fn(fs, today)
        except Exception as e:  # noqa: BLE001
            print(f"{label}: FAILED {type(e).__name__}: {e}", flush=True)
            continue
        if new.empty:
            print(f"{label}: no rows", flush=True)
            continue
        daily = merge(daily, new, [col])
        print(f"{label}: {len(new)} days {new.index.min():%Y-%m-%d}..{new.index.max():%Y-%m-%d}", flush=True)
    if not only or "NL" in only:
        try:
            new = netherlands()
            monthly = merge(monthly, new, ["NL_biomethane"])
            print(f"Netherlands (CBS): {len(new)} months {new.index.min():%Y-%m}..{new.index.max():%Y-%m}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"Netherlands (CBS): FAILED {type(e).__name__}: {e}", flush=True)
    daily = daily.sort_index().round(3).dropna(how="all")
    monthly = monthly.sort_index().round(2).dropna(how="all")
    annual = annual.sort_index().dropna(how="all")
    daily.index.name, monthly.index.name, annual.index.name = "date", "month", "year"
    if daily.empty and monthly.empty:
        print("no biomethane data")
        return
    print("days per column:", daily.notna().sum().to_dict())
    print("annual TWh (daily sources):\n", (daily.resample("YS").sum(min_count=1) / 1000).round(2).to_string())
    print("annual TWh (monthly sources):\n", (monthly.resample("YS").sum(min_count=1) / 1000).round(2).to_string())
    lines = [
        "Europe - biomethane (upgraded biogas) injected into the gas grids, operator / statistical sources", "",
        "Columns",
        "FR_biomethane (sheet Daily, GWh per gas day): ODRE open data (https://odre.opendatasoft.com), dataset prod-def-reg-jour-biom-reseau-grtgrd - "
        "daily definitive regionalised biomethane production of the sites injecting into the GRTgaz (NaTran), Teréga and distribution "
        "(GRDF, regional DSO) grids, summed over regions; MWh converted to GWh (PCS, gross calorific value). Biomethane only, not raw biogas. "
        "Recent days are provisional ('Meilleur Statut': definitive value where available, else provisional).",
        "DK_biomethane (sheet Daily, GWh per gas day): Energinet Energi Data Service dataset Gasflow, field KWhFromBiogas "
        "(https://api.energidataservice.dk/dataset/Gasflow) - upgraded biogas injected into the Danish gas network (biomethane, not raw biogas). "
        "It is ALREADY inside DK_total of europe_tso_gas_demand_daily.xlsx (Denmark consumption = gas from the transmission system + biogas).",
        "NL_biomethane (sheet Monthly, GWh per month): CBS StatLine 86103NED 'Aardgasbalans', 'Productie uit andere bronnen' "
        "(https://opendata.cbs.nl/ODataApi/odata/86103NED) - natural gas obtained by converting other energy carriers, chiefly upgraded "
        "biogas (green gas), a little refinery-gas conversion. Million m3 (35.17 MJ/m3) x 9.769 kWh/m3.",
        "DE_biomethane (sheet Annual, GWh per year): see the notes below (filled only if a source was found).", "",
        "Notes",
        "Biomethane here means gas upgraded to grid quality and injected into the grid; raw biogas burned on site is excluded. "
        "Which TSO consumption/production series already include it must be checked per country (Denmark yes; France ODRE consumption is "
        "grid offtake and therefore includes it, but ENTSOG production does not).",
        f"Daily sources re-fetch the last {REVISION_DAYS} days each run plus gaps; history from {args.start}.", "",
        "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC"]
    sheets = {"Daily": daily, "Monthly": monthly}
    if not annual.empty:
        sheets["Annual"] = annual
    xlsx_notes.write_workbook(path, sheets, lines, {"Columns", "Notes", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
