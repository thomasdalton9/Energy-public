"""
Taiwan generation by fuel from Taipower's LIVE snapshot (output/Data and Chart Outputs/taiwan_live_daily.xlsx).

Taipower's open-data file d006001 (data.gov.tw 8931, 'real-time generation of every unit incl. purchased power', refreshed
every 10 minutes, NO history) lists every unit incl. independent power producers, solar, wind and storage. This script is
polled about hourly; each new snapshot (by its own DateTime stamp, Taiwan local time) is added to its day's sums and the
workbook keeps only the DAILY roll-up:

  Live daily   date x polls (snapshots counted that day), <Fuel>_GW = mean over those snapshots, <Fuel>_sumMW = running sums
               (so the next poll can be added without a history file), first/last snapshot stamp, charted = polls >= 20.
  Coverage     monthly mean of charted days against E-STAT national generation for the same month (months with >= 90% of
               their days charted only; E-STAT is two months in arrears, so this fills as the months arrive).

This is a SEPARATE series from taiwan_generation_rollup.xlsx: a sample of instantaneous readings (not interval energy),
only the polls actually made, nothing filled; a missed poll is simply not counted. It starts at the first poll and exists
because Taipower's 10-minute history file lags about three months. Subtotal rows ('小計') are skipped.

    python3 taiwan/TAIWAN_LIVE_SNAPSHOT.py --out "output/Data and Chart Outputs/taiwan_live_daily.xlsx"
"""
import argparse
import calendar
import json
import os
import re
import sys

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

URL = "https://service.taipower.com.tw/data/opendata/apply/file/d006001/001.json"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
DATA = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(DATA, "taiwan_live_daily.xlsx")
ESIST = os.path.join(DATA, "taiwan_esist_monthly.xlsx")
MIN_POLLS = 20
# Taipower label (group of units) -> stored column
GROUPS = {"燃氣": "Gas_Taipower", "民營電廠-燃氣": "Gas_IPP", "燃煤": "Coal_Taipower", "民營電廠-燃煤": "Coal_IPP", "汽電共生": "Cogeneration",
          "燃料油": "Oil", "太陽能": "Solar", "風力": "Wind", "水力": "Hydro", "其它再生能源": "Other_renewables",
          "儲能": "Storage_discharge", "儲能負載": "Storage_load"}
COLS = ["Gas_Taipower", "Gas_IPP", "Coal_Taipower", "Coal_IPP", "Cogeneration", "Oil", "Solar", "Wind", "Hydro", "Other_renewables",
        "Storage_discharge", "Storage_load"]
# charted fuels: IPP rows are inside Gas and Coal below (Gas_IPP and Coal_IPP are 'of which' memo columns)
ESTAT = {"Hydro": "Renewable Energy - Hydro", "Solar": "Renewable Energy - Solar PV", "Wind": "Renewable Energy - Wind",
         "Coal": "Thermal - Coal-Fired", "Gas": "Thermal - LNG-Fired", "Oil": "Thermal - Oil-Fired"}


def parse(js):
    stamp = js["DateTime"]
    sums = dict.fromkeys(COLS, 0.0)
    n = 0
    for r in js["aaData"]:
        lab = re.sub(r"<[^>]+>", "", str(r.get("機組類型", ""))).strip()
        lab = re.sub(r"\(Energy Storage System Load\)", "", lab).strip()
        col = GROUPS.get(lab)
        name = re.sub(r"<[^>]+>", "", str(r.get("機組名稱", "")))
        if col is None or "小計" in name:
            continue
        try:
            sums[col] += float(r["淨發電量(MW)"])
            n += 1
        except (ValueError, TypeError):
            continue
    return stamp, sums, n


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    old = None
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Live daily", index_col=0)
            old.index = pd.to_datetime(old.index)
        except Exception as e:  # noqa: BLE001
            print(f"stored sheet unreadable ({type(e).__name__}); starting again", file=sys.stderr)
    r = None
    for i in range(3):
        try:
            r = requests.get(URL, headers=UA, timeout=60)
            r.raise_for_status()
            js = json.loads(r.content.decode("utf-8-sig"))
            break
        except Exception as e:  # noqa: BLE001
            print(f"  attempt {i + 1}/3: {type(e).__name__}: {e}", file=sys.stderr)
            js = None
    if js is None:
        print("snapshot not reachable; nothing added")
        return
    stamp, sums, n = parse(js)
    day = pd.Timestamp(stamp[:10])
    print(f"snapshot {stamp}: {n} unit rows, " + ", ".join(f"{k} {v:.0f}" for k, v in sums.items()))
    if n < 150:
        print("too few unit rows (< 150): snapshot not used")
        return
    sum_cols = [f"{c}_sumMW" for c in COLS]
    if old is None:
        old = pd.DataFrame(columns=["polls"] + sum_cols + ["first_snapshot", "last_snapshot"])
        old.index.name = "date"
    if len(old) and day in old.index and str(old.loc[day, "last_snapshot"]) >= stamp:
        print("snapshot already counted")
    else:
        if day not in old.index:
            old.loc[day, ["polls"] + sum_cols] = [0] + [0.0] * len(sum_cols)
            old.loc[day, "first_snapshot"] = stamp
        old.loc[day, "polls"] = int(old.loc[day, "polls"]) + 1
        for c in COLS:
            old.loc[day, f"{c}_sumMW"] = float(old.loc[day, f"{c}_sumMW"]) + sums[c]
        old.loc[day, "last_snapshot"] = stamp
    old = old.sort_index()
    for c in ["polls"] + sum_cols:
        old[c] = pd.to_numeric(old[c])
    d = old.copy()
    g = lambda c: d[f"{c}_sumMW"] / d["polls"] / 1000.0   # noqa: E731
    d["Gas_GW"], d["Coal_GW"] = (g("Gas_Taipower") + g("Gas_IPP")).round(4), (g("Coal_Taipower") + g("Coal_IPP")).round(4)
    for c in ("Cogeneration", "Oil", "Solar", "Wind", "Hydro", "Other_renewables", "Storage_discharge", "Storage_load",
              "Gas_IPP", "Coal_IPP"):
        d[f"{c}_GW"] = g(c).round(4)
    d["Total_GW"] = d[[f"{c}_GW" for c in ("Gas", "Coal", "Cogeneration", "Oil", "Solar", "Wind", "Hydro", "Other_renewables",
                                           "Storage_discharge")]].sum(axis=1).round(4)
    d["charted"] = (d["polls"] >= MIN_POLLS).astype(int)
    gw = [f"{c}_GW" for c in ("Gas", "Coal", "Cogeneration", "Oil", "Solar", "Wind", "Hydro", "Other_renewables", "Storage_discharge",
                              "Storage_load", "Gas_IPP", "Coal_IPP", "Total")]
    d = d[["polls", "charted"] + gw + sum_cols + ["first_snapshot", "last_snapshot"]]
    d.index.name = "date"
    sheets = {"Live daily": d}
    cov = []
    try:
        em = pd.read_excel(ESIST, sheet_name="GEN M", index_col=0)
        em.index = pd.to_datetime(em.index)
        ok = d[d["charted"] == 1]
        for t, g in ok.groupby(ok.index.to_period("M").to_timestamp()):
            if t not in em.index or len(g) < 0.9 * t.days_in_month:
                continue
            hrs = t.days_in_month * 24
            for f, col in list(ESTAT.items()) + [("Total", "Grand Total")]:
                avg = g[f"{f}_GW"].mean()
                es = em.loc[t, col] / hrs
                cov.append(["Live snapshots", f"{t:%Y-%m}", len(g), f, round(avg, 3), round(es, 3), round(avg / es * 100, 1)])
    except Exception as e:  # noqa: BLE001
        print(f"coverage not computed: {type(e).__name__}: {e}")
    sheets["Coverage"] = pd.DataFrame(cov, columns=["series", "month", "days_charted", "fuel", "series_avg_GW", "E-STAT_avg_GW",
                                                    "series_as_pct_of_E-STAT"]).set_index("series")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES, {"UNITS", "SOURCE", "RULES"})
    print(f"Live daily: {len(d)} days {d.index.min():%Y-%m-%d}..{d.index.max():%Y-%m-%d}; last day polls {int(d['polls'].iloc[-1])}")


NOTES = [
    "UNITS",
    "<Fuel>_GW = mean over the day's snapshots of the summed net output of all units of the fuel (MW / 1000); <Fuel>_sumMW = "
    "running sum of those snapshot values; polls = snapshots counted that day; charted = 1 when polls >= 20. Gas and Coal "
    "INCLUDE the independent power producers (Gas_IPP, Coal_IPP are 'of which' columns); Storage_discharge = pumped-storage "
    "and battery output, Storage_load = pumping and battery charging (negative); Other_renewables = geothermal, biofuel and "
    "other small renewables as grouped by Taipower. Total_GW = all generation incl. storage discharge, excluding storage load. "
    "Nuclear is absent: Taiwan's last reactor stopped in May 2025.",
    "",
    "SOURCE",
    "Taiwan Power Company open data d006001 'real-time generation information of every unit incl. purchased power' "
    "(service.taipower.com.tw/data/opendata/apply/file/d006001/001.json; data.gov.tw 8931; Open Government Data License v1.0). "
    "Taipower refreshes it every 10 minutes and keeps no history.",
    "",
    "RULES",
    "Polled about hourly by GitHub Actions; each snapshot (stamped in Taiwan local time) is counted once. A missed poll is not "
    "filled. The daily value is a mean of instantaneous readings, not a metered daily energy, and is only as good as the number "
    "of polls (see polls). Subtotal rows are skipped. History starts at the first poll; Taipower's 10-minute history file "
    "(taiwan_generation_rollup.xlsx) lags about three months and is a different, EMS-only series.",
]

if __name__ == "__main__":
    main()
