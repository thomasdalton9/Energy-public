"""
Poland's gas entry and exit points from Gaz-System's own Market Information Module (raw operator data):

  output/Data and Chart Outputs/poland_gazsystem_points_daily.xlsx
    sheet "Daily": date (gas day), GWh per day, one column per Gaz-System point (KspRealization, "Actual quantity of gas transmitted",
        billing values, operative values for the latest days):
        BY_wysokoje      Wysokoje (572407): the Yamal-Europe entry from Belarus (Gazprom / EuRoPol Gaz gas, Kondratki), 2021 - May 2022
        BY_tietierowka   Tietierowka (572405): the second Belarus entry (small)
        DE_pwp           Point of Interconnection PWP (172434): Germany - Poland, Mallnow (reverse flow on the Yamal line) and Lasow
        DE_ontras        GCP GAZ-SYSTEM / ONTRAS (172435)
        UA_gcp           GCP GAZ-SYSTEM / UA TSO (172437): Ukraine entry (Hermanowice, Drozdowicze)
        CZ_cieszyn       Cieszyn (370001) entry from Czechia;  CZ_branice  Branice (372414)
        LNG              Terminal LNG Swinoujscie (777001) entry
        PROD             domestic production entries (E-gas, L-gas, nitrogen removal plant: 907029, 907030, 907033)
        STO_out / STO_in storage entries (Kawerna, Sanok, Wierzchowice withdrawals) and exits (injections)
        EXIT_pwp, EXIT_ontras, EXIT_ua, EXIT_cz   exits at the same interconnection points
    sheet "Units": source and definitions

ENTSOG carries no Kondratki / Wysokoje rows, so Poland's Russian (Belarus) imports of 2021-22 are missing from the ENTSOG-based balance;
the master adds BY_wysokoje + BY_tietierowka to Poland's pipeline imports (EUROPE_MASTER.point_fix_args).

Incremental: reads the committed workbook, re-fetches the last 45 days plus any gap; history from 2021-01-01.
Usage: python3 GAZSYSTEM_ENTRIES_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

GS = "https://swi.gaz-system.pl/mir/api/v1/en/KspRealization/data"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "poland_gazsystem_points_daily.xlsx"
RELOAD_DAYS = 45
ZONES = {"BY_wysokoje": ["572407"], "BY_tietierowka": ["572405"], "DE_pwp": ["172434"], "DE_ontras": ["172435"], "UA_gcp": ["172437"],
         "CZ_cieszyn": ["370001"], "CZ_branice": ["372414"], "LNG": ["777001"], "PROD": ["907029", "907030", "907033"],
         "STO_out": ["178001", "178002", "178003"], "STO_in": ["108001", "108002", "108003"],
         "EXIT_pwp": ["102434"], "EXIT_ontras": ["102435"], "EXIT_ua": ["102437"], "EXIT_cz": ["300001"]}
CODE_TO_COL = {z: c for c, zs in ZONES.items() for z in zs}


def num(v):
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))        # "646 036 768" (thin/no-break spaces as thousands separators)
    return float("nan") if not s or s == "-" else float(s.replace(",", ""))


def fetch(d0, d1):
    rec = {}
    s = d0
    while s <= d1:
        e = min(s + timedelta(days=30), d1)
        start, got = 0, 0
        while True:
            body = {"start": start, "length": 2000, "globalFilter": {"value": ""}, "lang": "en",
                    "sort": [{"column": "gasDay", "order": "asc"}, {"column": "zoneCode", "order": "asc"}],
                    "customFilters": {"filters": [{"filterName": "gasDay$from", "values": s.isoformat()},
                                                  {"filterName": "gasDay$to", "values": e.isoformat()}]}}
            for i in range(4):
                try:
                    r = requests.post(GS, json=body, headers={"User-Agent": UA}, timeout=(15, 120))
                    if r.ok:
                        break
                except requests.RequestException:
                    pass
                time.sleep(5 * (i + 1))
            r.raise_for_status()
            j = r.json()
            rows = j.get("data", [])
            for x in rows:
                col = CODE_TO_COL.get(str(x.get("zoneCode")))
                if col:
                    day = rec.setdefault(pd.Timestamp(x["gasDay"]), {})
                    day[(col, str(x["zoneCode"]))] = num(x.get("realization"))
            got += len(rows)
            if not rows or got >= int(j.get("total") or 0):
                break
            start += len(rows)
        print(f"  {s} .. {e}: {got} rows", flush=True)
        s = e + timedelta(days=1)
    out = {}
    for day, vals in rec.items():
        row = {}
        for (col, _), v in vals.items():
            row[col] = (row.get(col, 0.0) if row.get(col) == row.get(col) else 0.0) + (0.0 if v != v else v)
        out[day] = row
    d = pd.DataFrame.from_dict(out, orient="index").sort_index() / 1e6      # kWh -> GWh
    return d.reindex(columns=list(ZONES))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = pd.DataFrame()
    if os.path.exists(path):
        old = pd.read_excel(path, sheet_name="Daily", index_col=0, parse_dates=True)
    fs = start if old.empty else max(start, old.dropna(how="all").index.max().date() - timedelta(days=RELOAD_DAYS))
    # fill gaps inside the history too
    if not old.empty:
        full = pd.date_range(start, old.index.max())
        miss = full.difference(old.dropna(how="all").index)
        if len(miss):
            fs = min(fs, miss.min().date())
    print(f"Gaz-System MIR from {fs}", flush=True)
    new = fetch(fs, today)
    comb = old.reindex(old.index.union(new.index)) if not old.empty else new.copy()
    comb.loc[new.index, new.columns] = new
    comb = comb.sort_index().round(3).dropna(how="all")
    comb.index.name = "date"
    print((comb.resample("YS").sum() / 1000).round(2).T.to_string())
    lines = ["Poland - gas entry and exit points (Gaz-System, raw operator data)", "",
             "Source", "Gaz-System Market Information Module, 'Actual quantity of gas transmitted' (KspRealization) per zone and gas day: "
             "https://swi.gaz-system.pl/mir/#/public/bil/ksp-realization. Billing values, operative values for the latest days. Free, no key.",
             "", "Units and definitions",
             "Sheet Daily: GWh per gas day (kWh in the source / 1e6). Zones: BY_wysokoje = Wysokoje 572407 and BY_tietierowka = Tietierowka 572405 "
             "(entries from Belarus: the Yamal-Europe pipeline at Kondratki, which ENTSOG does not carry); DE_pwp 172434 and DE_ontras 172435 = Germany - Poland "
             "interconnection entries (Mallnow / Lasow); UA_gcp 172437 = Ukraine entry; CZ_cieszyn 370001, CZ_branice 372414 = Czech entries; "
             "LNG = Swinoujscie 777001; PROD = domestic production entries 907029 + 907030 + 907033; STO_out / STO_in = storage entries / exits; "
             "EXIT_* = exit-direction zones at the same interconnection points (102434, 102435, 102437, 300001).",
             f"Re-fetches the last {RELOAD_DAYS} days each run plus gaps; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} days {comb.index.min():%Y-%m-%d} to {comb.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print("saved", FILE)


if __name__ == "__main__":
    main()
