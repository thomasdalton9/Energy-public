"""
Gatun Lake (Panama Canal) water level - free, public, no key. Gatun is
the Canal's main reservoir; its level directly limits vessel draft
(and so Canal transit capacity) and is closely watched the same way
gas storage or hydro reservoirs are elsewhere in this repo.

Source: Panama Canal Authority (ACP)'s own water-level dashboard
(evtms-rpts.pancanal.com/eng/h2o/index.html), which links a plain CSV -
confirmed via GATUN_RHINE_PANAMA_DISCOVERY.py:
  DATE_LOG, GATUN_LAKE_LEVEL(FEET) - daily, back to 1965-01-01.

Outputs (gatun_lake_level.xlsx):
  Daily   date, level_ft, country ("Panama")

Usage: python3 GATUN_LAKE_LEVEL.py [--out gatun_lake_level.xlsx]
"""
print("STARTING", flush=True)

import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://evtms-rpts.pancanal.com/eng/h2o/Download_Gatun_Lake_Water_Level_History.csv"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
COUNTRY = "Panama"
OUT_DEFAULT = "gatun_lake_level.xlsx"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=OUT_DEFAULT)
    args = parser.parse_args()

    print("Downloading Gatun Lake water level history...", flush=True)
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    df.columns = ["date", "level_ft"]
    df["date"] = pd.to_datetime(df["date"])
    df["country"] = COUNTRY
    df = df.sort_values("date")

    print(f"  {len(df)} days, {df['date'].min().date()} to {df['date'].max().date()}", flush=True)
    print(df.tail().to_string(index=False), flush=True)

    notes = [
        "UNITS",
        "level_ft: Gatun Lake's surface level in feet above mean sea level, as published by ACP.",
        "",
        "WHY THIS MATTERS",
        "Gatun is the Panama Canal's main reservoir - its level directly caps vessel draft (and so how loaded "
        "ships can transit), which ACP adjusts via periodic draft restrictions when the lake runs low. Low "
        "Gatun levels (e.g. the 2023-2024 drought) directly reduced Canal transit capacity.",
        "",
        "SOURCE",
        "Panama Canal Authority (ACP), evtms-rpts.pancanal.com/eng/h2o - 'Gatun Water Level Indicators' "
        "dashboard's underlying CSV. Daily, back to 1965-01-01.",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": df}, notes, {"UNITS", "WHY THIS MATTERS", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
