"""
Great Britain daily generation by fuel from NESO (National Energy System Operator, the GB TSO), Historic
Generation Mix dataset on the NESO Data Portal (CKAN datastore, no key). Half-hourly MW by fuel; the daily MWh
is the day's mean MW x 24 (days with fewer than 40 of 48 periods are skipped as incomplete).

WIND is transmission-metered wind, WIND_EMB is NESO's estimate of embedded (distribution-connected) wind; both
are counted. SOLAR is NESO's embedded-solar estimate. IMPORTS and STORAGE are not generation and are left out.
GB left ENTSO-E's platform after Brexit, so NESO is the only TSO feed for it.
"""
import sys
from datetime import date, datetime

import pandas as pd

import europe_common as ec

URL = "https://api.neso.energy/api/3/action/datastore_search"
RESOURCE = "f93d1835-75bc-43e5-84ad-12472b180a98"
MAP = {"GAS": "Gas_MWh", "COAL": "Coal_MWh", "NUCLEAR": "Nuclear_MWh", "WIND": "Wind_MWh", "WIND_EMB": "Wind_MWh",
       "HYDRO": "Hydro_MWh", "BIOMASS": "Bioenergy_MWh", "SOLAR": "Solar_MWh", "OTHER": "Other_MWh"}
PAGE = 10000


def fetch(start, end):
    rows, offset, oldest = [], 0, None
    while True:
        j = ec.get_json(URL, {"resource_id": RESOURCE, "limit": PAGE, "offset": offset, "sort": "DATETIME desc"})
        recs = j["result"]["records"]
        if offset == 0:
            print("fields:", [f["id"] for f in j["result"]["fields"]], file=sys.stderr)
        if not recs:
            break
        rows.extend(recs)
        oldest = recs[-1]["DATETIME"]
        print(f"  offset {offset}: down to {oldest}", file=sys.stderr, flush=True)
        if pd.Timestamp(oldest).date() < start:
            break
        offset += PAGE
    d = pd.DataFrame(rows)
    d["date"] = pd.to_datetime(d["DATETIME"]).dt.date
    d = d[(d["date"] >= start) & (d["date"] <= end)]
    cols = [c for c in MAP if c in d.columns]
    missing = [c for c in MAP if c not in d.columns]
    if missing:
        print("columns not in the dataset:", missing, file=sys.stderr)
    for c in cols:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    g = d.groupby("date")
    count = g["DATETIME"].count()
    mean = g[cols].mean() * 24.0
    mean = mean[count >= 40]
    out = pd.DataFrame(index=mean.index)
    for c in cols:
        out[MAP[c]] = out.get(MAP[c], 0) + mean[c].fillna(0)
    return ec.to_daily(out)


NOTES = ec.STANDARD_NOTES + [
    "GB: NESO's Historic Generation Mix. Half-hourly MW by fuel -> daily MWh (mean MW x 24). Wind = metered + "
    "embedded estimate; Solar = embedded estimate; Bioenergy = biomass. Imports and storage excluded.",
    "",
    "COVERAGE",
    f"Daily from {ec.DEFAULT_START} to yesterday; the last {ec.RELOAD_DAYS} days are re-pulled every run.",
    "",
    "SOURCE",
    "NESO Data Portal, Historic GB Generation Mix: https://www.neso.energy/data-portal/historic-generation-mix",
]

if __name__ == "__main__":
    ec.run("great_britain_power_generation_daily.xlsx", fetch, NOTES, __doc__)
