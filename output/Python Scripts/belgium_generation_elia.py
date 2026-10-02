"""
Belgium daily generation by fuel from Elia (the Belgian TSO) open data, dataset ods201 "Total generation -
aggregated by fuel type" (Opendatasoft API, no key): 15-minute MW by ENTSO-E fuel type. The query sums MW per
UTC day, fuel type and resolution on the server and converts to MWh here (MW x hours per interval).
"""
import sys
from datetime import date

import pandas as pd

import europe_common as ec

URL = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/ods201/records"
HOURS = {"PT15M": 0.25, "PT30M": 0.5, "PT1H": 1.0, "PT60M": 1.0}


def fetch(start, end):
    frames = []
    for a, b in ec.month_chunks(start, end):
        rows, offset = [], 0
        where = f"datetime >= date'{a.isoformat()}' AND datetime < date'{(b + pd.Timedelta(days=1)).date().isoformat()}'"
        while True:
            j = ec.get_json(URL, {"select": "sum(generatedpower) as s", "where": where, "limit": 100, "offset": offset,
                                  "group_by": "year(datetime) as y, month(datetime) as m, day(datetime) as d, "
                                              "fueltypeentsoe, resolutioncode"})
            part = j.get("results", [])
            rows.extend(part)
            if len(part) < 100:
                break
            offset += 100
        print(f"  {a} to {b}: {len(rows)} rows", file=sys.stderr, flush=True)
        if rows:
            frames.append(pd.DataFrame(rows))
    if not frames:
        return pd.DataFrame()
    d = pd.concat(frames)
    d["date"] = pd.to_datetime(dict(year=d["y"], month=d["m"], day=d["d"])).dt.date
    d["mwh"] = pd.to_numeric(d["s"], errors="coerce") * d["resolutioncode"].map(HOURS).fillna(0.25)
    d["col"] = d["fueltypeentsoe"].map(ec.bucket)
    print("fuel types:", sorted(d["fueltypeentsoe"].dropna().unique()), file=sys.stderr)
    d = d.dropna(subset=["col"])
    wide = d.pivot_table(index="date", columns="col", values="mwh", aggfunc="sum")
    return ec.to_daily(wide)


NOTES = ec.STANDARD_NOTES + [
    "Belgium: Elia ods201 (total generation by ENTSO-E fuel type, 15-minute MW) summed to MWh per UTC day. "
    "Energy storage excluded; waste and other fuels are in Other_MWh.",
    "",
    "COVERAGE",
    f"Daily from {ec.DEFAULT_START} (or Elia's first record) to yesterday; last {ec.RELOAD_DAYS} days re-pulled each run.",
    "",
    "SOURCE",
    "Elia Open Data, ods201: https://opendata.elia.be/explore/dataset/ods201/",
]

if __name__ == "__main__":
    ec.run("belgium_power_generation_daily.xlsx", fetch, NOTES, __doc__)
