"""
One-off probe: hourly WEM (Perth / SWIS) generation mix for one calendar day,
including battery charging (negative) and discharging, from AEMO's WEMDE
facility SCADA (MWh per 5-minute dispatch interval per facility).

A WEMDE file FacilityScada_YYYYMMDD.zip covers the trading day 08:00 to 08:00
AWST, so a calendar day needs that day's file and the previous day's.
Fuels come from AU_WEM_GENERATION.fuel_of (facility-code rules). Utility-scale
only: behind-the-meter rooftop solar is not in facility SCADA.

Prints the hourly table (mean MW per fuel, AWST hour starting) as CSV to stdout.

Usage: python3 WEM_INTRADAY_DAY_PROBE.py [YYYY-MM-DD]   (default 2026-01-14, a Wednesday)
"""
import io
import json
import os
import sys
import zipfile

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "australia_nz"))
sys.path.insert(0, ROOT)
from AU_WEM_GENERATION import NEW, fuel_of, get  # noqa: E402

day = pd.Timestamp(sys.argv[1] if len(sys.argv) > 1 else "2026-01-14")
frames = []
for d in (day - pd.Timedelta(days=1), day):
    r = get(NEW.format(d=d.strftime("%Y%m%d")), ok404=True)
    if r is None:
        print(f"missing file for {d.date()}")
        continue
    z = zipfile.ZipFile(io.BytesIO(r.content))
    rows = json.loads(z.read(z.namelist()[0]))["data"]["facilityScadaDispatchIntervals"]
    frames.append(pd.DataFrame(rows))
d = pd.concat(frames, ignore_index=True)
print("columns:", list(d.columns))
print(d.head(3).to_string())

tcol = next(c for c in d.columns if "interval" in c.lower() or "time" in c.lower())
d["t"] = pd.to_datetime(d[tcol], errors="coerce")
if d["t"].dt.tz is not None:
    d["t"] = d["t"].dt.tz_convert("Australia/Perth").dt.tz_localize(None)
d = d[(d["t"] >= day) & (d["t"] < day + pd.Timedelta(days=1))].copy()
d["mw"] = pd.to_numeric(d["quantity"], errors="coerce") * 12   # MWh per 5 min -> MW
d["fuel"] = d["code"].map(fuel_of)
bat = d[d["fuel"] == "Battery_discharge"]
print("battery codes:", sorted(bat["code"].unique()))
print("battery MW min/max:", bat["mw"].min(), bat["mw"].max())
d.loc[d["fuel"] == "Battery_discharge", "fuel"] = "Battery"

five = d.pivot_table(index="t", columns="fuel", values="mw", aggfunc="sum").fillna(0)
print("5-min intervals:", len(five))
hourly = five.groupby(five.index.floor("h")).mean().round(1)
hourly.index = hourly.index.strftime("%H:%M")
print("=== HOURLY CSV START ===")
print(hourly.to_csv())
print("=== HOURLY CSV END ===")
