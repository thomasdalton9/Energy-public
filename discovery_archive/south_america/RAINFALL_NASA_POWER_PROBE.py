"""Probe: does NASA POWER (power.larc.nasa.gov daily point API, PRECTOTCORR) answer from GitHub Actions for a couple of
South American catchment points over a ~8-year range? Prints status, elapsed time, number of days, missing codes,
the first/last date returned and the water-year totals."""
import datetime as dt
import time

import pandas as pd
import requests

URL = "https://power.larc.nasa.gov/api/temporal/daily/point"
POINTS = {"Guatape (Antioquia)": (6.3, -75.2), "Furnas (Grande)": (-20.7, -46.3)}
start, end = dt.date(2018, 10, 1), dt.date.today()
for name, (lat, lon) in POINTS.items():
    t0 = time.time()
    try:
        r = requests.get(URL, params={"parameters": "PRECTOTCORR", "community": "AG", "latitude": lat, "longitude": lon,
                                      "start": start.strftime("%Y%m%d"), "end": end.strftime("%Y%m%d"),
                                      "format": "JSON"}, timeout=300)
        print(name, "HTTP", r.status_code, f"{time.time() - t0:.1f}s", len(r.content), "bytes", flush=True)
        r.raise_for_status()
        js = r.json()
        p = js["properties"]["parameter"]["PRECTOTCORR"]
        s = pd.Series({pd.Timestamp(k): v for k, v in p.items()})
        miss = (s <= -998).sum()
        s = s[s > -998]
        print("  days", len(p), "missing codes", miss, "first", s.index.min().date(), "last", s.index.max().date())
        print("  header:", js.get("header", {}).get("api", {}), js.get("parameters"))
        wy = s.groupby([(s.index.year + (s.index.month >= 10)).rename("wy")]).sum().round(0)
        print("  water-year totals mm (wy = year of the Sep end):", wy.to_dict())
    except Exception as exc:  # noqa: BLE001
        print(name, "FAILED", type(exc).__name__, exc, f"{time.time() - t0:.1f}s", flush=True)
