"""Probe 3: which daily/monthly country-level requests of sis-energy-derived-reanalysis return data (probe 2: every June-2020 daily request failed UrlNoDataError)? Also the form of the successor sis-energy-global-reanalysis."""
import json, os, requests
from concurrent.futures import ThreadPoolExecutor
import cdsapi
B = "https://cds.climate.copernicus.eu/api"
ID = "sis-energy-derived-reanalysis"
cl = cdsapi.Client(url=B, key=os.environ["CDS_API_KEY"], quiet=True)
cc = dict(energy_product_type="capacity_factor_ratio")
def v(var, temp, y, m=None, spatial="country_level", **kw):
    d = dict(variable=[var], spatial_aggregation=spatial, temporal_aggregation=temp, year=[str(y)], **kw)
    if m: d["month"] = [m]
    return d
V = [
 v("wind_speed_at_10m", "daily", 2020, "01"), v("wind_speed_at_10m", "daily", 2000, "01"), v("wind_speed_at_10m", "daily", 2025, "01"),
 v("wind_speed_at_10m", "monthly", 2020), v("wind_speed_at_10m", "annual", 2020),
 v("wind_power_generation_onshore", "daily", 2020, "01", **cc), v("wind_power_generation_onshore", "monthly", 2020, **cc),
 v("solar_photovoltaic_power_generation", "daily", 2020, "01", **cc), v("solar_photovoltaic_power_generation", "hourly", 2020, "01", **cc),
 v("wind_speed_at_10m", "daily", 2020), v("wind_speed_at_10m", "daily", 2020, "01", data_format="netcdf"),
]
def run(i):
    q = V[i]
    try:
        cl.retrieve(ID, q, f"/tmp/p3_{i}")
        return f"OK {i} {json.dumps(q)} {os.path.getsize(f'/tmp/p3_{i}')} {open(f'/tmp/p3_{i}','rb').read(6)}"
    except Exception as ex:
        return f"ERR {i} {json.dumps(q)} {str(ex)[:200]!r}"
with ThreadPoolExecutor(6) as ex:
    for r in ex.map(run, range(len(V))):
        print(r, flush=True)
try:
    f = requests.get(f"{B}/catalogue/v1/collections/sis-energy-global-reanalysis/constraints.json", timeout=60).json()
    for e in f[:40]:
        print(json.dumps({k: (x if len(x) < 8 else f"{x[0]}..{x[-1]} ({len(x)})") for k, x in e.items()}))
    g = requests.get(f"{B}/catalogue/v1/collections/sis-energy-global-reanalysis", timeout=60).json()
    print(g.get("description", "")[:1500])
except Exception as ex:
    print("successor", ex)
