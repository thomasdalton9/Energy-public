"""Probe 2: compact summary of the dataset form/constraints, then ONE small request (daily country-level, 2 variables, one month) to see licence status, file format and variable/column names."""
import json, os, sys, zipfile, requests
B = "https://cds.climate.copernicus.eu/api"
ID = "sis-energy-derived-reanalysis"
c = requests.get(f"{B}/catalogue/v1/collections/{ID}/constraints.json", timeout=60).json()
for e in c:
    d = {k: (v if len(v) < 8 else f"{v[0]}..{v[-1]} ({len(v)})") for k, v in e.items()}
    print(json.dumps(d))
f = requests.get(f"{B}/catalogue/v1/collections/{ID}/form.json", timeout=60).json()
for w in f:
    det = w.get("details", {})
    print("WIDGET", w.get("name"), w.get("type"), w.get("label"), "|", json.dumps(det.get("values") or det.get("groups") or det.get("text") or det)[:700])
col = requests.get(f"{B}/catalogue/v1/collections/{ID}", timeout=60).json()
print("licences:", json.dumps(col.get("licences") or col.get("links"))[:600])
for k in ("description", "update_frequency", "temporal_coverage"):
    print(k, str(col.get(k))[:800])
import cdsapi
cl = cdsapi.Client(url=B, key=os.environ["CDS_API_KEY"], quiet=False)
base = {"spatial_aggregation": "country_level", "temporal_aggregation": "daily"}
V = [
 dict(base, variable=["solar_photovoltaic_power_generation"], energy_product_type="capacity_factor_ratio", year=["2020"], month=["06"]),
 dict(base, variable=["wind_power_generation_onshore"], energy_product_type="capacity_factor_ratio", year=["2020"], month=["06"], data_format="netcdf"),
 dict(base, variable=["wind_speed_at_10m"], year=["2020"], month=["06"]),
 dict(base, variable=["surface_downwelling_shortwave_radiation"], year=["2020"], month=["06"]),
 dict(base, variable=["2m_air_temperature"], year=["2020"], month=["06"], data_format="csv"),
 dict(spatial_aggregation="maritime_country_level", temporal_aggregation="daily", variable=["wind_power_generation_offshore"], energy_product_type="capacity_factor_ratio", year=["2020"], month=["06"]),
]
ok = None
for i, variant in enumerate(V):
    try:
        t = cl.retrieve(ID, variant, f"/tmp/c3s_test{i}")
        print("OK", i, variant, t, os.path.getsize(f"/tmp/c3s_test{i}"))
        ok = ok or f"/tmp/c3s_test{i}"
    except Exception as ex:
        print("ERR", i, variant, type(ex).__name__, str(ex)[:600])
if not ok:
    sys.exit(0)
p = ok
print(os.path.getsize(p), open(p, "rb").read(8))
if zipfile.is_zipfile(p):
    z = zipfile.ZipFile(p); print(z.namelist())
    z.extractall("/tmp/c3s_x"); files = [os.path.join("/tmp/c3s_x", n) for n in z.namelist()]
else:
    files = [p]
os.system("pip -q install xarray netcdf4 h5netcdf >/dev/null 2>&1")
import xarray as xr
for fn in files:
    ds = xr.open_dataset(fn)
    print(fn); print(ds)
    for v in ds.data_vars:
        print(v, ds[v].attrs)
    print({k: ds[k].values[:40].tolist() for k in ds.coords if ds[k].size < 400 and k != "time"})
    print(ds.to_dataframe().head(5).to_string())
