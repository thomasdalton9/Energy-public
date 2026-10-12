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
req = {"variable": ["solar_photovoltaic_power_generation", "wind_power_generation_onshore"], "spatial_aggregation": "country_level",
       "temporal_aggregation": "daily", "energy_product_type": "capacity_factor_ratio", "year": ["2025"], "month": ["06"], "data_format": "netcdf"}
for variant in (req,):
    try:
        t = cl.retrieve(ID, variant, "/tmp/c3s_test")
        print("downloaded", t)
    except Exception as ex:
        print("ERR", type(ex).__name__, str(ex)[:3000])
        sys.exit(0)
p = "/tmp/c3s_test"
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
