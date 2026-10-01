"""
One-off: download US balancing-authority (control area) boundaries from
HIFLD's public ArcGIS service, simplify them, and save a small GeoJSON for
the North America coverage map (maps/us_balancing_authorities.geojson).
Run in GitHub Actions (services1.arcgis.com is blocked from the sandbox).
"""
import json
import os

import geopandas as gpd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URLS = [
    "https://services1.arcgis.com/Hp6G80Pky0om7QvQ/arcgis/rest/services/Control_Areas/FeatureServer/0/query",
    "https://services5.arcgis.com/HDRa0B57OVrv2E1q/arcgis/rest/services/Control_Areas/FeatureServer/0/query",
]
OUT = os.path.join(ROOT, "maps", "us_balancing_authorities.geojson")


def fetch(url):
    feats, offset = [], 0
    while True:
        r = requests.get(url, params={"where": "1=1", "outFields": "*", "f": "geojson", "outSR": 4326,
                                      "resultOffset": offset, "resultRecordCount": 200}, timeout=120)
        r.raise_for_status()
        page = r.json().get("features", [])
        feats += page
        print(f"  {len(page)} features at offset {offset}", flush=True)
        if len(page) < 200:
            return feats
        offset += 200


for url in URLS:
    try:
        print("trying", url, flush=True)
        feats = fetch(url)
        if feats:
            break
    except Exception as e:  # noqa: BLE001
        print("  failed:", type(e).__name__, e, flush=True)
else:
    raise SystemExit("no source worked")

gdf = gpd.GeoDataFrame.from_features(feats, crs=4326)
print("columns:", list(gdf.columns))
keep = [c for c in gdf.columns if c.upper() in ("NAME", "ID", "TYPE", "COUNTRY", "STATE", "WEBSITE")] + ["geometry"]
gdf = gdf[keep]
gdf["geometry"] = gdf.geometry.simplify(0.05, preserve_topology=True)
print(gdf.drop(columns="geometry").to_string())
os.makedirs(os.path.dirname(OUT), exist_ok=True)
for c in gdf.columns:
    if c != "geometry":
        gdf[c] = gdf[c].astype(object).where(gdf[c].notna(), None).map(lambda v: None if v is None else str(v))
with open(OUT, "w") as f:
    f.write(gdf.to_json())
print("saved", OUT, os.path.getsize(OUT), "bytes")
