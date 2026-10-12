"""Probe: Eurostat Urban Audit city populations (urb_cpop1, DE1001V) + GISCO URAU label points. Writes raw results."""
import json, os, sys, requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "degree_days")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0"}
def get(url, **kw):
    r = requests.get(url, headers=UA, timeout=180, **kw)
    print(url[:120], r.status_code, len(r.content), flush=True)
    return r
r = get("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/urb_cpop1", params={"indic_ur": "DE1001V", "format": "JSON", "lang": "en"})
if r.ok:
    open(os.path.join(OUT, "urb_cpop1.json"), "wb").write(r.content)
for yr in ("2021", "2020", "2024"):
    r = get(f"https://gisco-services.ec.europa.eu/distribution/v2/urau/geojson/URAU_LB_{yr}_4326.geojson")
    if r.ok:
        open(os.path.join(OUT, f"urau_lb_{yr}.geojson"), "wb").write(r.content)
r = get("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/demo_pjan", params={"time": "2025", "sex": "T", "age": "TOTAL", "format": "JSON", "lang": "en"})
if r.ok:
    open(os.path.join(OUT, "demo_pjan_2025.json"), "wb").write(r.content)
r = get("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/demo_pjan", params={"time": "2024", "sex": "T", "age": "TOTAL", "format": "JSON", "lang": "en"})
if r.ok:
    open(os.path.join(OUT, "demo_pjan_2024.json"), "wb").write(r.content)
