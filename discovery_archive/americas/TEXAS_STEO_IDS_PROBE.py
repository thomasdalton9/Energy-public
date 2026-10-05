"""Probe (Actions only): EIA STEO series ids for regional dry gas production, with values around now."""
import os, re, requests
KEY = os.environ.get("EIA_API_KEY", "")
def eia(route, params=None):
    r = requests.get("https://api.eia.gov/v2/" + route, params={"api_key": KEY, **(params or {})}, timeout=(10, 90))
    print(route, r.status_code)
    r.raise_for_status()
    return r.json()["response"]
fac = eia("steo/facet/seriesId", {"length": 5000})["facets"]
print(len(fac), "series")
hits = [f for f in fac if re.search(r"permian|eagle|haynes|marcellus|bakken|anadarko|niobrara|appalach|rest of|lower 48|NGPR|NGMP|ngnw", (f["id"] + str(f.get("name", ""))).lower())]
for f in hits:
    print("  ", f["id"], "|", f.get("name"))
for f in hits:
    if re.search(r"permian|eagle|haynes|marcellus|bakken|NGPR", (f["id"] + str(f.get("name", ""))).lower()):
        d = eia("steo/data/", {"frequency": "monthly", "data[0]": "value", "facets[seriesId][]": f["id"], "start": "2026-01",
                                "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 60})["data"]
        print(f["id"], len(d), d[0]["period"] if d else None, d[-1]["period"] if d else None, d[0].get("unit") if d else None,
              [(r["period"], r["value"]) for r in d[::6]])
