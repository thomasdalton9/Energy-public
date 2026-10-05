"""Probe: does EIA API v2 operating-generator-capacity carry a balancing-authority field/facet?
Needs EIA_API_KEY (Actions only). Prints metadata, facets, sample rows and BA-code counts for the latest month."""
import os, requests, collections
URL = "https://api.eia.gov/v2/electricity/operating-generator-capacity/"
key = os.environ["EIA_API_KEY"]
m = requests.get(URL, params={"api_key": key}, timeout=60).json()["response"]
print("endPeriod", m.get("endPeriod"))
print("facets:", [f["id"] for f in m.get("facets", [])])
print("data cols:", list(m.get("data", {}).keys()))
for f in m.get("facets", []):
    if "balanc" in f["id"].lower() or "ba" == f["id"].lower():
        r = requests.get(URL + "facet/" + f["id"], params={"api_key": key}, timeout=60).json()["response"]
        print(f["id"], r.get("totalFacets"), [x["id"] for x in r.get("facets", [])][:200])
end = m["endPeriod"]
d = requests.get(URL + "data/", params={"api_key": key, "frequency": "monthly", "data[0]": "net-summer-capacity-mw",
                 "start": end, "end": end, "length": 5, "facets[status][]": "OP"}, timeout=120).json()["response"]
print("total", d.get("total")); 
for r in d["data"]: print(r)
