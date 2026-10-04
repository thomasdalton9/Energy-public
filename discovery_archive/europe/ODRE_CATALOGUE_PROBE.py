"""
Probe: ODRE (Open Data Reseaux-Energies) dataset catalogue entries about gas network losses / own consumption / balance / compressor fuel /
LNG / exports, to find what French gas consumption (offtake) leaves out. Prints dataset ids and titles. Prints only.
"""
import re
import sys

import requests

B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets"
pat = re.compile(r"perte|propre|bilan|compress|gaz de service|ecart|écart|consommation|export|expédition|frontiere|frontière|interconnexion|physique|flux|stock", re.I)
off, n = 0, 0
while True:
    r = requests.get(B, params={"limit": 100, "offset": off, "where": "search(title,'gaz') OR search(title,'gas')", "select": "dataset_id,metas"}, timeout=90)
    if not r.ok:
        print("HTTP", r.status_code, r.text[:200])
        break
    res = r.json().get("results", [])
    if not res:
        break
    for d in res:
        t = (d.get("metas", {}).get("default", {}) or {}).get("title", "")
        if pat.search(t) or pat.search(d["dataset_id"]):
            print(d["dataset_id"], "|", t)
            n += 1
    off += 100
    if off > 1500:
        break
print("matched", n)
sys.exit(0)
