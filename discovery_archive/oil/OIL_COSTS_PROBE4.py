"""Oil cost curve probe 4 (Actions): EIA STEO values for the regional oil series (production, existing-production change, wells completed/drilled, rigs)."""
import os, requests
OUT = "discovery_archive/results/oil"
KEY = os.environ["EIA_API_KEY"]
IDS = ["COPRPM", "COPRBK", "COPREF", "COPRR48", "COPRAP", "COPRHA", "COPRPUS", "TOPRPM", "TOPRBK", "TOPREF", "TOPRNI", "TOPRAC", "TOPRMP", "TOPRWF", "TOPRR48", "TOPRL48",
       "COEOPPM", "COEOPBK", "COEOPEF", "CONWRPM", "CONWRBK", "CONWREF", "CONWPM", "CONWBK", "CONWEF", "NWCPM", "NWCBK", "NWCEF", "NWDBK", "NWDEF", "NWDPM", "RIGSPM", "RIGSBK", "RIGSEF", "RIGSR48", "DUCSPM", "DUCSEF"]
lines = []
for i in IDS:
    d = requests.get("https://api.eia.gov/v2/steo/data", params={"api_key": KEY, "frequency": "monthly", "data[0]": "value", "facets[seriesId][]": i, "start": "2024-01", "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 100}, timeout=60)
    print(i, d.status_code, flush=True)
    if d.status_code == 200:
        for x in d.json()["response"]["data"]:
            lines.append(f"{i}\t{x['period']}\t{x['value']}\t{x.get('seriesDescription','')}\t{x.get('unit','')}")
open(f"{OUT}/eia_steo_regional_oil.txt", "w").write("\n".join(lines))
# last release date of STEO
r = requests.get("https://api.eia.gov/v2/steo", params={"api_key": KEY}, timeout=60)
open(f"{OUT}/eia_steo_meta.txt", "w").write(r.text[:3000])
