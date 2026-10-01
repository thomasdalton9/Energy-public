"""
Round 6 (Oct-2026): download the CNE data-repository plant files
("Generacion y Consumo en Estacion <year>.xlsx": monthly generation and fuel
consumption by power station) so their layout can be inspected offline.
The repository (https://datacne.gob.do/repositorio-estadistico) lists a Dropbox
folder through POST /api/dropbox/list {path}; POST /api/dropbox/preview {path}
returns a temporary download link.
Usage: python3 CARIBBEAN_POWER_GAS_PROBE6.py do6
"""

print("STARTING", flush=True)

import os
import re
import sys

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)
API = "https://datacne.gob.do/api/dropbox/"


def do6():
    r = S.post(API + "list", json={"path": "/generación y consumo por central"}, timeout=60)
    files = [f for f in r.json()["files"] if f["type"] == "file"]
    for f in files:
        if not re.search(r"20(21|25|26)", f["name"]):
            continue
        p = S.post(API + "preview", json={"path": f["path_lower"]}, timeout=60)
        link = p.json().get("link")
        print(f["name"], p.status_code, bool(link), flush=True)
        if link:
            d = S.get(link, timeout=120)
            print("  ", d.status_code, len(d.content), flush=True)
            with open(os.path.join(OUT, re.sub(r"[^\w.]+", "_", f["name"])), "wb") as fh:
                fh.write(d.content)
    r = S.post(API + "list", json={"path": "/precios de energia"}, timeout=60)
    for f in r.json().get("files", []):
        print("precios:", f.get("type"), f.get("path_display"), f.get("modifiedAt"))


if __name__ == "__main__":
    {"do6": do6}[sys.argv[1]]()
    print("DONE", flush=True)
