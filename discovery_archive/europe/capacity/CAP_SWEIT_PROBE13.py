"""Probe 13: Terna download-centre recordsv2 (POST JSON) for CapacityRenewableSources / CapacityThermal / GenerationPlants - granularity of the capacity datasets.
Probe 12 finding: the page POSTs a JSON body to /api/sitecore/dati/downloadcenter/recordsv2 (and downloads from .../downloadfile?)."""
import json, subprocess
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def post(u, body, n=2500):
    cmd = ["curl", "-sL", "-m", "120", "-A", UA, "-H", "Content-Type: application/json", "-H", "Accept: application/json", "-X", "POST", "--data", json.dumps(body),
           "-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=150)
    t = o.stdout
    print("\n=====", u, json.dumps(body)[:200], "|", t[t.rfind("#HTTP"):].strip(), flush=True)
    print(t[:t.rfind("#HTTP")][:n].replace("\n", " "), flush=True)

API = "https://dati.terna.it/api/sitecore/dati/downloadcenter/recordsv2"
for ds in ["CapacityRenewableSources", "CapacityThermal", "GenerationPlants"]:
    for extra in [{"filterYear": "2025"}, {"filterYear": "2024", "filterMonth": "12"}]:
        body = {"filterDataset": ds, **extra, "orderByColumn": "Anno", "orderByDir": "desc", "pageSize": "6", "pageIndex": "0", "db": "dati"}
        post(API, body)
