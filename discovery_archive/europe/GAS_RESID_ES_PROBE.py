"""Probe (EU27+UK gas residual, Spain): save Enagas monthly bulletin PDFs (full balance tables), the raw Enagas demand JSON, links on Enagas energy-data
pages, and raw ALSI country+facility records for ES/IT/FR/GB. Writes into discovery_archive/europe/results/residual/."""
import json, os, re, sys, time
from datetime import date
import requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
def get(url, **kw):
    kw.setdefault("timeout", 60); h = dict(UA); h.update(kw.pop("headers", {}))
    r = requests.get(url, headers=h, **kw); print(r.status_code, url[:110], len(r.content), flush=True); return r
# 1. bulletins
PAGE = "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"
files = {}
today = date.today()
for y, m in [(2026, k) for k in range(1, 11)] + [(2025, k) for k in range(8, 13)] + [(2024, 12), (2023, 12)]:
    try:
        r = get(PAGE, params={"category": "", "month": m, "year": y})
        for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
            files.setdefault(x, (y, m))
    except Exception as e:
        print("list fail", y, m, type(e).__name__)
print("bulletins", len(files)); json.dump(files, open(os.path.join(OUT, "enagas_bulletin_list.json"), "w"), indent=1)
for f, ym in sorted(files.items(), key=lambda kv: kv[1])[-4:] + [kv for kv in files.items() if kv[1] in ((2024, 12), (2023, 12))][:2]:
    try:
        c = get("https://www.enagas.es" + f).content
        open(os.path.join(OUT, "enagas_" + f.split("/")[-1]), "wb").write(c)
    except Exception as e:
        print("pdf fail", f, type(e).__name__)
# 2. Enagas JSON raw
h = {"Accept": "application/json, text/javascript, */*; q=0.01", "X-Requested-With": "XMLHttpRequest",
     "Referer": "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/"}
U = ("https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/demanda/historico/jcr:content/responsiveGrid/"
     "container_copy_19796/realdemand_copy_copy.realdemand.json")
try:
    j = get(U, params={"date": "30/06/2026"}, headers=h).json()
    json.dump(j, open(os.path.join(OUT, "enagas_demand_raw.json"), "w"))
    print("top keys", list(j)[:20]); a = j.get("actual", [])
    print("n actual", len(a), a[:2])
except Exception as e:
    print("json fail", type(e).__name__, e)
# 3. links on energy-data pages
links = {}
for u in ["https://www.enagas.es/en/technical-management-system/energy-data/", "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/",
          "https://www.enagas.es/en/technical-management-system/energy-data/publications/", "https://www.enagas.es/en/technical-management-system/energy-data/supply/",
          "https://www.enagas.es/en/technical-management-system/energy-data/stocks/", "https://www.enagas.es/en/technical-management-system/energy-data/balance/"]:
    try:
        r = get(u)
        links[u] = sorted(set(re.findall(r'href="([^"]*energy-data[^"]*)"', r.text)))[:80]
    except Exception as e:
        links[u] = [type(e).__name__]
json.dump(links, open(os.path.join(OUT, "enagas_links.json"), "w"), indent=1)
# 4. ALSI raw
key = os.environ.get("GIE_API_KEY", "")
if key:
    raw = {}
    for cc in ("ES", "IT", "FR", "GB"):
        try:
            r = requests.get("https://alsi.gie.eu/api", params={"country": cc, "from": "2025-06-01", "to": "2025-06-03", "size": 50}, headers=dict(UA, **{"x-key": key}), timeout=60)
            raw[cc] = r.json(); print("ALSI", cc, r.status_code, str(r.json())[:300])
        except Exception as e:
            raw[cc] = str(e)
    json.dump(raw, open(os.path.join(OUT, "alsi_raw.json"), "w"), indent=1)
