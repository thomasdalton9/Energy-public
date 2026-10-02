"""One-off probe: what US->Mexico pipeline capacity / flow data is reachable from GitHub Actions.
Prints status, size and a preview for each candidate. Output: discovery_archive/americas/mexico_border_probe.txt"""
import os, sys, json
import requests

UA = {"User-Agent": "Mozilla/5.0 (energy-data research)"}
out = []
def log(*a):
    s = " ".join(str(x) for x in a); print(s); out.append(s)

def probe(name, url, params=None, n=600):
    try:
        r = requests.get(url, params=params, headers=UA, timeout=60)
        log(f"## {name}\n{url}\n-> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
        log(r.text[:n].replace("\n", " ") if "text" in (r.headers.get("content-type") or "") or "json" in (r.headers.get("content-type") or "") else "(binary)")
        return r
    except Exception as e:
        log(f"## {name}\n{url}\n-> ERROR {type(e).__name__}: {str(e)[:150]}")

key = os.environ.get("EIA_API_KEY", "")
# EIA API v2: monthly exports by point of exit (pipeline) - Mexico crossings
r = probe("EIA API v2 exports by point of exit", "https://api.eia.gov/v2/natural-gas/move/poe2/data/",
          {"api_key": key, "frequency": "monthly", "data[0]": "value", "facets[duoarea][]": "NUS-NMX",
           "sort[0][column]": "period", "sort[0][direction]": "desc", "length": 40}, 1500)
r = probe("EIA API v2 poe2 facets", "https://api.eia.gov/v2/natural-gas/move/poe2/", {"api_key": key}, 1500)
# EIA border crossing capacity reference
for n, u in [
    ("EIA pipelines page", "https://www.eia.gov/naturalgas/pipelines/EIA-NaturalGasPipelines.xlsx"),
    ("EIA border crossing xls", "https://www.eia.gov/naturalgas/pipelines/bordercrossing.xlsx"),
    ("EIA pipelines overview", "https://www.eia.gov/naturalgas/data.php"),
    ("EIA US-Mexico crossings (pipeline projects)", "https://www.eia.gov/naturalgas/pipelines/EIA-NaturalGasPipelines.xlsx"),
    ("CENAGAS home", "https://www.gob.mx/cenagas"),
    ("CENAGAS boletin electronico", "https://www.cenagas.gob.mx/"),
    ("CENAGAS SISTRANGAS flujos", "https://www.gob.mx/cenagas/acciones-y-programas/sistema-de-informacion-de-operacion-del-sistrangas"),
    ("Kinder Morgan EPNG ebb", "https://pipeline2.kindermorgan.com/Capacity/OpAvailEntryDate.aspx?code=EPNG"),
    ("Energy Transfer Trans-Pecos", "https://tgpl.kindermorgan.com/"),
    ("EIA Weekly Update", "https://www.eia.gov/naturalgas/weekly/"),
]:
    probe(n, u, n=300)
open(os.path.join(os.path.dirname(__file__), "mexico_border_probe.txt"), "w").write("\n".join(out))
