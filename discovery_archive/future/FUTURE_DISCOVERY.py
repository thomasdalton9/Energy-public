"""
One-off discovery for the "future" masters (prints only): pipeline / plan / outlook sources.
  NA   EIA-860M monthly generator workbook (Planned, Retired sheets), EIA STEO API, EIA API status facets,
       Canada Energy Regulator Energy Futures data
  ANZ  AEMO NEM Generation Information workbook, AEMO ESOO/GSOO data pages, MBIE electricity demand and
       generation scenarios (EDGS)
  SA   ANEEL SIGA (plants under construction), Chile CNE projects in construction, Colombia UPME project register
       (datos.gov.co), Argentina CAMMESA new projects
"""
import io
import os
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
KEY = os.environ.get("EIA_API_KEY", "")


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=120, **kw)
        print(f"{r.status_code} {len(r.content):>11,} B  {r.headers.get('content-type', '')[:35]:35}  "
              f"{r.url.replace(KEY, 'KEY')[:200] if KEY else r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url[:150]}", flush=True)


def links(r, pat, n=25, base=""):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    return [u if u.startswith("http") else base + u for u in hit]


def xl(content, rows=8, cols=14, sheets=8):
    try:
        x = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print("   not excel", type(e).__name__, content[:150])
        return
    print("   sheets:", x.sheet_names[:30])
    for sh in x.sheet_names[:sheets]:
        d = pd.read_excel(x, sh, header=None, nrows=rows)
        print(f"   --- {sh}\n" + d.iloc[:, :cols].to_string(max_colwidth=22)[:1800])


print("=================== NA: EIA-860M workbook")
r = get("https://www.eia.gov/electricity/data/eia860m/")
xs = links(r, r"generator\d{4}\.xlsx", 6, "https://www.eia.gov/electricity/data/eia860m/")
if xs:
    rr = get(xs[0])
    if rr is not None and rr.ok:
        xl(rr.content, rows=4, cols=30, sheets=5)

print("=================== NA: EIA STEO API")
if KEY:
    r = get("https://api.eia.gov/v2/steo/", params={"api_key": KEY})
    if r is not None and r.ok:
        print("   ", str(r.json().get("response", {}))[:800])
    for sid in ("NGHHUUS", "NGPRPUS", "NGTCPUS", "NGEXPUS_LNG", "EPEOPUS", "NGEPGEN", "ELWNPUS"):
        r = get("https://api.eia.gov/v2/steo/data/", params={"api_key": KEY, "frequency": "monthly", "data[0]": "value",
                                                               "facets[seriesId][]": sid, "sort[0][column]": "period",
                                                               "sort[0][direction]": "desc", "length": 3})
        if r is not None and r.ok:
            print("   ", sid, str(r.json().get("response", {}).get("data", []))[:400])
    r = get("https://api.eia.gov/v2/electricity/operating-generator-capacity/facet/status/", params={"api_key": KEY})
    if r is not None and r.ok:
        print("   status facets:", str(r.json().get("response", {}).get("facets", []))[:1200])
else:
    print("   no EIA_API_KEY")

print("=================== NA: Canada Energy Regulator Energy Futures")
for u in ("https://www.cer-rec.gc.ca/en/data-analysis/canada-energy-future/",
          "https://www.cer-rec.gc.ca/open/energy/energyfutures2023/",
          "https://www.cer-rec.gc.ca/en/data-analysis/canada-energy-future/2023/",
          "https://open.canada.ca/data/api/action/package_search?q=canada%20energy%20futures&rows=5"):
    r = get(u)
    if r is not None and r.ok and "json" in r.headers.get("content-type", ""):
        for p in r.json()["result"]["results"][:5]:
            print("   pkg:", p["title"][:80], [(x.get("format"), x.get("url")) for x in p.get("resources", [])][:6])
    else:
        links(r, r"\.(csv|xlsx?)(\?|$)|open/energy|data", 20, "https://www.cer-rec.gc.ca")

print("=================== ANZ: AEMO Generation Information")
for u in ("https://aemo.com.au/en/energy-systems/electricity/national-electricity-market-nem/nem-forecasting-and-planning/forecasting-and-planning-data/generation-information",
          "https://www.aemo.com.au/en/energy-systems/electricity/national-electricity-market-nem/nem-forecasting-and-planning/forecasting-and-planning-data/generation-information"):
    links(get(u), r"generation[-_]information.*\.xlsx|\.xlsx", 15, "https://aemo.com.au")
for y, mm in ((2026, "jul"), (2026, "july"), (2026, "apr"), (2026, "april"), (2026, "jan"), (2026, "january"),
              (2025, "oct"), (2025, "october")):
    for pat in (f"https://www.aemo.com.au/-/media/files/electricity/nem/planning_and_forecasting/generation_information/{y}/nem-generation-information-{mm}-{y}.xlsx",
                f"https://aemo.com.au/-/media/files/electricity/nem/planning_and_forecasting/generation_information/{y}/nem-generation-information-{mm}-{y}.xlsx"):
        r = get(pat)
        if r is not None and r.ok and r.content[:2] == b"PK":
            xl(r.content, rows=6, cols=16, sheets=12)
            break
    else:
        continue
    break

print("=================== ANZ: AEMO ESOO / GSOO / ISP pages")
for u in ("https://aemo.com.au/en/energy-systems/electricity/national-electricity-market-nem/nem-forecasting-and-planning/forecasting-and-reliability/nem-electricity-statement-of-opportunities-esoo",
          "https://aemo.com.au/en/energy-systems/gas/gas-forecasting-and-planning/gas-statement-of-opportunities-gsoo",
          "https://aemo.com.au/energy-systems/major-publications/integrated-system-plan-isp",
          "https://forecasting.aemo.com.au/"):
    links(get(u), r"\.(xlsx?|csv|zip)(\?|$)", 15)

print("=================== ANZ: MBIE EDGS")
r = get("https://www.mbie.govt.nz/building-and-energy/energy-and-natural-resources/energy-statistics-and-modelling/energy-modelling/electricity-demand-and-generation-scenarios/")
for u in links(r, r"\.(xlsx?|csv)(\?|$)", 15, "https://www.mbie.govt.nz")[:2]:
    rr = get(u)
    if rr is not None and rr.ok:
        xl(rr.content, rows=6, cols=12, sheets=6)

print("=================== SA: ANEEL SIGA phases")
r = get("https://dadosabertos.aneel.gov.br/api/3/action/package_search?q=siga&rows=5")
if r is not None and r.ok:
    for p in r.json()["result"]["results"][:5]:
        print("   pkg:", p["name"], [(x.get("format"), x.get("url")) for x in p.get("resources", [])][:4])

print("=================== SA: Chile CNE / Coordinador projects in construction")
for u in ("https://www.cne.cl/estadisticas/electricidad/", "https://www.cne.cl/tarificacion/electricidad/",
          "https://www.coordinador.cl/desarrollo/documentos/proyectos-en-construccion/"):
    links(get(u), r"construcci|proyecto|\.xlsx", 20)

print("=================== SA: Colombia UPME project register (datos.gov.co)")
r = get("https://www.datos.gov.co/api/catalog/v1?q=registro%20proyectos%20generacion&limit=8")
if r is not None and r.ok:
    for x in r.json().get("results", [])[:8]:
        res = x.get("resource", {})
        print("   ", res.get("id"), res.get("name", "")[:90], res.get("updatedAt"))

print("=================== SA: Argentina CAMMESA / Secretaría de Energía new projects")
for u in ("https://cammesaweb.cammesa.com/informe-sintesis-mensual/",
          "https://datos.gob.ar/api/3/action/package_search?q=proyectos%20generacion&rows=5"):
    r = get(u)
    if r is not None and r.ok and "json" in r.headers.get("content-type", ""):
        for p in r.json()["result"]["results"][:5]:
            print("   pkg:", p["title"][:80], [(x.get("format"), x.get("url")) for x in p.get("resources", [])][:4])
    else:
        links(r, r"proyecto|ingreso|nuevo", 15)
