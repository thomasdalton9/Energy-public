"""
Discovery (Oct-2026): raw official sources for Puerto Rico, Jamaica and the
Dominican Republic power generation by type and gas use, from 2021.

  eia     EIA API: Puerto Rico EIA-923 generation and fuel consumption by fuel
          (electric-power-operational-data), plus any natural-gas route that
          carries Puerto Rico (imports by point of entry, consumption)
  oc      Organismo Coordinador del SENI (oc.do / oc.org.do) - daily / monthly
          operation reports, generation by fuel, fuel consumption by plant
  dogas   Dominican Republic gas: CNE, MEM, MICM fuel imports, Banco Central
  jm      Jamaica: JPS, OUR, MSET, PCJ, STATIN
  prlocal Puerto Rico local: Genera PR, LUMA, PREB (energia.pr.gov)

Prints data-file links (xls/xlsx/csv/pdf/zip/json) and pages that mention
generation / statistics / gas. Usage: python3 CARIBBEAN_POWER_GAS_DISCOVERY.py <target>
"""

print("STARTING", flush=True)

import json
import os
import re
import socket
import sys
from urllib.parse import urljoin, urlparse

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
DATA = re.compile(r"\.(xlsx?|csv|zip|pdf|ods|json)(\?|$)", re.I)
SKIP = re.compile(r"facebook|twitter|linkedin|youtube|instagram|whatsapp|mailto:|tel:|javascript:|wp-json|xmlrpc|"
                  r"\.(jpg|jpeg|png|gif|svg|webp|css|ico|mp4)(\?|$)", re.I)
KEY = os.environ.get("EIA_API_KEY")


def get(url, **kw):
    try:
        r = S.get(url, timeout=45, **kw)
        print(f"[{r.status_code}] {r.url[:180]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def dns(hosts):
    for h in hosts:
        try:
            print(h, "->", socket.gethostbyname(h), flush=True)
        except OSError as e:
            print(h, "-> no DNS", e, flush=True)


def crawl(starts, follow, budget=45, hosts=None):
    hosts = hosts or {urlparse(u).netloc.replace("www.", "") for u in starts}
    queue, seen, data = list(starts), set(), {}
    n = 0
    while queue and n < budget:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        r = get(u)
        n += 1
        if r is None or "html" not in (r.headers.get("content-type") or ""):
            continue
        t = r.text
        title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        print("   title:", re.sub(r"\s+", " ", title.group(1)).strip()[:100] if title else "", flush=True)
        for m in re.finditer(r"<iframe[^>]*src=[\"']([^\"']+)", t, re.I):
            print("   IFRAME", urljoin(r.url, m.group(1))[:200], flush=True)
        for m in re.finditer(r"(https?://[^\"' ]*(api|ashx|asmx|svc|json)[^\"' ]*)", t, re.I):
            print("   APIish", m.group(1)[:200], flush=True)
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", t, re.I | re.S):
            href, text = m.group(1).strip(), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if SKIP.search(href):
                continue
            link = urljoin(r.url, href)
            if DATA.search(link) or "powerbi" in link or "tableau" in link:
                data.setdefault(link, text[:90])
                continue
            host = urlparse(link).netloc.replace("www.", "")
            if any(host.endswith(h) for h in hosts) and re.search(follow, link + " " + text, re.I) and link not in seen:
                queue.append(link)
    print(f"-- {len(data)} data links", flush=True)
    for link, text in data.items():
        print(f"   DATA {link[:220]}  [{text}]", flush=True)
    print(f"-- queue left: {queue[:40]}", flush=True)


def eia_q(route, params):
    p = {"api_key": KEY, **params}
    r = S.get(f"https://api.eia.gov/v2/{route}", params=p, timeout=60)
    print(f"[{r.status_code}] {route} {json.dumps({k: v for k, v in params.items()})[:300]}", flush=True)
    try:
        return r.json()
    except ValueError:
        print(r.text[:500])
        return {}


def eia():
    if not KEY:
        print("no EIA_API_KEY")
        return
    route = "electricity/electric-power-operational-data/data/"
    for data_col in ["generation", "consumption-for-eg", "total-consumption", "consumption-for-eg-btu"]:
        j = eia_q(route, {"frequency": "monthly", "data[]": data_col, "facets[location][]": "PR",
                          "facets[sectorid][]": "99", "facets[fueltypeid][]": ["NG", "ALL", "COL", "RFO", "DFO", "SUN",
                                                                                "WND", "HYC", "BIO", "LFG"],
                          "start": "2025-01", "sort[0][column]": "period", "sort[0][direction]": "desc", "length": 60})
        for row in (j.get("response", {}).get("data") or [])[:24]:
            print("   ", {k: row.get(k) for k in ("period", "fueltypeid", "sectorid", data_col, f"{data_col}-units")})
        print("   total", j.get("response", {}).get("total"))
    j = eia_q(route, {"frequency": "monthly", "data[]": "generation", "facets[location][]": "PR",
                      "facets[sectorid][]": "99", "facets[fueltypeid][]": "ALL", "start": "2020-12",
                      "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 5})
    print("   first", (j.get("response", {}).get("data") or [])[:2])
    # gas consumption by sector, NG for power by sector (utility vs IPP)
    j = eia_q(route, {"frequency": "monthly", "data[]": "consumption-for-eg", "facets[location][]": "PR",
                      "facets[fueltypeid][]": "NG", "start": "2026-01", "length": 60})
    for row in (j.get("response", {}).get("data") or [])[:30]:
        print("   ", {k: row.get(k) for k in ("period", "sectorid", "sectorDescription", "consumption-for-eg",
                                               "consumption-for-eg-units")})
    # plant-level fuel consumption, PR
    j = eia_q("electricity/facility-fuel/data/", {"frequency": "monthly", "data[]": ["generation", "consumption-for-eg"],
                                                   "facets[state][]": "PR", "facets[fuel2002][]": "NG",
                                                   "start": "2026-01", "length": 60})
    for row in (j.get("response", {}).get("data") or [])[:30]:
        print("   ", {k: row.get(k) for k in ("period", "plantName", "primeMover", "generation",
                                               "consumption-for-eg", "consumption-for-eg-units")})
    # natural gas routes that might hold Puerto Rico
    for route2 in ["natural-gas/move/poe2/", "natural-gas/move/poe1/", "natural-gas/move/impc/", "natural-gas/cons/sum/",
                   "natural-gas/move/expc/", "natural-gas/move/state/"]:
        j = eia_q(route2.rstrip("/") + "/facet/duoarea/", {})
        fac = j.get("response", {}).get("facets") or []
        hits = [f for f in fac if re.search(r"puerto|pr\b|penuelas|peñuelas|san juan", json.dumps(f), re.I)]
        print(f"   {route2}: {len(fac)} duoareas; PR-ish: {hits[:10]}")
        j = eia_q(route2.rstrip("/") + "/facet/process/", {})
        print("    processes:", [(f.get("id"), f.get("name")) for f in (j.get("response", {}).get("facets") or [])][:40])
    j = eia_q("natural-gas/move/poe2/data/", {"frequency": "monthly", "data[]": "value", "start": "2025-06",
                                              "facets[process][]": "IML", "length": 500})
    names = sorted({(r.get("duoarea"), r.get("area-name"), r.get("series-description")) for r in
                    (j.get("response", {}).get("data") or [])})
    for n in names:
        print("   IML", n)
    j = eia_q("international/data/", {"frequency": "annual", "data[]": "value", "facets[countryRegionId][]": ["PRI", "JAM", "DOM"],
                                      "facets[productId][]": "26", "start": "2021", "length": 200})
    for row in (j.get("response", {}).get("data") or [])[:40]:
        print("   INTL", {k: row.get(k) for k in ("period", "countryRegionId", "productName", "activityName", "value", "unit")})


def oc():
    dns(["oc.do", "www.oc.do", "oc.org.do", "www.oc.org.do", "apps.oc.do", "sie.gob.do", "www.cne.gob.do"])
    crawl(["https://www.oc.do/", "https://www.oc.org.do/", "https://oc.do/Informes", "https://www.oc.do/Informes/Operacion",
           "https://www.oc.do/Informes/Operaci%C3%B3n-Real"],
          r"inform|operac|generac|estad|diari|mensual|combust|reporte|publica|sistema|despacho", 70,
          hosts={"oc.do", "oc.org.do"})


def dogas():
    dns(["www.cne.gob.do", "cne.gob.do", "mem.gob.do", "www.mem.gob.do", "micm.gob.do", "www.micm.gob.do",
         "www.bancentral.gov.do", "www.one.gob.do", "www.aesdominicana.com.do", "www.sie.gob.do"])
    crawl(["https://www.cne.gob.do/", "https://www.cne.gob.do/estadisticas/", "https://www.cne.gob.do/balance-energetico/"],
          r"estad|balance|gas|import|combust|hidrocarb|electric|generac|informe|publica|document", 50)
    crawl(["https://micm.gob.do/direcciones/combustibles/", "https://micm.gob.do/",
           "https://micm.gob.do/transparencia/estadisticas-institucionales/"],
          r"combust|import|estad|gas|precio|hidrocarb|natural", 40)
    crawl(["https://mem.gob.do/", "https://mem.gob.do/estadisticas/", "https://mem.gob.do/transparencia/estadisticas-institucionales/"],
          r"estad|balance|gas|import|combust|hidrocarb|electric|generac|informe|publica|document|energ", 40)
    crawl(["https://www.bancentral.gov.do/a/d/2537-sector-externo"],
          r"externo|import|petrol|comercio|serie", 20)


def jm():
    dns(["www.jpsco.com", "jpsco.com", "our.org.jm", "www.our.org.jm", "www.mset.gov.jm", "mset.gov.jm",
         "www.pcj.com", "statinja.gov.jm", "www.statinja.gov.jm", "www.gpt.gov.jm", "www.mstem.gov.jm"])
    crawl(["https://www.jpsco.com/", "https://www.jpsco.com/investors/", "https://www.jpsco.com/about-us/"],
          r"report|annual|statistic|generat|investor|quarter|financial|performance|fuel", 40)
    crawl(["https://our.org.jm/", "https://our.org.jm/publications/", "https://our.org.jm/electricity/"],
          r"electric|statistic|report|annual|generat|performance|quarter|data|publication|document|fuel", 60)
    crawl(["https://www.mset.gov.jm/", "https://www.mset.gov.jm/energy-statistics/", "https://www.mstem.gov.jm/",
           "https://www.mstem.gov.jm/energy-statistics/"],
          r"energy|statistic|report|petroleum|electric|data|publication|import|fuel|ng|lng", 50,
          hosts={"mset.gov.jm", "mstem.gov.jm"})
    crawl(["https://www.pcj.com/", "https://www.pcj.com/energy-statistics/"], r"statistic|energy|data|report|import", 30)
    crawl(["https://statinja.gov.jm/", "https://statinja.gov.jm/Trade-Econ%20Statistics/InternationalMerchandiseTrade/",
           "https://statinja.gov.jm/NationalAccounting.aspx"],
          r"trade|import|energy|electric|statistic|merchandise|production|fuel", 40)


def prlocal():
    dns(["genera-pr.com", "www.genera-pr.com", "lumapr.com", "www.lumapr.com", "energia.pr.gov", "www.aeepr.com",
         "aeepr.com", "estadisticas.pr", "indicadores.pr", "www.prepa.pr.gov"])
    crawl(["https://genera-pr.com/", "https://genera-pr.com/en/", "https://genera-pr.com/generacion"],
          r"generac|generat|data|dato|report|informe|estad|diari|daily|fuel|combust", 40)
    crawl(["https://lumapr.com/", "https://lumapr.com/sistema-electrico/", "https://lumapr.com/datos-del-sistema/",
           "https://lumapr.com/resources/", "https://lumapr.com/system-overview/"],
          r"generac|generat|data|dato|report|informe|estad|diari|daily|fuel|system|sistema|overview|resumen", 40)
    crawl(["https://energia.pr.gov/", "https://energia.pr.gov/datos/", "https://energia.pr.gov/en/data/",
           "https://energia.pr.gov/numero-de-caso/"],
          r"dato|data|generac|generat|report|informe|estad|monthly|mensual|fuel|combust|indicad", 50)
    crawl(["https://aeepr.com/es-pr/", "https://aeepr.com/es-pr/QuienesSomos/Paginas/ReporteMensual.aspx",
           "https://indicadores.pr/"],
          r"generac|generat|data|dato|report|informe|estad|mensual|monthly|fuel|combust|indicad|energ", 40,
          hosts={"aeepr.com", "indicadores.pr"})


if __name__ == "__main__":
    {"eia": eia, "oc": oc, "dogas": dogas, "jm": jm, "prlocal": prlocal}[sys.argv[1]]()
    print("DONE", flush=True)
