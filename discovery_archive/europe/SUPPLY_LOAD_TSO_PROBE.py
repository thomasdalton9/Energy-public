"""
Probe: national TSO / statistics sources for load and generation (Spain REData, Belgium Elia open data, Slovenia, Serbia, Lithuania),
to compare against ENTSO-E actual total load. Prints only.
"""
import json
import re
import signal
import sys

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json, text/html, */*"}


def _timeout(*_):
    raise TimeoutError("hard timeout")


signal.signal(signal.SIGALRM, _timeout)


def get(url, n=300, **kw):
    try:
        signal.alarm(45)   # hard cap: requests' timeout is per socket read, a trickling host can hang forever
        r = requests.get(url, headers=UA, timeout=(10, 25), **kw)
        signal.alarm(0)
        print(f"GET {url[:150]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}b", flush=True)
        if n:
            print("   ", r.text[:n].replace("\n", " "), flush=True)
        return r
    except BaseException as e:  # noqa: BLE001
        signal.alarm(0)
        print(f"GET {url[:150]} -> ERR {str(e)[:100]}", flush=True)


print("=== SPAIN REData balance (MWh per year)")
for y in (2023, 2024, 2025):
    r = get(f"https://apidatos.ree.es/en/datos/balance/balance-electrico?start_date={y}-01-01T00:00&end_date={y}-12-31T23:59&time_trunc=year", 0)
    if r is not None and r.ok:
        try:
            for grp in r.json()["included"]:
                tot = grp["attributes"].get("total")
                print(f"  {y} {grp['type']:35s} total={tot}")
                for it in grp["attributes"]["content"]:
                    t = it["attributes"].get("total")
                    print(f"       {it['type'][:38]:40s} {t/1e6 if t else t:.3f} TWh")
        except Exception as e:  # noqa: BLE001
            print("  parse", e, r.text[:200])
for y in (2024,):
    get(f"https://apidatos.ree.es/en/datos/demanda/evolucion?start_date={y}-01-01T00:00&end_date={y}-12-31T23:59&time_trunc=year", 1500)
    get(f"https://apidatos.ree.es/en/datos/intercambios/enlaces-francia?start_date={y}-01-01T00:00&end_date={y}-12-31T23:59&time_trunc=year", 600)

print("\n=== BELGIUM Elia open data")
r = get("https://opendata.elia.be/api/explore/v2.1/catalog/datasets?limit=100&select=dataset_id,title&where=search(%22load%22)%20or%20search(%22generation%22)%20or%20search(%22flow%22)", 0)
if r is not None and r.ok:
    for d in r.json().get("results", []):
        print("  ", d.get("dataset_id"), "|", (d.get("title") or d.get("metas", {}).get("default", {}).get("title", ""))[:90])
for ds in ("ods001", "ods003", "ods002", "ods032", "ods177"):
    get(f"https://opendata.elia.be/api/explore/v2.1/catalog/datasets/{ds}/records?limit=1", 700)

print("\n=== SLOVENIA")
for u in ("https://www.eles.si/", "https://www.eles.si/en/transparency", "https://pxweb.stat.si/SiStatData/api/v1/en/Data/",
          "https://pxweb.stat.si/SiStatData/api/v1/en/Data/1818501S.px", "https://www.stat.si/StatWeb/en/News/Index/",
          "https://www.agen-rs.si/", "https://www.nek.si/en/about-nek/key-data"):
    get(u, 250)

print("\n=== SERBIA")
for u in ("https://ems.rs/", "https://ems.rs/en/", "https://www.ems.rs/", "https://data.stat.gov.rs/?caller=SDDB", "https://www.eps.rs/eng/Pages/default.aspx",
          "https://www.aers.rs/"):
    get(u, 250)

print("\n=== LITHUANIA")
for u in ("https://www.litgrid.eu/", "https://www.litgrid.eu/index.php/power-system/power-system-information/electricity-balance/",
          "https://www.litgrid.eu/index.php/energetikos-sistema/", "https://energy.litgrid.eu/", "https://osp.stat.gov.lt/", "https://www.ena.lt/",
          "https://api.energy-charts.info/public_power?country=lt&start=2024-01-01&end=2024-01-02"):
    r = get(u, 250)
    if r is not None and r.ok and "litgrid" in u:
        for m in sorted(set(re.findall(r'href="([^"]+)"', r.text)))[:80]:
            if re.search(r"xls|csv|balans|load|sistem|energy|data|api", m, re.I):
                print("      link", m[:140])

print("\n=== BELGIUM national alternative: Statbel / FPS economy")
get("https://statbel.fgov.be/en/themes/energy/electricity", 200)
sys.exit(0)
