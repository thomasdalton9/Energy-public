"""
Third probe for the Europe master. Closes what EU_TSO_PROBE2.py left open:

  - Energy-Charts: throttled (it returns 429 after ~8 quick calls) - remaining
    countries, history depth, rate-limit headers
  - Gas TSO data pages: GTS dataport, GAZ-SYSTEM, THE, Transgaz, Enagas,
    National Gas (looks for the API behind the JS app), Terna, Elia
  - France: the real sector-level gas consumption dataset on ODRE
  - ENTSOG: which countries have distribution / final-consumer / power points
  - AGSI burst test (rate limit), Statnett history depth

Usage: python3 EU_TSO_PROBE3.py   (writes eu_tso_probe3_report.json)
"""
import json
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, timedelta
from urllib.parse import urljoin

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
}
TIMEOUT = (10, 60)
TODAY = date.today()
YDAY = TODAY - timedelta(days=1)
WEEK_AGO = TODAY - timedelta(days=8)
RESULTS = []
DATA_EXT = re.compile(r"\.(csv|xlsx?|json|zip|xml)(\?|$)", re.I)


def get(url, **kw):
    t0 = time.time()
    try:
        r = requests.get(url, headers=kw.pop("headers", HEADERS), timeout=TIMEOUT, **kw)
        return r, round(time.time() - t0, 1), None
    except requests.RequestException as e:
        return None, round(time.time() - t0, 1), f"{type(e).__name__}: {e}"[:160]


def squash(text, n=300):
    return re.sub(r"\s+", " ", text)[:n]


def probe(group, name, url, extract=None, retry429=0, **kw):
    r, secs, err = get(url, **kw)
    tries = 0
    while r is not None and r.status_code == 429 and tries < retry429:
        tries += 1
        time.sleep(int(r.headers.get("Retry-After", 20)) if str(r.headers.get("Retry-After", "")).isdigit() else 20)
        r, secs, err = get(url, **kw)
    rec = {"group": group, "name": name, "url": url, "secs": secs}
    if r is None:
        rec.update(status="ERR", detail=err)
    else:
        rec.update(status=r.status_code, bytes=len(r.content))
        detail = ""
        if extract and r.status_code == 200:
            try:
                detail = extract(r)
            except Exception as e:
                detail = f"extract failed: {type(e).__name__}: {e}"
        rec["detail"] = squash(detail or r.text, 600)
        if r.status_code == 429:
            rec["detail"] += f" | headers={ {k: v for k, v in r.headers.items() if 'rate' in k.lower() or 'retry' in k.lower()} }"
    RESULTS.append(rec)
    print(f"[{group}] {name}: {rec['status']} {rec.get('bytes', '')}B {secs}s | {rec['detail'][:520]}", flush=True)
    return r


def page_links(group, name, url, pattern=None, n=40):
    """Landing page: list hrefs that are data files or match `pattern`."""
    def ex(r):
        hrefs = sorted({urljoin(r.url, h) for h in re.findall(r'href="([^"#]+)"', r.text)})
        keep = [h for h in hrefs if DATA_EXT.search(h) or (pattern and re.search(pattern, h, re.I))]
        scripts = re.findall(r'<script[^>]+src="([^"]+)"', r.text)
        return f"links={len(hrefs)} kept({len(keep)})={keep[:n]} scripts={scripts[:4]}"
    return probe(group, name, url, extract=ex, allow_redirects=True)


def j(r):
    return r.json()


def name_of(x):
    n = x["name"]
    return n if isinstance(n, str) else n.get("en")


# ---------------------------------------------------------- Energy-Charts ---
def energy_charts():
    def ex(r):
        d = j(r)
        ts = d.get("unix_seconds", [])
        names = [name_of(x) for x in d["production_types"]]
        return (f"series={len(names)} points={len(ts)} first={time.strftime('%Y-%m-%d', time.gmtime(ts[0])) if ts else None} "
                f"last={time.strftime('%Y-%m-%d', time.gmtime(ts[-1])) if ts else None} names={names}")
    codes = ["ch", "cz", "dk", "se", "no", "fi", "pt", "gr", "hu", "ro", "bg", "ie", "gb", "hr", "sk", "si", "lt", "lv", "ee", "ua"]
    for c in codes:
        probe("ec", f"public_power {c}", f"https://api.energy-charts.info/public_power?country={c}&start={WEEK_AGO}&end={YDAY}",
              extract=ex, retry429=2)
        time.sleep(4)
    for c, d0, d1 in (("de", "2015-01-01", "2015-01-03"), ("es", "2016-01-01", "2016-01-03"), ("pl", "2018-01-01", "2018-01-03"),
                      ("nl", "2018-01-01", "2018-01-03"), ("gb", "2017-01-01", "2017-01-03")):
        probe("ec", f"history {c} {d0}", f"https://api.energy-charts.info/public_power?country={c}&start={d0}&end={d1}",
              extract=lambda r: f"points={len(j(r).get('unix_seconds', []))}", retry429=2)
        time.sleep(4)
    probe("ec", "installed_power ES", "https://api.energy-charts.info/installed_power?country=es&time_step=yearly&installation_decommission=false",
          extract=lambda r: f"series={[name_of(x) for x in j(r)['production_types']][:8]} time={j(r)['time'][:2]}..{j(r)['time'][-1]}", retry429=2)
    time.sleep(4)
    probe("ec", "price ES bidding zone", f"https://api.energy-charts.info/price?bzn=ES&start={WEEK_AGO}&end={YDAY}",
          extract=lambda r: f"points={len(j(r).get('price', []))} unit={j(r).get('unit')}", retry429=2)


# -------------------------------------------------------- gas TSO pages ----
def gas_pages():
    page_links("gas", "GTS dataport", "https://www.gasunietransportservices.nl/en/network-operations/transparency/dataport",
               pattern=r"dataport|api|download|export")
    page_links("gas", "GTS transparency", "https://www.gasunietransportservices.nl/en/network-operations/transparency",
               pattern=r"transparen|data")
    page_links("gas", "GAZ-SYSTEM data transparency", "https://www.gaz-system.pl/en/for-customers/market-information/data-transparency.html",
               pattern=r"transparen|data|dane|download")
    page_links("gas", "THE DE publications", "https://www.tradinghub.eu/en-gb/Publications/Transparency/Aggregated-consumption-data",
               pattern=r"consumption|download|publications")
    page_links("gas", "Transgaz operational data", "https://www.transgaz.ro/en/clients/operational-data",
               pattern=r"operational|flow|consum")
    page_links("gas", "Enagas energy data", "https://www.enagas.es/en/technical-management-system/energy-data/",
               pattern=r"energy-data|gas-system|demand|supply|informe|datos")
    page_links("power", "Terna download center", "https://dati.terna.it/download-center",
               pattern=r"download|generation|produzione|dataset")
    page_links("power", "Terna developer portal", "https://developer.terna.it/",
               pattern=r"api|doc|subscri")
    # National Gas portal is a JS app: fetch its bundle and pull out API URLs
    r = probe("gas", "National Gas portal html", "https://data.nationalgas.com/",
              extract=lambda r: f"scripts={re.findall(r'src=.([^ >]+.js)', r.text)[:6]}")
    if r is not None and r.status_code == 200:
        for src in re.findall(r'src="([^"]+\.js[^"]*)"', r.text)[:3]:
            u = urljoin("https://data.nationalgas.com/", src)
            rr = probe("gas", f"National Gas bundle {src}", u, extract=lambda rr: "bundle ok")
            if rr is not None and rr.status_code == 200:
                apis = sorted(set(re.findall(r'https?://[\w.\-/]*(?:api|nationalgas)[\w.\-/?=&{}$]*', rr.text)))[:25]
                print(f"   API-looking URLs in bundle: {apis}", flush=True)
                paths = sorted(set(re.findall(r'["\'](/(?:api|data|operational)[\w\-/]*)["\']', rr.text)))[:25]
                print(f"   API-looking paths in bundle: {paths}", flush=True)


# ------------------------------------------------------------------ Elia ---
def elia():
    base = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets"
    probe("power", "Elia datasets (generation)", f'{base}?limit=40&select=dataset_id,title&where=search(title,"generation")',
          extract=lambda r: str([(d["dataset_id"], d["title"][:50]) for d in j(r)["results"]]))
    probe("power", "Elia datasets (total count)", f"{base}?limit=1", extract=lambda r: f"total_count={j(r).get('total_count')}")


# ---------------------------------------------------------- France gas -----
def france_gas():
    base = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets"
    for term in ("consommation", "gaz", "secteur", "industr"):
        probe("gas", f"ODRE titles '{term}'", f'{base}?limit=100&select=dataset_id,title&where=search(title,"{term}")',
              extract=lambda r: f"n={j(r).get('total_count')} " + str([(d["dataset_id"], (d.get("title") or "")[:42]) for d in j(r)["results"]]))
    probe("gas", "ODRE conso-jour records latest", f"{base}/conso-jour-nat-eldgrd-grt/records?limit=2&order_by=date%20desc",
          extract=lambda r: f"latest={[x['date'] for x in j(r)['results']]} sectors={[x['secteur_d_activite'] for x in j(r)['results']]}")
    probe("gas", "ODRE conso-jour sector values", f"{base}/conso-jour-nat-eldgrd-grt/records?limit=0&group_by=secteur_d_activite,operateur",
          extract=lambda r: str(j(r).get("results")))


# --------------------------------------------------------------- ENTSOG ----
def entsog():
    r = probe("gas", "ENTSOG interconnections", "https://transparency.entsog.eu/api/v1/interconnections?limit=-1",
              extract=lambda r: f"n={len(j(r)['interconnections'])} keys={list(j(r)['interconnections'][0])}")
    if r is None or r.status_code != 200:
        return
    ics = j(r)["interconnections"]
    by = defaultdict(Counter)
    for i in ics:
        for side in ("to", "from"):
            t = i.get(f"{side}InfrastructureTypeLabel")
            if t in ("Distribution", "Final Consumers"):
                by[i.get(f"{side}Country") or i.get(f"{side}CountryKey") or i.get(f"{side}CountryLabel")][t] += 1
    print(f"   demand-type points by country: { {k: dict(v) for k, v in sorted(by.items(), key=lambda kv: str(kv[0]))} }", flush=True)
    fc = [(i["pointKey"], i["pointLabel"]) for i in ics
          if i.get("toInfrastructureTypeLabel") == "Final Consumers" or i.get("fromInfrastructureTypeLabel") == "Final Consumers"]
    print(f"   Final Consumers points (first 40 distinct): {sorted(set(fc))[:40]}", flush=True)
    power = sorted({lab for _, lab in fc if re.search(r"power|electric|plant|CCGT", lab, re.I)})[:30]
    print(f"   power-plant style labels: {power}", flush=True)


# ----------------------------------------------------------- misc ---------
def misc():
    ok = 0
    t0 = time.time()
    for k in range(12):
        r, _, _ = get(f"https://agsi.gie.eu/api?country=FR&from={WEEK_AGO}&to={YDAY}")
        if r is not None and r.status_code == 200:
            ok += 1
        else:
            print(f"   AGSI burst call {k} -> {None if r is None else r.status_code}", flush=True)
    print(f"[gas] AGSI burst: {ok}/12 ok in {round(time.time() - t0, 1)}s", flush=True)
    RESULTS.append({"group": "gas", "name": "AGSI burst 12 calls", "status": 200 if ok == 12 else "PARTIAL", "detail": f"{ok}/12 ok"})
    probe("power", "Statnett history from 2015", "https://driftsdata.statnett.no/restapi/ProductionConsumption/GetData?From=2015-01-01",
          extract=lambda r: f"start={j(r).get('StartPointUTC')} end={j(r).get('EndPointUTC')} tick_ms={j(r).get('PeriodTickMs')} prod_keys={[x.get('title') for x in j(r).get('Production', [])][:8]}")


def main():
    for fn in (gas_pages, elia, france_gas, entsog, misc, energy_charts):
        try:
            fn()
        except Exception as e:
            print(f"!! {fn.__name__} failed: {type(e).__name__}: {e}", file=sys.stderr)
    with open("eu_tso_probe3_report.json", "w") as f:
        json.dump(RESULTS, f, indent=1, default=str)
    print(f"\nSUMMARY: {sum(1 for r in RESULTS if r['status'] == 200)}/{len(RESULTS)} returned 200")
    for r in RESULTS:
        print(f"  {r['status']!s:>7}  {r['group']:6} {r['name']}")


if __name__ == "__main__":
    main()
