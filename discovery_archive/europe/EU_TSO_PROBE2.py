"""
Second probe for the Europe master (follow-up to EU_TSO_POWER_GAS_PROBE.py).
Fixes the first run's failures and goes deeper on what worked:

  - PSE (Poland): print the real OData error / entity names
  - NESO (GB) and TenneT (NL): try current hostnames
  - GTS / Snam / GAZ-SYSTEM / National Gas / THE / Gassco / Enagas / Terna:
    scan landing pages for data-looking links instead of guessing deep URLs
  - GIE AGSI+/ALSI: history depth and payload without a key
  - ENTSOG: interconnection-point categories (distribution / industrial /
    power) as a gas-demand proxy, and the aggregated endpoint spelling
  - Energy-Charts: which countries, history depth, capacity and price endpoints
  - ODRE France gas: contents of the national consumption-by-sector datasets
  - Eurostat monthly electricity / gas as an accuracy benchmark

Usage: python3 EU_TSO_PROBE2.py   (writes eu_tso_probe2_report.json)
"""
import json
import re
import sys
import time
from collections import Counter
from datetime import date, timedelta

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
ENTITY_RE = re.compile(r"EntitySet Name=.([\w-]+).")
LINK_WORDS = ("api", "csv", "xls", "json", "download", "transparen", "open-data", "opendata",
              "data", "dataset", "report")


def get(url, **kw):
    t0 = time.time()
    try:
        r = requests.get(url, headers=kw.pop("headers", HEADERS), timeout=TIMEOUT, **kw)
        return r, round(time.time() - t0, 1), None
    except requests.RequestException as e:
        return None, round(time.time() - t0, 1), f"{type(e).__name__}: {e}"[:160]


def squash(text, n=240):
    return re.sub(r"\s+", " ", text)[:n]


def probe(group, name, url, extract=None, **kw):
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
        rec["detail"] = squash(detail or r.text, 500)
    RESULTS.append(rec)
    print(f"[{group}] {name}: {rec['status']} {rec.get('bytes', '')}B {secs}s | {rec['detail'][:420]}", flush=True)
    return r


def scan(group, name, url):
    """Landing page: status, <title>, and up to 15 data-looking links."""
    def ex(r):
        title = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.S | re.I)
        links = sorted({m for m in re.findall(r'href="([^"#]+)"', r.text)
                        if any(w in m.lower() for w in LINK_WORDS)})
        return f"title={squash(title.group(1), 80) if title else None} links({len(links)})={links[:15]}"
    return probe(group, name, url, extract=ex, allow_redirects=True)


def j(r):
    return r.json()


# ------------------------------------------------------------- Poland ------
def pse():
    r = probe("power", "PSE $metadata entity sets", "https://api.raporty.pse.pl/api/$metadata",
              extract=lambda r: "EntitySets=" + str(ENTITY_RE.findall(r.text)[:40]))
    probe("power", "PSE his-wlk-cal $first=1", "https://api.raporty.pse.pl/api/his-wlk-cal?$first=1",
          extract=lambda r: str(j(r))[:400])
    probe("power", "PSE his-wlk-cal filter (2018)", "https://api.raporty.pse.pl/api/his-wlk-cal?$filter=business_date eq '2018-06-01'&$first=2",
          extract=lambda r: str(j(r))[:400])
    probe("power", "PSE error text for bad property (shows property names)",
          "https://api.raporty.pse.pl/api/his-wlk-cal?$filter=doba%20eq%20'2018-06-01'",
          extract=lambda r: r.text)


# ------------------------------------------------------- GB / NL hosts ------
def hosts():
    for h in ("https://data.nationalgrideso.com/api/3/action/package_show?id=historic-generation-mix",
              "https://api.neso.energy/api/3/action/package_show?id=historic-generation-mix",
              "https://www.neso.energy/data-portal"):
        probe("power", f"NESO {h.split('/')[2]}", h,
              extract=lambda r: str([(x["name"], x["id"]) for x in j(r)["result"]["resources"]][:6])
              if "package_show" in r.url else squash(r.text, 200))
    for h in ("https://www.tennet.eu/", "https://www.tennet.eu/electricity-market/transparency-pages/",
              "https://api.tennet.eu/", "https://transparency.tennet.eu/"):
        scan("power", f"TenneT {h}", h)


# ------------------------------------------------------ landing pages -------
def landing():
    for name, url in [
        ("GTS NL root", "https://www.gasunietransportservices.nl/en"),
        ("Snam IT root", "https://www.snam.it/en/"),
        ("Snam IT transparency alt", "https://jarvis.snam.it/"),
        ("GAZ-SYSTEM PL root", "https://www.gaz-system.pl/en/"),
        ("National Gas data portal", "https://data.nationalgas.com/"),
        ("THE DE publications", "https://www.tradinghub.eu/en-gb/Publications/Transparency/Aggregated-consumption-data"),
        ("Gassco umm", "https://umm.gassco.no/"),
        ("Gassco flows", "https://www.gassco.no/en/our-activities/flows-and-capacity/"),
        ("Enagas data", "https://www.enagas.es/en/technical-management-system/energy-data/"),
        ("Terna transparency", "https://www.terna.it/en/electric-system/transparency-report/actual-generation"),
        ("Terna dataset portal", "https://dati.terna.it/"),
        ("Fluxys (BE)", "https://www.fluxys.com/en/products-services/transmission/transparency"),
        ("Elia (BE) open data", "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?limit=3"),
        ("Transgaz (RO)", "https://www.transgaz.ro/en"),
    ]:
        scan("landing", name, url)


# ------------------------------------------------------------- GIE ----------
def gie():
    for label, base, cty, d0 in [("AGSI", "https://agsi.gie.eu/api", "DE", "2011-01-01"),
                                 ("AGSI", "https://agsi.gie.eu/api", "DE", WEEK_AGO),
                                 ("ALSI", "https://alsi.gie.eu/api", "ES", "2016-01-01"),
                                 ("ALSI", "https://alsi.gie.eu/api", "ES", WEEK_AGO)]:
        probe("gas", f"{label} {cty} from {d0}", f"{base}?country={cty}&from={d0}&to={d0 if d0 != WEEK_AGO else YDAY}&size=300",
              extract=lambda r: f"total={j(r).get('total')} rows={len(j(r).get('data', []))} first={str(j(r).get('data', [None])[0])[:300]}")
    probe("gas", "AGSI EU aggregate", f"https://agsi.gie.eu/api?continent=eu&from={WEEK_AGO}&to={YDAY}",
          extract=lambda r: f"total={j(r).get('total')} first={str(j(r).get('data', [None])[0])[:300]}")
    probe("gas", "ALSI list of facilities", "https://alsi.gie.eu/api/about?show=listing",
          extract=lambda r: str(j(r))[:300])


# ------------------------------------------------------------ ENTSOG --------
def entsog():
    base = "https://transparency.entsog.eu/api/v1"
    r = probe("gas", "ENTSOG interconnections (all)", f"{base}/interconnections?limit=-1",
              extract=lambda r: f"n={len(j(r)['interconnections'])}")
    if r is not None and r.status_code == 200:
        ics = j(r)["interconnections"]
        for key in ("fromInfrastructureTypeLabel", "toInfrastructureTypeLabel"):
            print(f"   {key}: {Counter(i.get(key) for i in ics).most_common(12)}", flush=True)
        cons = [i for i in ics if any(w in str(i.get("toInfrastructureTypeLabel", "")).lower() + str(i.get("fromInfrastructureTypeLabel", "")).lower()
                                      for w in ("distribution", "final consumer", "power", "industr"))]
        print(f"   demand-type points: {len(cons)}; sample={[ (c['pointKey'], c['pointLabel']) for c in cons[:6] ]}", flush=True)
    for spelling in ("AggregatedData", "aggregateddata", "aggregatedData"):
        probe("gas", f"ENTSOG {spelling}", f"{base}/{spelling}?indicator=Physical%20Flow&periodType=day&from={WEEK_AGO}&to={YDAY}&timezone=CET&limit=3",
              extract=lambda r: str(j(r))[:300])
    probe("gas", "ENTSOG physical flow, DE distribution/power points sample",
          f"{base}/operationalData?indicator=Physical%20Flow&periodType=day&from={WEEK_AGO}&to={YDAY}&timezone=CET&limit=5&pointKey=DIS-00001",
          extract=lambda r: str([(x["pointLabel"], x["periodFrom"][:10], x["value"]) for x in j(r)["operationalData"]]))
    probe("gas", "ENTSOG 5y limit (2021-10)", f"{base}/operationalData?indicator=Physical%20Flow&periodType=day&from=2021-10-01&to=2021-10-03&timezone=CET&limit=3",
          extract=lambda r: f"rows={len(j(r)['operationalData'])}")
    probe("gas", "ENTSOG TP archive pointer", "https://transparency.entsog.eu/#/archive")


# --------------------------------------------------------- Energy-Charts ----
def energy_charts():
    codes = ["de", "fr", "es", "it", "pl", "nl", "be", "at", "ch", "cz", "dk", "se", "no", "fi", "pt", "gr", "hu", "ro", "bg", "ie", "gb", "hr", "sk", "si", "lt", "lv", "ee"]
    for c in codes:
        def ex(r, c=c):
            d = j(r)
            ts = d.get("unix_seconds", [])
            names = [x["name"] if isinstance(x["name"], str) else x["name"].get("en") for x in d["production_types"]]
            return f"series={len(names)} points={len(ts)} last={time.strftime('%Y-%m-%d', time.gmtime(ts[-1])) if ts else None} sample={names[:5]}"
        probe("ec", f"Energy-Charts public_power {c}", f"https://api.energy-charts.info/public_power?country={c}&start={WEEK_AGO}&end={YDAY}", extract=ex)
    probe("ec", "Energy-Charts history (DE 2015)", "https://api.energy-charts.info/public_power?country=de&start=2015-01-01&end=2015-01-03",
          extract=lambda r: f"points={len(j(r).get('unix_seconds', []))}")
    probe("ec", "Energy-Charts history (ES 2016)", "https://api.energy-charts.info/public_power?country=es&start=2016-01-01&end=2016-01-03",
          extract=lambda r: f"points={len(j(r).get('unix_seconds', []))}")
    probe("ec", "Energy-Charts installed_power (DE yearly)", "https://api.energy-charts.info/installed_power?country=de&time_step=yearly&installation_decommission=false",
          extract=lambda r: f"series={[x['name'] if isinstance(x['name'], str) else x['name'].get('en') for x in j(r)['production_types']][:6]} years={j(r)['time'][:3]}..{j(r)['time'][-1]}")
    probe("ec", "Energy-Charts price (DE-LU)", f"https://api.energy-charts.info/price?bzn=DE-LU&start={WEEK_AGO}&end={YDAY}",
          extract=lambda r: f"points={len(j(r).get('price', []))} unit={j(r).get('unit')}")
    probe("ec", "Energy-Charts signal/other country list", "https://api.energy-charts.info/",
          extract=lambda r: squash(r.text, 200))


# --------------------------------------------------------- France gas -------
def france_gas():
    base = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets"
    for ds in ("pg2024-conso-nat-secteur", "conso-jour-nat-eldgrd-grt"):
        probe("gas", f"ODRE {ds} meta", f"{base}/{ds}",
              extract=lambda r: f"title={j(r)['metas']['default'].get('title')} records={j(r).get('has_records')} fields={[f['name'] for f in j(r)['fields']][:18]}")
        probe("gas", f"ODRE {ds} records", f"{base}/{ds}/records?limit=3",
              extract=lambda r: f"total={j(r).get('total_count')} first={str(j(r)['results'][0])[:350]}")
    probe("gas", "ODRE gas dataset titles", f'{base}?where=search(title,"gaz")%20or%20search(title,"gas")&limit=60&select=dataset_id,title',
          extract=lambda r: str([(d["dataset_id"], (d.get("title") or "")[:45]) for d in j(r)["results"]]))


# --------------------------------------------------------------- misc -------
def misc():
    probe("bench", "Eurostat monthly electricity DE (nrg_cb_pem)",
          "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem?geo=DE&lastTimePeriod=3&format=JSON&lang=EN",
          extract=lambda r: f"keys={list(j(r))[:8]} dims={j(r).get('id')} n_values={len(j(r).get('value', {}))}")
    probe("bench", "Eurostat monthly gas DE (nrg_cb_gasm)",
          "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm?geo=DE&lastTimePeriod=3&format=JSON&lang=EN",
          extract=lambda r: f"dims={j(r).get('id')} n_values={len(j(r).get('value', {}))}")
    probe("gas", "Energinet gas datasets in catalogue", "https://api.energidataservice.dk/meta/dataset?limit=200",
          extract=lambda r: str([d.get("datasetName") or d.get("name") for d in (j(r) if isinstance(j(r), list) else j(r).get("records", [])) if "gas" in str(d).lower()][:20]))
    probe("power", "Statnett history endpoints", "https://driftsdata.statnett.no/restapi/ProductionConsumption/GetData?From=2026-09-20",
          extract=lambda r: f"keys={list(j(r))[:6]}")


def main():
    for fn in (pse, hosts, landing, gie, entsog, energy_charts, france_gas, misc):
        try:
            fn()
        except Exception as e:
            print(f"!! {fn.__name__} failed: {type(e).__name__}: {e}", file=sys.stderr)
    with open("eu_tso_probe2_report.json", "w") as f:
        json.dump(RESULTS, f, indent=1, default=str)
    print(f"\nSUMMARY: {sum(1 for r in RESULTS if r['status'] == 200)}/{len(RESULTS)} returned 200")
    for r in RESULTS:
        print(f"  {r['status']!s:>4}  {r['group']:7} {r['name']}")


if __name__ == "__main__":
    main()
