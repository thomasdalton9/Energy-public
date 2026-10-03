"""
Probe for the planned Europe master: which national TSO / operator power and
gas feeds are open (no key), how deep their history goes, and how they compare
with the EU aggregators (ENTSO-E, ENTSOG, GIE AGSI+/ALSI).

Nothing here is a pull - each check makes one or two small requests and prints
a one-line verdict (HTTP status, size, a data snippet, history depth where it
can be read cheaply). Endpoints that need a key are probed without one: 401/403
means "reachable, key required"; a connection error means blocked/down.

Usage: python3 EU_TSO_POWER_GAS_PROBE.py
Writes eu_tso_probe_report.json next to the script's working directory.
"""
import json
import sys
import time
from datetime import date, timedelta

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
}
TIMEOUT = (10, 45)
TODAY = date.today()
YDAY = TODAY - timedelta(days=1)
WEEK_AGO = TODAY - timedelta(days=8)

RESULTS = []


def get(url, **kw):
    t0 = time.time()
    try:
        r = requests.get(url, headers=kw.pop("headers", HEADERS), timeout=TIMEOUT, **kw)
        return r, round(time.time() - t0, 1), None
    except requests.RequestException as e:
        return None, round(time.time() - t0, 1), f"{type(e).__name__}: {e}"[:200]


def probe(group, name, url, note=None, extract=None, **kw):
    """One request; extract(response) -> short string with data/history facts."""
    r, secs, err = get(url, **kw)
    rec = {"group": group, "name": name, "url": url, "secs": secs}
    if r is None:
        rec.update(status="ERR", detail=err)
    else:
        rec.update(status=r.status_code, bytes=len(r.content),
                   ctype=r.headers.get("Content-Type", "")[:40])
        detail = ""
        if extract and r.status_code == 200:
            try:
                detail = extract(r)
            except Exception as e:  # layout drift is the thing we want to see
                detail = f"extract failed: {type(e).__name__}: {e}"[:200]
        if not detail:
            detail = r.text[:160].replace("\n", " ")
        rec["detail"] = detail
    if note:
        rec["note"] = note
    RESULTS.append(rec)
    print(f"[{group}] {name}: {rec['status']} {rec.get('bytes', '')}B {secs}s | {rec['detail'][:230]}",
          flush=True)
    return r


def iso(ms):
    return time.strftime("%Y-%m-%d", time.gmtime(ms / 1000))


# ---------------------------------------------------------------- power ----
def smard():
    # 1223 lignite, 4067 wind onshore, 4068 solar, 4071 gas ... region DE, index lists chunk start times
    r = probe("power", "SMARD DE index (solar 4068)", "https://www.smard.de/app/chart_data/4068/DE/index_hour.json",
              extract=lambda r: (lambda ts: f"{len(ts)} weekly chunks, first {iso(ts[0])}, last {iso(ts[-1])}")(r.json()["timestamps"]))
    if r is not None and r.status_code == 200:
        ts = r.json()["timestamps"]
        probe("power", "SMARD DE latest chunk", f"https://www.smard.de/app/chart_data/4068/DE/4068_DE_hour_{ts[-1]}.json",
              extract=lambda r: f"{len(r.json()['series'])} hourly points, last={r.json()['series'][-1]}")


def rte():
    base = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/eco2mix-national-cons-def/records"
    probe("power", "RTE eCO2mix consolidated (latest)", f"{base}?limit=1&order_by=date_heure%20desc",
          extract=lambda r: f"total={r.json().get('total_count')} latest={r.json()['results'][0].get('date_heure')}")
    probe("power", "RTE eCO2mix consolidated (earliest)", f"{base}?limit=1&order_by=date_heure%20asc",
          extract=lambda r: f"earliest={r.json()['results'][0].get('date_heure')} keys={list(r.json()['results'][0])[:14]}")
    probe("power", "RTE eCO2mix real-time", "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/eco2mix-national-tr/records?limit=1&order_by=date_heure%20desc",
          extract=lambda r: f"latest={r.json()['results'][0].get('date_heure')}")


def ree():
    base = "https://apidatos.ree.es/en/datos/generacion/estructura-generacion"
    probe("power", "REE REData generation (last week, daily)",
          f"{base}?start_date={WEEK_AGO}T00:00&end_date={YDAY}T23:59&time_trunc=day",
          extract=lambda r: f"series={[s['type'] for s in r.json()['included']][:8]}")
    probe("power", "REE REData generation (2015, monthly)", f"{base}?start_date=2015-01-01T00:00&end_date=2015-12-31T23:59&time_trunc=month",
          extract=lambda r: f"2015 ok, {len(r.json()['included'])} series")


def neso():
    probe("power", "NESO CKAN historic generation mix",
          "https://api.nationalgrideso.com/api/3/action/package_show?id=historic-generation-mix",
          extract=lambda r: f"resources={[(x['name'], x['id']) for x in r.json()['result']['resources']][:6]}")
    probe("power", "Elexon BMRS FUELINST (last day)",
          f"https://data.elexon.co.uk/bmrs/api/v1/datasets/FUELINST?publishDateTimeFrom={YDAY}T00:00Z&publishDateTimeTo={YDAY}T01:00Z&format=json",
          extract=lambda r: f"rows={len(r.json()['data'])} first={r.json()['data'][0]}")
    probe("power", "Elexon BMRS FUELHH (2016)",
          "https://data.elexon.co.uk/bmrs/api/v1/datasets/FUELHH?publishDateTimeFrom=2016-01-01T00:00Z&publishDateTimeTo=2016-01-01T01:00Z&format=json",
          extract=lambda r: f"rows={len(r.json()['data'])} (2016 history ok)")
    probe("power", "Carbon Intensity API GB mix", f"https://api.carbonintensity.org.uk/generation/{YDAY}T00:00Z/{YDAY}T03:00Z",
          extract=lambda r: f"{len(r.json()['data'])} half-hours, mix={r.json()['data'][0]['generationmix'][:3]}")


def energinet():
    base = "https://api.energidataservice.dk/dataset"
    probe("power", "Energinet ElectricityProdex5MinRealtime", f"{base}/ElectricityProdex5MinRealtime?limit=1",
          extract=lambda r: f"total={r.json().get('total')} rec={r.json()['records'][0]}")
    probe("power", "Energinet ProductionConsumptionSettlement (hourly)", f"{base}/ProductionConsumptionSettlement?limit=1&sort=HourUTC%20asc",
          extract=lambda r: f"earliest={r.json()['records'][0].get('HourUTC')} keys={list(r.json()['records'][0])[:12]}")
    probe("gas", "Energinet Gasflow", f"{base}/Gasflow?limit=1",
          extract=lambda r: f"total={r.json().get('total')} rec={r.json()['records'][0]}")
    probe("gas", "Energinet GasConsumption/other gas datasets", f"{base}?limit=1",
          note="catalogue root; look for gas dataset names")


def pse():
    probe("power", "PSE Poland his-wlk-cal (generation by source)", "https://api.raporty.pse.pl/api/his-wlk-cal?$top=1&$orderby=udtczas%20desc",
          extract=lambda r: f"rec={r.json()['value'][0]}")
    probe("power", "PSE Poland his-wlk-cal (2018)", "https://api.raporty.pse.pl/api/his-wlk-cal?$filter=doba%20eq%20'2018-06-01'&$top=1",
          extract=lambda r: f"2018 rec={r.json()['value'][0]}")


def others_power():
    probe("power", "Statnett (Norway) detailed overview", "https://driftsdata.statnett.no/restapi/ProductionConsumption/GetLatestDetailedOverview",
          extract=lambda r: f"keys={list(r.json())[:10]}")
    probe("power", "Fingrid open data (no key)", "https://data.fingrid.fi/api/datasets/75/data?pageSize=1",
          note="expect 401 = key required (free)")
    probe("power", "Terna transparency (page)", "https://www.terna.it/en/electric-system/transparency-report/actual-generation")
    probe("power", "Terna API portal", "https://developer.terna.it/")
    probe("power", "TenneT NL open data", "https://www.tennet.org/english/operational_management/export_data.aspx")
    probe("power", "EirGrid SmartGrid (already pulled) ping", "https://www.smartgriddashboard.com/DashboardService.svc/data?area=generationactual&region=ALL&datefrom=01-Oct-2026&dateto=02-Oct-2026",
          extract=lambda r: f"len={len(r.text)}")


def aggregators_power():
    probe("power", "Energy-Charts public_power DE", f"https://api.energy-charts.info/public_power?country=de&start={WEEK_AGO}&end={YDAY}",
          extract=lambda r: f"types={[x['name']['en'] if isinstance(x['name'], dict) else x['name'] for x in r.json()['production_types']][:8]}")
    probe("power", "Energy-Charts public_power FR", f"https://api.energy-charts.info/public_power?country=fr&start={WEEK_AGO}&end={YDAY}",
          extract=lambda r: f"n_series={len(r.json()['production_types'])}")
    probe("power", "Energy-Charts public_power IT", f"https://api.energy-charts.info/public_power?country=it&start={WEEK_AGO}&end={YDAY}",
          extract=lambda r: f"n_series={len(r.json()['production_types'])}")
    probe("power", "ENTSO-E API (no key)", "https://web-api.tp.entsoe.eu/api?documentType=A75&processType=A16&in_Domain=10Y1001A1001A83F&periodStart=202609010000&periodEnd=202609020000",
          note="expect 401 = reachable, key required")


# ------------------------------------------------------------------ gas ----
def entsog():
    base = "https://transparency.entsog.eu/api/v1"
    probe("gas", "ENTSOG Physical Flow (recent)",
          f"{base}/operationalData?indicator=Physical%20Flow&periodType=day&from={WEEK_AGO}&to={YDAY}&timezone=CET&limit=3",
          extract=lambda r: f"rows={len(r.json()['operationalData'])} first={ {k: r.json()['operationalData'][0].get(k) for k in ('pointLabel', 'periodFrom', 'value', 'unit')} }")
    probe("gas", "ENTSOG Physical Flow (2015)",
          f"{base}/operationalData?indicator=Physical%20Flow&periodType=day&from=2015-10-01&to=2015-10-03&timezone=CET&limit=3",
          extract=lambda r: f"2015 rows={len(r.json()['operationalData'])}")
    probe("gas", "ENTSOG aggregatedData (consumption proxy)", f"{base}/AggregatedData?indicator=Physical%20Flow&periodType=day&from={WEEK_AGO}&to={YDAY}&limit=3",
          extract=lambda r: f"keys={list(r.json())} first={r.json().get('AggregatedData', r.json().get('aggregatedData', [{}]))[0]}")
    probe("gas", "ENTSOG interconnection points", f"{base}/interconnections?limit=2",
          extract=lambda r: f"first={r.json()['interconnections'][0]}")


def gie():
    probe("gas", "GIE AGSI+ (no key)", "https://agsi.gie.eu/api?country=DE&from=2026-09-01&to=2026-09-02",
          note="expect 401/403 = key required (free)")
    probe("gas", "GIE ALSI (no key)", "https://alsi.gie.eu/api?country=ES&from=2026-09-01&to=2026-09-02",
          note="expect 401/403 = key required (free)")


def gas_tsos():
    probe("gas", "GRTgaz/Teréga via ODRE datasets (consommation)",
          'https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?where=search(title,"consommation")%20and%20search(title,"gaz")&limit=8',
          extract=lambda r: f"{[d['dataset_id'] for d in r.json()['results']]}")
    probe("gas", "ODRE gas flows/transport datasets",
          'https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?where=search(title,"grtgaz")&limit=8',
          extract=lambda r: f"{[d['dataset_id'] for d in r.json()['results']]}")
    probe("gas", "National Gas (UK) data portal", "https://data.nationalgas.com/")
    probe("gas", "Trading Hub Europe (DE) aggregated consumption",
          "https://www.tradinghub.eu/en-gb/Publications/Transparency/Aggregated-consumption-data")
    probe("gas", "GTS (NL) transparency", "https://www.gasunietransportservices.nl/en/shippers/transparency")
    probe("gas", "Gassco (NO) flows", "https://umm.gassco.no/")
    probe("gas", "Snam (IT) transparency", "https://www.snam.it/en/transport/transparency/")
    probe("gas", "Enagas (ES) data", "https://www.enagas.es/en/technical-management-system/energy-data/")
    probe("gas", "GAZ-SYSTEM (PL)", "https://www.gaz-system.pl/en/customers/transparency-platform.html")


def main():
    for fn in (smard, rte, ree, neso, energinet, pse, others_power, aggregators_power, entsog, gie, gas_tsos):
        try:
            fn()
        except Exception as e:
            print(f"!! {fn.__name__} failed: {type(e).__name__}: {e}", file=sys.stderr)
    with open("eu_tso_probe_report.json", "w") as f:
        json.dump(RESULTS, f, indent=1, default=str)
    ok = sum(1 for r in RESULTS if r["status"] == 200)
    print(f"\nSUMMARY: {ok}/{len(RESULTS)} returned 200")
    for r in RESULTS:
        print(f"  {r['status']!s:>4}  {r['group']:5} {r['name']}")


if __name__ == "__main__":
    main()
