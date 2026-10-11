"""
Discovery: reservoir / lake / river LEVEL or STORAGE data for West African hydro
(Akosombo, Bui, Kainji, Jebba, Shiroro, Zungeru, Manantali, Felou, Gouina, Kompienga,
Soubre, Kossou, Kaleta, Souapiti, Mount Coffee, Lagdo, Song Loulou), satellite
reservoir products (G-REALM, Hydroweb, GRLM, Copernicus lakes, DAHITI, ESA CCI) and
river gauges (Niger, Volta, GRDC, Niger-HYCOS). Probes URLs, prints status/type/size,
a snippet, level-related keywords and data-file links. Nothing committed; no Ember.
Writes west_africa_hydro_levels_output.txt (uploaded as artifact).
"""
import re
import signal
import sys
import requests


class Hard(Exception):
    pass


def _alarm(*a):
    raise Hard('hard 45s deadline')


signal.signal(signal.SIGALRM, _alarm)

OUT = open("west_africa_hydro_levels_output.txt", "w", encoding="utf-8")


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    OUT.write(s + "\n")
    OUT.flush()


H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
KW = re.compile(r"(lake level|water level|reservoir level|niveau|cote|elevation|storage|"
                r"akosombo|kainji|jebba|shiroro|manantali|bui dam|volta lake|hydrolog|"
                r"stage|discharge|debit)", re.I)
LINK = re.compile(r'href=["\']([^"\']+\.(?:csv|xlsx?|txt|json|zip|nc|pdf))["\']', re.I)

URLS = [
    # Operators
    ("VRA home", "https://www.vra.com"),
    ("VRA lake level", "https://www.vra.com/our_mandate/hydro_generation.php"),
    ("VRA search", "https://www.vra.com/?s=lake+level"),
    ("Bui Power Authority", "https://www.buipower.com"),
    ("Mainstream Energy", "https://www.mainstreamenergy.com.ng"),
    ("NISO", "https://www.niso.gov.ng"),
    ("TCN", "https://www.tcn.org.ng"),
    ("NIHSA", "https://nihsa.gov.ng"),
    ("Niger Basin Authority", "https://www.abn.ne"),
    ("Niger-HYCOS", "https://www.abn.ne/index.php/fr/hycos"),
    ("OMVS", "https://www.omvs.org"),
    ("SOGEM", "https://www.sogem-omvs.org"),
    ("Volta Basin Authority", "https://www.abv-volta.org"),
    ("Volta HYCOS", "https://www.abv-volta.org/en/"),
    ("CIE Cote d'Ivoire", "https://www.cie.ci"),
    ("ANARE CI", "https://www.anare.ci"),
    ("EDG Guinea", "https://www.edg.com.gn"),
    ("LEC Liberia", "https://www.lec.com.lr"),
    ("ENEO Cameroon", "https://www.eneocameroon.cm"),
    ("Cameroon Ministry water", "https://www.minee.cm"),
    ("WAPP", "https://www.ecowapp.org"),
    ("WAPP info platform", "https://www.ecowapp.org/en/content/wapp-information-and-coordination-center-icc"),
    # Satellite / global products
    ("G-REALM index", "https://ipad.fas.usda.gov/cropexplorer/global_reservoir/"),
    ("G-REALM lake list", "https://ipad.fas.usda.gov/cropexplorer/global_reservoir/gr_regional_chart.aspx?regionid=afr"),
    ("G-REALM Volta txt", "https://ipad.fas.usda.gov/lakes/images/lake0020.10d.2.txt"),
    ("Hydroweb.next STAC", "https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac"),
    ("Hydroweb.next collections", "https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac/collections"),
    ("Hydroweb old", "https://hydroweb.theia-land.fr"),
    ("DAHITI", "https://dahiti.dgfi.tum.de/en/"),
    ("DAHITI Volta search", "https://dahiti.dgfi.tum.de/en/map/?search=Volta"),
    ("GRLM zenodo search", "https://zenodo.org/api/records?q=Global+Reservoir+and+Lake+Monitor&size=5"),
    ("GRLM figshare search", "https://api.figshare.com/v2/articles/search?search_for=GRLM%20reservoir"),
    ("Copernicus lakes WL", "https://land.copernicus.eu/en/products/water/lake-water-level-daily-1km-global"),
    ("Copernicus lakes WL (CDS)", "https://cds.climate.copernicus.eu/datasets/satellite-lake-water-level"),
    ("ESA CCI lakes (CEDA)", "https://catalogue.ceda.ac.uk/uuid/ab3f7eef8d594a81a2b6a3a7a0d4c4b2/"),
    ("HydroLAKES/GRanD", "https://www.globaldamwatch.org/grand"),
    ("ResOpsUS-like global ResOps", "https://zenodo.org/api/records?q=reservoir+operations+Africa+storage&size=5"),
    # River gauges
    ("GRDC portal", "https://portal.grdc.bafg.de"),
    ("GRDC station catalogue", "https://grdc.bafg.de/data/data_portal/"),
    ("GRDC Niamey search", "https://portal.grdc.bafg.de/KiWebPortal/rest/grdc/stations?search=Niamey"),
    ("Niger-HYCOS (WMO hydrohub)", "https://hydrohub.wmo.int"),
    ("HydroSOS", "https://hydrosos.org"),
    ("GloFAS point API (Open-Meteo flood)", "https://flood-api.open-meteo.com/v1/flood?latitude=13.51&longitude=2.11&daily=river_discharge&past_days=30&forecast_days=1"),
    ("GloFAS Akosombo (Open-Meteo flood)", "https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&past_days=30&forecast_days=1"),
    ("GloFAS Kainji (Open-Meteo flood)", "https://flood-api.open-meteo.com/v1/flood?latitude=10.4&longitude=4.6&daily=river_discharge&past_days=30&forecast_days=1"),
    ("GloFAS Manantali (Open-Meteo flood)", "https://flood-api.open-meteo.com/v1/flood?latitude=13.2&longitude=-10.4&daily=river_discharge&past_days=30&forecast_days=1"),
    ("AGRHYMET", "https://www.agrhymet.ne"),
    ("SIEREM / HydroSciences", "https://www.hydrosciences.fr"),
    ("Niger basin DSS (ABN SAP)", "https://www.abn.ne/index.php/fr/"),
    ("Wikimedia Akosombo data", "https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=Lake+Volta+level+Akosombo+feet&format=json"),
]


def probe(name, url):
    log("=" * 100)
    log(f"## {name}\n{url}")
    try:
        signal.alarm(45)
        r = requests.get(url, headers=H, timeout=(10, 30), allow_redirects=True)
        signal.alarm(0)
    except BaseException as e:
        signal.alarm(0)
        log("  ERROR", type(e).__name__, str(e)[:200])
        return
    ct = r.headers.get("content-type", "")
    log(f"  status {r.status_code} final {r.url} type {ct} bytes {len(r.content)}")
    if r.status_code >= 400:
        return
    txt = r.text if ("text" in ct or "json" in ct or "xml" in ct) else ""
    if txt:
        body = re.sub(r"\s+", " ", re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", txt, flags=re.S | re.I))
        plain = re.sub(r"<[^>]+>", " ", body)
        plain = re.sub(r"\s+", " ", plain)
        log("  snippet:", plain[:350])
        hits = sorted({m.group(0).lower() for m in KW.finditer(plain)})
        log("  keywords:", hits[:20])
        links = LINK.findall(txt)
        if links:
            log("  data links:", links[:15])
        for m in list(KW.finditer(plain))[:4]:
            s = max(0, m.start() - 80)
            log("   ctx:", plain[s:m.end() + 120])
    else:
        log("  (non-text)", r.content[:120])


for n, u in URLS:
    probe(n, u)

# Follow-ups: Open-Meteo JSON shape
log("=" * 100)
log("Done")
OUT.close()
