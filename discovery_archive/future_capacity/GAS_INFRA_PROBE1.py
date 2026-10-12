"""Probe 1 (Europe future gas infrastructure): which official sources answer from Actions, and what links they expose."""
import os, re, sys, json, time
import requests
OUT = "discovery_archive/results/future_capacity"
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
KEY = os.environ.get("GIE_API_KEY", "")
URLS = [
 "https://www.gie.eu/transparency-platform/gle-map/", "https://www.gie.eu/publications/maps/", "https://www.gie.eu/publications/maps/gie-lng-map/",
 "https://www.gie.eu/publications/maps/gie-storage-map/", "https://www.gie.eu/transparency-platform/GIE_LNG_Map_Dataset.xlsx",
 "https://alsi.gie.eu/", "https://agsi.gie.eu/",
 "https://tyndp.entsog.eu/", "https://tyndp.entsog.eu/api/projects", "https://tyndp.entsog.eu/projects/",
 "https://www.entsog.eu/tyndp", "https://www.entsog.eu/tyndp-2024-projects",
 "https://transparency.entsog.eu/api/v1/operatorpointdirections?limit=3",
 "https://factpages.sodir.no/en/field/Overview/", "https://factpages.sodir.no/downloads/csv/field.csv",
 "https://factpages.sodir.no/public?/Factpages/external/tableview/field&rs:Format=CSV",
 "https://www.sodir.no/en/facts/data-and-services/", "https://www.sodir.no/en/whats-new/publications/",
 "https://www.nstauthority.co.uk/data-centre/", "https://www.nstauthority.co.uk/data-centre/data-downloads-and-publications/",
 "https://energy.ec.europa.eu/topics/infrastructure/projects-common-interest-and-projects-mutual-interest_en",
 "https://ec.europa.eu/energy/infrastructure/transparency_platform/map-viewer/main.html",
 "https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:32024R0001",
 "https://energy.ec.europa.eu/system/files/2024-04/PCI_PMI_list_2023.xlsx",
]
def probe(u, hdr=None, save=None):
    try:
        r = requests.get(u, headers=dict(UA, **(hdr or {})), timeout=(10, 60), allow_redirects=True)
    except Exception as e:
        print("FAIL", u, type(e).__name__, str(e)[:100]); return None
    ct = r.headers.get("content-type", "")
    print(r.status_code, len(r.content), ct[:40], u, "->", r.url if r.url != u else "")
    if r.status_code == 200 and "html" in ct:
        links = sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf|json|zip)[^"]*)"', r.text, re.I)))
        for l in links[:40]: print("    LINK", l)
        apis = sorted(set(re.findall(r'["\'](/?api[^"\']{3,80})["\']', r.text)))[:15]
        for a in apis: print("    API?", a)
    if save and r.status_code == 200:
        open(os.path.join(OUT, save), "wb").write(r.content)
    return r
for u in URLS: probe(u)
if KEY:
    H = {"x-key": KEY}
    r = probe("https://alsi.gie.eu/api/about", H, "alsi_about.json")
    if r is not None and r.status_code == 200: print(r.text[:1500])
    r = probe("https://alsi.gie.eu/api?country=DE&from=2026-10-01&to=2026-10-05&size=5", H, "alsi_de.json")
    if r is not None and r.status_code == 200: print(r.text[:1500])
    r = probe("https://agsi.gie.eu/api/about", H, "agsi_about.json")
    if r is not None and r.status_code == 200: print(r.text[:1500])
    r = probe("https://agsi.gie.eu/api?country=DE&from=2026-10-01&to=2026-10-02&size=5", H, "agsi_de.json")
    if r is not None and r.status_code == 200: print(r.text[:2500])
