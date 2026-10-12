"""Probe 3: TYNDP 2024 site media, PCI list via CELLAR, Sodir factmaps layers, NSTA hub datasets, operator pages (keyword context)."""
import os, re, json
import requests
OUT = "discovery_archive/results/future_capacity"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
def get(u, hdr=None, **kw):
    try:
        r = requests.get(u, headers=dict(UA, **(hdr or {})), timeout=(10, 90), **kw)
        print(r.status_code, len(r.content), r.headers.get("content-type", "")[:30], u[:140]); return r
    except Exception as e:
        print("FAIL", u[:120], type(e).__name__); return None
def text(r):
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
def hits(r, pats, n=3, w=110):
    t = text(r)
    for p in pats:
        for m in list(re.finditer(p, t, re.I))[:n]:
            print("   >", p, "::", t[max(0, m.start() - w): m.end() + w])
# TYNDP 2024 site media
for q in ["annex", "project", "xlsx", "table"]:
    r = get(f"https://tyndp2024.entsog.eu/wp-json/wp/v2/media?search={q}&per_page=50")
    if r is not None and r.status_code == 200:
        try:
            for m in r.json(): print("   M", m.get("source_url"), m.get("title", {}).get("rendered", "")[:60])
        except Exception as e: print("  json err", e)
r = get("https://tyndp2024.entsog.eu/infrastructure-report-introduction/promoters-submissions-for-natural-gas-to-tyndp-2024/")
if r is not None:
    for l in sorted(set(re.findall(r'(?:href|src)="([^"]+\.(?:xlsx?|csv|pdf|json|zip))"', r.text, re.I)))[:30]: print("   L", l)
    print("  ", text(r)[:600])
# PCI list
for u in ["https://publications.europa.eu/resource/celex/32024R1041", "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32024R1041",
          "https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:32024R1041", "https://energy.ec.europa.eu/topics/infrastructure/projects-common-interest-and-projects-mutual-interest/pci-and-pmi-selection-process_en",
          "https://energy.ec.europa.eu/topics/infrastructure/projects-common-interest-and-projects-mutual-interest/union-lists-projects-common-interest-and-projects-mutual-interest_en"]:
    r = get(u, {"Accept": "application/xhtml+xml, text/html, application/pdf"}, allow_redirects=True)
    if r is not None and r.status_code == 200:
        print("   ", text(r)[:200])
        for l in sorted(set(re.findall(r'href="([^"]+)"', r.text)))[:400]:
            if re.search(r"1041|union.list|pci.*(list|annex)|xlsx|\.pdf", l, re.I): print("   L", l[:170])
# Sodir factmaps layers
r = get("https://factmaps.sodir.no/api/rest/services/Factmaps/FactMapsWGS84/MapServer?f=json")
lay = {}
if r is not None and r.status_code == 200:
    for l in r.json().get("layers", []):
        lay[l["id"]] = l["name"]
    print("  layers:", {k: v for k, v in lay.items() if re.search(r"field|discover|pipeline|facilit|plan|PDO|surface", v, re.I)})
for k, v in lay.items():
    if re.search(r"^(Fields?|Discoveries|Discovery|Fields?, |Discoveries, )", v, re.I) or v.lower() in ("field", "fields", "discovery", "discoveries"):
        rr = get(f"https://factmaps.sodir.no/api/rest/services/Factmaps/FactMapsWGS84/MapServer/{k}/query?where=1%3D1&outFields=*&returnGeometry=false&resultRecordCount=3&f=json")
        if rr is not None and rr.status_code == 200: print("   Q", v, rr.text[:900])
# NSTA hub datasets
r = get("https://opendata-nstauthority.hub.arcgis.com/api/v3/datasets?q=fields&page%5Bsize%5D=100")
if r is not None and r.status_code == 200:
    for d in r.json().get("data", []):
        a = d.get("attributes", {})
        pass
# operator pages
DOCS = {
 "https://www.tap-ag.com/news/news-stories/tap-triggers-the-first-level-of-capacity-expansion": [r"1\.2 ?bcm", r"2027", r"1 January 2026"],
 "https://energy-terminal.de/en/terminals": [r"Stade", r"Brunsb", r"Wilhelmshaven"],
 "https://energy-terminal.de/en/construction-diary-stade": [r"billion cubic", r"2027", r"November"],
 "https://www.gaz-system.pl/en/transmission-system/investment-projects/fsru-terminal.html": [r"bcm|billion", r"202[6-9]"],
 "https://www.gasunie.nl/en/projects/eemsenergyterminal": [r"2036", r"2027", r"bcm"],
 "https://lng.hr/en/2026/01/01/5-years-of-operation-of-the-lng-terminal-croatia/": [r"6\.1", r"billion"],
 "https://www.gastrade.gr/the-project/": [r"5[.,]5", r"Thrace"],
 "https://gasgrid.fi/en/lng-terminal/inkoo-lng-terminal/": [r"GWh", r"lease|charter"],
 "https://totalenergies.com/news/press-releases/france-totalenergies-demobilizes-its-floating-lng-terminal-le-havre": [r"demobiliz", r"2026"],
 "https://www.snam.it/en/our-businesses/energy-morning/necp,-the-moees-sends-its-proposal-to-brussels.html": [r"Vado", r"2026"],
 "https://www.omvpetrom.com/en/news": [r"Neptun"],
}
for u, pats in DOCS.items():
    r = get(u)
    if r is not None and r.status_code == 200: hits(r, pats)
