"""
Probe 2 (CEE capacity group): link crawl of the TSO / regulator landing pages, Energy-Charts error bodies,
CKAN (opendata.swiss) BFE dataset list, PSE API root, and downloads of the Austrian E-Control / APG workbooks.
Writes everything to discovery_archive/europe/capacity/results/probe2.txt (the workflow commits it).
"""
import io, json, os, re, sys, time
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe2.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=H, timeout=(10, 60), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
KW = re.compile(r"instal|capacit|kapacit|výkon|vykon|putere|teljes|beép|beep|leistung|moc|mesi[cč]n|monthly|bestand|engpass|erzeug|production|produkc|generation|xls|csv|ogd|capacit|мощност|мощ", re.I)
def crawl(name, url, maxn=60):
    r = get(url); 
    if r is None: return
    p(f"== {name} {url} -> {r.status_code} {len(r.content)}")
    links = re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I)
    seen = set(); n = 0
    for h, t in links:
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()[:80]
        if KW.search(h + " " + t) and h not in seen:
            seen.add(h); p("   ", t, "|", h[:160]); n += 1
            if n >= maxn: break
p("=== Energy-Charts")
for u in ["https://api.energy-charts.info/installed_power?country=pl&time_step=monthly&installation_decommission=false",
          "https://api.energy-charts.info/installed_power?country=pl&time_step=monthly",
          "https://api.energy-charts.info/installed_power?country=de&time_step=monthly",
          "https://api.energy-charts.info/installed_power?country=de&time_step=yearly",
          "https://api.energy-charts.info/installed_power?country=pl",
          "https://api.energy-charts.info/openapi.json"]:
    r = get(u)
    if r is not None:
        p(r.status_code, u, r.headers.get("content-type"), re.sub(r"\s+", " ", r.text)[:300])
        if r.status_code == 200 and "json" in r.headers.get("content-type", "") and "openapi" not in u:
            j = r.json(); p("   keys", list(j)[:8], "n", len(j.get("time", [])), j.get("time", [])[:2], j.get("time", [])[-2:], [s["name"] for s in j.get("production_types", [])])
        if "openapi" in u and r.status_code == 200:
            j = r.json(); p("   paths", [k for k in j["paths"] if "inst" in k or "power" in k])
            p("   installed_power params", json.dumps(j["paths"].get("/installed_power", {}))[:900])
time.sleep(2)
p("=== PSE")
for u in ["https://api.raporty.pse.pl/api", "https://api.raporty.pse.pl/api/", "https://api.raporty.pse.pl/api/$metadata", "https://api.raporty.pse.pl/", "https://www.pse.pl/web/pse-eng/", "https://www.pse.pl/dane-systemowe", "https://www.pse.pl/en/data"]:
    r = get(u)
    if r is not None: p(r.status_code, u, r.headers.get("content-type"), re.sub(r"\s+", " ", r.text)[:500])
crawl("pse.pl en", "https://www.pse.pl/web/pse-eng/")
crawl("pse.pl dane", "https://www.pse.pl/dane-systemowe")
crawl("URE potencjal", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze", 80)
crawl("ARE", "https://www.are.waw.pl/")
p("=== CZ")
crawl("ERU home", "https://www.eru.gov.cz/")
crawl("ERU en", "https://www.eru.gov.cz/en")
crawl("CEPS all-data", "https://www.ceps.cz/en/all-data", 80)
crawl("OTE", "https://www.ote-cr.cz/en/statistics/yearly-market-report")
p("=== SK")
crawl("SEPS en", "https://www.sepsas.sk/en/", 50)
crawl("SEPS sk", "https://www.sepsas.sk/", 50)
crawl("URSO", "https://www.urso.gov.sk/", 40)
p("=== HU")
for u in ["https://www.mavir.hu/", "https://www.mavir.hu/en", "https://www.mavir.hu/web/mavir/adatok", "https://www.mavir.hu/web/mavir-en/home"]:
    crawl("MAVIR " + u, u, 60)
crawl("MEKH hu", "https://mekh.hu/")
p("=== RO")
crawl("Transelectrica", "https://www.transelectrica.ro/en/web/tel/productie", 60)
crawl("Transelectrica home", "https://www.transelectrica.ro/", 60)
crawl("ANRE", "https://www.anre.ro/", 60)
p("=== BG")
crawl("ESO", "https://www.eso.bg/", 60)
for u in ["https://www.dker.bg/", "https://www.dker.bg/bg/", "https://www.dker.bg/en/", "https://www.dker.bg/uploads/_CKFinder_/"]:
    r = get(u)
    if r is not None: p(r.status_code, u, len(r.content))
p("=== AT")
r = get("https://www.e-control.at/statistik/e-statistik/archiv/bestandsstatistik")
if r is not None:
    for m in re.finditer(r'href=["\']([^"\']+\.xlsx?)["\']', r.text): p("   ", m.group(1))
r = get("https://www.e-control.at/statistik/e-statistik/archiv/betriebsstatistik")
if r is not None:
    for m in re.finditer(r'href=["\']([^"\']+\.xlsx?)["\']', r.text): p("   BS", m.group(1))
crawl("econtrol strom", "https://www.e-control.at/statistik/strom", 60)
crawl("APG installed de", "https://markttransparenz.apg.at/de/markt/Markttransparenz/Erzeugung/Installierte-Leistung", 40)
def dl(url, name):
    r = get(url)
    if r is None or r.status_code != 200: p("dl fail", url, r.status_code if r is not None else None); return None
    fn = os.path.join("/tmp", name); open(fn, "wb").write(r.content); p("dl ok", url, len(r.content), r.headers.get("last-modified")); return fn
import pandas as pd
for url, nm in [("https://www.e-control.at/documents/1785851/1811756/BeStGes-2025_KW1EPLKls.xlsx", "a.xlsx"),
                ("https://www.e-control.at/documents/1785851/1811756/BeStGes-2025_KW2EPLTyp.xlsx", "b.xlsx"),
                ("https://www.e-control.at/documents/1785851/1811756/BeStGes-JR_KWEPL.xlsx", "c.xlsx"),
                ("https://pb1-medien.apg.at/im/dl/apg-680580447/Installierte Erzeugungskapazität pro Produktionstyp.xlsx", "d.xlsx")]:
    fn = dl(url, nm)
    if fn:
        try:
            x = pd.ExcelFile(fn); p("   sheets", x.sheet_names)
            for s in x.sheet_names[:4]:
                d = x.parse(s, header=None); p("   --", s, d.shape); p(d.head(14).to_string()[:2500])
        except Exception as e: p("   xl err", e)
p("=== CH CKAN")
r = get("https://ckan.opendata.swiss/api/3/action/package_search?q=organization:bundesamt-fur-energie-bfe&rows=200")
if r is not None and r.status_code == 200:
    for pk in r.json()["result"]["results"]:
        t = pk["title"].get("de") or pk["title"].get("en") or str(pk["title"])
        if re.search(r"anlage|leistung|kapaz|produktion|strom|elektr|photovolt|wind|wasserkraft|pronovo|speicher", t, re.I):
            p("  ", t[:90], "|", [rs.get("download_url") or rs.get("url") for rs in pk["resources"]][:3])
