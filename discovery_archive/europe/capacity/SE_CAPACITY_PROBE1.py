"""
Probe 1 (South-East Europe capacity group: TR, GR, HR, SI, RS, BA, ME, MK, AL, XK, MT). Output: results/se_probe1.txt
Fetches landing pages of each candidate source (requests, then curl_cffi if blocked), prints status, title and links whose
text/href looks like capacity data (kurulu, capacity, kapacit, instal, snaga, moc, xls, pdf, csv, json, api).
Findings are summarised in the final probe docstring of SE_CAPACITY_PROBE4.py.
"""
import os, re, sys, json
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en,tr;q=0.8,el;q=0.7,hr;q=0.6"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe1.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(10, 40), **kw)
        if r.status_code in (403, 503, 429, 412):
            raise RuntimeError(f"http {r.status_code}")
        return r
    except Exception as e:
        try:
            from curl_cffi import requests as cr
            r = cr.get(url, impersonate="chrome120", timeout=40, **kw)
            p("   (curl_cffi used after", type(e).__name__, str(e)[:50] + ")")
            return r
        except Exception as e2:
            p("ERR", url, type(e).__name__, str(e)[:80], "|", type(e2).__name__, str(e2)[:80]); return None
KW = re.compile(r"kurulu|g[uü]c|capacit|kapasit|install|snag|mo[cć]|isko|instal|ισχ|εγκατ|rae|xls|csv|json|api|zip|monthly|aylik|ayl[ıi]k|mjese|meseč|mesec|μηνια|statist|izvje|report|raport|godi", re.I)
def crawl(url, n=60, show=True):
    r = get(url)
    if r is None: return None
    t = re.search(r"<title[^>]*>(.*?)</title>", r.text[:50000], re.S | re.I)
    p(f"== {url} -> {r.status_code} {len(r.content)}B {r.headers.get('content-type','')[:40]} title={re.sub(chr(10),' ',(t.group(1) if t else '')).strip()[:80]}")
    if not show or r.status_code != 200: return r
    seen = set(); k = 0
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        h = urljoin(url, m.group(1)); tx = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:90]
        if h in seen or h.startswith("javascript"): continue
        if KW.search(h + " " + tx):
            seen.add(h); k += 1; p("   ", tx, "->", h[:170])
            if k >= n: break
    return r
SEEDS = [
 # Turkey
 "https://www.teias.gov.tr/kurulu-guc-raporlari", "https://www.teias.gov.tr/turkiye-elektrik-uretim-iletim-istatistikleri",
 "https://www.epdk.gov.tr/Detay/Icerik/3-0-0-98/elektrik-piyasasi-sektor-raporlari", "https://www.epdk.gov.tr/",
 "https://seffaflik.epias.com.tr/", "https://seffaflik.epias.com.tr/electricity-service/v1/dashboard/installed-capacity",
 "https://seffaflik.epias.com.tr/electricity-service/v1/dashboard/realtime-generation",
 "https://seffaflik.epias.com.tr/electricity-service/technical/tr/index.html",
 "https://enerji.gov.tr/bilgi-merkezi-enerji-elektrik", "https://enerji.gov.tr/",
 # Greece
 "https://www.admie.gr/en/market/market-statistics/detailed-monthly-reports", "https://www.admie.gr/en/market/market-statistics/monthly-energy-balance",
 "https://www.dapeep.gr/en/", "https://www.dapeep.gr/category/statistika/", "https://deddie.gr/en/",
 "https://www.rae.gr/", "https://geo.rae.gr/", "https://www.henex.gr/",
 # Croatia
 "https://www.hops.hr/", "https://www.hrote.hr/", "https://www.hera.hr/", "https://www.hrote.hr/statistika-oie",
 # Slovenia
 "https://www.eles.si/", "https://www.agen-rs.si/", "https://www.borzen.si/en-us/Home", "https://www.sodo.si/", "https://pxweb.stat.si/SiStatData/pxweb/en/Data/",
 # Serbia
 "https://ems.rs/", "https://www.mre.gov.rs/", "https://www.stat.gov.rs/",
 # Bosnia, Montenegro, N. Macedonia, Albania, Kosovo, Malta
 "https://www.nosbih.ba/", "https://www.derk.ba/", "https://www.cges.me/", "https://www.mepso.com.mk/", "https://www.ere.gov.mk/",
 "https://www.ost.al/", "https://www.ere.gov.al/", "https://www.kostt.com/", "https://www.ero-ks.org/",
 "https://www.rews.org.mt/", "https://nso.gov.mt/",
]
for u in SEEDS:
    crawl(u)
p("=== EPIAS guesses")
for ep in ["dashboard/installed-capacity", "dashboard/realtime-generation", "generation/data/installed-capacity", "generation/data/installed-capacity-group",
           "generation/data/license-based-installed-capacity", "dashboard/installed-power"]:
    u = "https://seffaflik.epias.com.tr/electricity-service/v1/" + ep
    for meth in ("GET", "POST"):
        try:
            r = requests.request(meth, u, headers={**H, "Content-Type": "application/json", "Accept": "application/json", "Referer": "https://seffaflik.epias.com.tr/"},
                                 data=json.dumps({"period": "2025-01-01T00:00:00+03:00"}) if meth == "POST" else None, timeout=30)
            p(meth, ep, r.status_code, r.text[:400].replace("\n", " "))
        except Exception as e:
            p(meth, ep, "ERR", type(e).__name__, str(e)[:80])
