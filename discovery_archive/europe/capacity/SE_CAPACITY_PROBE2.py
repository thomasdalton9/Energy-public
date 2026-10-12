"""
Probe 2 (SE Europe capacity). Output: results/se_probe2.txt. Lists file/data links on the pages probe 1 found promising:
TEIAS Kurulu Guc, EPDK monthly sector reports, EPIAS technical doc (installed-capacity services, dashboard paths),
HROTE, Borzen, RAAEY statistics, HOPS annual reports, DEDDIE islands, CGES/NOSBiH/MEPSO/ERE/ZRRE, SiStat PxWeb search.
"""
import os, re, sys, json, html
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en,tr;q=0.8,el;q=0.7,hr;q=0.6"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe2.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(10, 60), **kw)
        if r.status_code in (403, 503, 429, 412): raise RuntimeError(f"http {r.status_code}")
        return r
    except Exception as e:
        try:
            from curl_cffi import requests as cr
            return cr.get(url, impersonate="chrome120", timeout=60, **kw)
        except Exception as e2:
            p("ERR", url, str(e)[:60], "|", str(e2)[:60]); return None
def links(url, pat, n=45, text_pat=None):
    r = get(url)
    if r is None: return None
    p(f"== {url} -> {r.status_code} {len(r.content)}B")
    if r.status_code != 200: return r
    txt = r.content.decode(r.encoding or "utf-8", "replace") if not r.encoding else r.text
    seen, k = set(), 0
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', txt, re.S | re.I):
        h = urljoin(url, html.unescape(m.group(1))); tx = html.unescape(re.sub(r"<[^>]+>|\s+", " ", m.group(2))).strip()[:80]
        if h in seen: continue
        if re.search(pat, h, re.I) or (text_pat and re.search(text_pat, tx, re.I)):
            seen.add(h); k += 1; p("   ", tx, "->", h[:200])
            if k >= n: p("    ... more"); break
    if k == 0: p("    (no matching links)")
    return r
FILE = r"\.(xlsx?|csv|pdf|zip|json)(\?|$)"
# Turkey
links("https://www.teias.gov.tr/kurulu-guc-raporlari", FILE + r"|kurulu", 60, r"kurulu|20\d\d")
links("https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar", FILE + r"|Dosya|Download|sektor-rapor", 40, r"rapor|20\d\d")
links("https://www.epdk.gov.tr/Detay/Icerik/3-0-66-3/elektrikuretim-kapasite-projeksiyonlari", FILE, 20)
r = get("https://seffaflik.epias.com.tr/electricity-service/technical/tr/index.html")
if r is not None:
    t = r.text
    for key in ("_flow-rate-and-installed-power\"", "_new-installed-capacity\"", "dashboard"):
        for m in list(re.finditer(re.escape(key), t))[:3]:
            p("--- EPIAS doc", key, ":", re.sub(r"<[^>]+>|\s+", " ", t[m.start():m.start() + 1800])[:900])
    p("EPIAS 'dashboard' paths:", sorted(set(re.findall(r"/v1/dashboard/[a-z\-/]+", t)))[:40])
    p("EPIAS installed paths:", sorted(set(re.findall(r"/v1/[a-z\-/]*install[a-z\-/]*", t)))[:40])
    p("EPIAS auth hint:", re.sub(r"<[^>]+>|\s+", " ", t[t.find("TGT") - 200:t.find("TGT") + 400])[:500] if "TGT" in t else "no TGT")
# Croatia / Slovenia / Greece
links("https://www.hrote.hr/obnovljivi-izvori-energije-i-viskoucinkovita-kogeneracija", FILE + r"|izvjest|statist|mjesec", 40, r"izvje|statist|mjese")
links("https://www.hrote.hr/izvjestaji", FILE + r"|izvjest", 40, r"izvje|mjese|20\d\d")
links("https://www.hops.hr/godisnji-izvjestaji", FILE, 15)
links("https://borzen.si/en-us/electricity-production-support/data-and-reports/data-and-reports", FILE + r"|report|data", 40, r"report|monthly|statist")
links("https://www.eles.si/", r"poroc|report|statist|kapacit|moc", 20, r"poro|report|statist")
links("https://www.sodo.si/", r"statist|podatki|samooskrb|porocil", 25, r"statist|podat|samooskrb|poro")
links("https://www.raaey.gr/energeia/statistika/", FILE + r"|statist|ape|energeia/", 50, r"στατ|ΑΠΕ|εγκατ|ισχ")
links("https://deddie.gr/en/services/non-interconnected-islands/", FILE + r"|statist|ape|res", 40, r"statist|capacity|RES|monthly")
links("https://www.admie.gr/en/market/market-statistics", FILE + r"|monthly|statist|capacity", 40)
links("https://www.admie.gr/", r"statist|monthly|report|capacity", 20)
# Balkans
for u in ("https://www.cges.me/", "https://www.nosbih.ba/", "https://www.mepso.com.mk/", "https://www.ere.gov.al/", "https://www.ero-ks.org/", "https://www.derk.ba/", "https://ems.rs/"):
    links(u, FILE + r"|izvje|izvjes|kapacit|instal|godisn|annual|report|raport|izve", 25, r"capacity|kapacit|install|annual|godi|izvje|raport|report|годиш|извеш|снаг")
# Malta
links("https://www.rews.org.mt/", FILE + r"|statist|renewable|capacity", 20, r"statist|capacity|PV")
links("https://nso.gov.mt/", r"electric|energy|statist", 15, r"electric|energy")
# SiStat search
for q in ("electric", "capacity", "power plant", "energy"):
    r = get("https://pxweb.stat.si/SiStatData/api/v1/en/Data/")
    break
if r is not None and r.status_code == 200:
    try:
        j = r.json(); p("SiStat top:", [(x.get("id"), x.get("text")) for x in j][:60])
    except Exception as e: p("SiStat json fail", e, r.text[:200])
r = get("https://pxweb.stat.si/SiStatData/api/v1/en/Data/?query=*electric*&filter=*")
if r is not None: p("SiStat search:", r.status_code, r.text[:1500])
# EPIAS home page JS bundles: every service path the public site calls
r = get("https://seffaflik.epias.com.tr/")
if r is not None:
    srcs = re.findall(r'src=["\']([^"\']+\.js[^"\']*)', r.text); p("EPIAS js:", srcs)
    paths = set()
    for s in srcs[:12]:
        rr = get(urljoin("https://seffaflik.epias.com.tr/", s))
        if rr is not None and rr.status_code == 200:
            paths |= set(re.findall(r"electricity-service/v1/[A-Za-z0-9_\-/]+", rr.text))
            paths |= set(re.findall(r"[\"'](/?(?:generation|dashboard|market|consumption)/[A-Za-z0-9_\-/]+)[\"']", rr.text))
            for m in re.finditer(r"install", rr.text, re.I):
                p("   js ctx:", rr.text[max(0, m.start() - 80):m.start() + 120].replace("\n", " ")); break
    p("EPIAS paths:", sorted(paths)[:200])
