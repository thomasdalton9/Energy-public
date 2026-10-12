"""
Probe 3 (SE Europe capacity). Output: results/se_probe3.txt
 - EPIAS: documented params of powerplant-list / licensed-powerplant-investment-list / new+old-installed-capacity / aic;
   unauthenticated POST tries on the public-site service paths
 - EPDK monthly electricity sector report entries (titles) + first download inspected
 - TEIAS Kurulu Guc page structure (script data), HROTE monthly report PDF sample + year pages,
   Borzen report page, ADMIE file-download-api, SODO self-supply page, DEDDIE monthly/RES statistics pages
"""
import os, re, sys, json, html, io
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en,tr;q=0.8,el;q=0.7,hr;q=0.6"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe3.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(10, 90), **kw)
        if r.status_code in (403, 503, 429, 412): raise RuntimeError(f"http {r.status_code}")
        return r
    except Exception as e:
        try:
            from curl_cffi import requests as cr
            return cr.get(url, impersonate="chrome120", timeout=90, **kw)
        except Exception as e2:
            p("ERR", url, str(e)[:60], "|", str(e2)[:60]); return None
def clean(s, n=600): return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s)))[:n]
def pdf_text(content, pages=3, n=2500):
    import fitz
    d = fitz.open(stream=content, filetype="pdf"); p("   pdf pages", len(d))
    return "\n".join(d[i].get_text() for i in range(min(pages, len(d))))[:n]
BASE = "https://seffaflik.epias.com.tr/electricity-service/"
p("=== EPIAS doc")
r = get(BASE + "technical/tr/index.html")
if r is not None:
    t = r.content.decode("utf-8", "replace")
    for key in ("/v1/generation/data/powerplant-list\"", "/v1/generation/data/powerplant-list-for-date-range", "/v1/generation/data/licensed-powerplant-investment-list", "/v1/renewables/data/new-installed-capacity", "/v1/renewables/data/old-installed-capacity", "/v1/generation/data/aic"):
        i = t.find(key.strip('"'))
        if i < 0: p("not found", key); continue
        p("---", key, ":", clean(t[i:i + 9000], 1400))
    j = t.find("REQUEST_WITHOUT_TGT")
    p("--- TGT note:", clean(t[max(0, j - 3000):j - 500], 700))
p("=== EPIAS posts")
HP = {**H, "Content-Type": "application/json", "Accept": "application/json", "Referer": "https://seffaflik.epias.com.tr/"}
def post(ep, body):
    try:
        rr = requests.post(BASE + "v1/" + ep, headers=HP, data=json.dumps(body), timeout=60)
        p("POST", ep, json.dumps(body)[:100], "->", rr.status_code, len(rr.content), rr.text[:700].replace("\n", " "))
    except Exception as e: p("POST", ep, "ERR", type(e).__name__, str(e)[:100])
post("generation/data/powerplant-list", {"startDate": "2026-09-01T00:00:00+03:00", "endDate": "2026-09-01T00:00:00+03:00"})
post("generation/data/powerplant-list", {})
post("generation/data/powerplant-list-for-date-range", {"startDate": "2026-09-01T00:00:00+03:00", "endDate": "2026-09-30T00:00:00+03:00"})
post("generation/data/licensed-powerplant-investment-list", {})
post("generation/data/licensed-powerplant-investment-list", {"page": {"number": 1, "size": 5}})
post("renewables/data/new-installed-capacity", {"period": "2026-09-01T00:00:00+03:00"})
post("renewables/data/old-installed-capacity", {"period": "2026-09-01T00:00:00+03:00"})
post("generation/data/aic", {"startDate": "2026-10-01T00:00:00+03:00", "endDate": "2026-10-01T00:00:00+03:00"})
post("generation/data/sbfgp", {"startDate": "2026-10-01T00:00:00+03:00", "endDate": "2026-10-01T00:00:00+03:00"})
p("=== EPDK monthly electricity reports")
r = get("https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar")
if r is not None:
    t = r.text; seen = []
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']*DownloadDocument[^"\']*)["\'][^>]*>(.*?)</a>', t, re.S):
        ctx = clean(t[max(0, m.start() - 400):m.end() + 100], 300)
        seen.append((html.unescape(m.group(1)), ctx))
    p("download links", len(seen))
    for u, c in seen[:40]: p("  ", u, "|", c[-220:])
    if seen:
        rr = get(urljoin("https://www.epdk.gov.tr/", seen[0][0]))
        if rr is not None:
            p("first download:", rr.status_code, rr.headers.get("content-type"), rr.headers.get("content-disposition"), len(rr.content))
            if rr.content[:4] == b"%PDF":
                txt = pdf_text(rr.content, 12, 20000)
                for m in re.finditer(r"(?i)kurulu", txt):
                    p("   KURULU:", clean(txt[max(0, m.start() - 150):m.start() + 500], 650)); break
                p("   head:", clean(txt, 1200))
p("=== TEIAS")
r = get("https://www.teias.gov.tr/kurulu-guc-raporlari")
if r is not None:
    t = r.text
    for m in list(re.finditer(r"(?i)kurulu", t))[:12]: p("  ctx:", clean(t[max(0, m.start() - 150):m.start() + 250], 400))
    p("   file refs:", sorted(set(re.findall(r"[\w/\-\.%]+\.(?:xlsx?|pdf|zip)", t)))[:60])
    p("   hrefs /:", sorted(set(re.findall(r'href="(/[^"#]{3,80})"', t)))[:80])
p("=== HROTE")
r = get("https://files.hrote.hr/files/PDF/OIEIK/2023/Izvjestaji/Mjesecni_izvjestaj_prosinac_2023.pdf")
if r is not None and r.content[:4] == b"%PDF": p(pdf_text(r.content, 4, 3500))
links_hrote = get("https://www.hrote.hr/2025-godina")
if links_hrote is not None:
    p("2025-godina:", links_hrote.status_code, re.findall(r'href="([^"]+\.pdf)"', links_hrote.text)[:30])
r = get("https://www.hrote.hr/izvjestaji")
if r is not None:
    p("hrote monthly pdfs:", sorted(set(re.findall(r'href="([^"]*Mjesec[^"]*)"', r.text, re.I)))[:30])
p("=== Borzen")
r = get("https://borzen.si/en-us/electricity-production-support/data-and-reports/data-and-reports/pid/1195/mcat/1193/acat/2/artdatemonth/7-2026")
if r is not None:
    t = r.text; p("   status", r.status_code, "docs:", sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?))"', t)))[:20]); p("   text:", clean(t[t.find("mainContainer"):], 1500))
p("=== ADMIE file download api")
r = get("https://www.admie.gr/en/market/market-statistics/file-download-api")
if r is not None: p(r.status_code, clean(r.text[r.text.find("main-content"):], 2500))
for u in ("https://www.admie.gr/getFiletypeInfoEN", "https://www.admie.gr/getFiletypeInfo"):
    rr = get(u)
    if rr is not None: p(u, rr.status_code, rr.text[:3000].replace("\n", " "))
r = get("https://www.admie.gr/en/market/reports/monthly-energy-balance")
if r is not None:
    p("monthly-energy-balance", r.status_code, sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?|zip))"', r.text)))[:15])
p("=== SODO / DEDDIE")
r = get("https://www.sodo.si/sl/distibucija-elektricne-energije/stanje-vlog-za-samooskrbo")
if r is not None: p(r.status_code, clean(r.text[r.text.find("<main"):], 1500), re.findall(r'href="([^"]+\.(?:pdf|xlsx?))"', r.text)[:10])
for u in ("https://deddie.gr/en/services/non-interconnected-islands/monthly-reports-of-res-and-thermal-units-in-the-non-interconnected-islands/",
          "https://deddie.gr/en/services/non-interconnected-islands/statistics-on-res-stations-relating-to-the-w-or-wr-application-on-the-nii/"):
    r = get(u)
    if r is not None: p(u[-70:], r.status_code, re.findall(r'href="([^"]+\.(?:pdf|xlsx?|zip))"', r.text)[:12], clean(r.text[r.text.find("<main"):], 500))
