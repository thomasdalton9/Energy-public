"""
Probe 5 (SE Europe capacity). Output: results/se_probe5.txt
 - EPDK monthly electricity sector report: correct item list, installed-capacity table text for several months
 - DAPEEP ELAPE monthly bulletin PDFs (direct), HROTE year pages, SODO self-supply sheet layout, Borzen quarterly PDF
"""
import os, re, sys, json, html, io
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en,tr;q=0.8,el;q=0.7,hr;q=0.6"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe5.txt"), "w")
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
            p("ERR", url[:120], str(e)[:60], "|", str(e2)[:60]); return None
def clean(s, n=600): return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s)))[:n]
import fitz
p("=== EPDK")
r = get("https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar")
items = []
if r is not None:
    t = r.text
    for m in re.finditer(r'href="([^"]*DownloadDocument\?id=[^"]+)"', t):
        pre = clean(t[max(0, m.start() - 700):m.start()], 700)
        mm = re.findall(r"(\d{4}) Yılı Elektrik Piyasası (\w+) Ayı Sektör Raporu", pre)
        items.append((mm[-1] if mm else None, html.unescape(m.group(1))))
    named = [x for x in items if x[0]]
    p("downloads", len(items), "named", len(named), "first", named[:2], "last", named[-3:])
    p("years", sorted(set(x[0][0] for x in named)))
    for pick in [("2026", "Temmuz"), ("2025", "Ocak"), ("2023", "Ocak"), ("2021", "Ocak"), ("2019", "Ocak")]:
        hit = [x for x in named if x[0] == pick]
        if not hit: p("no item", pick); continue
        rr = get(urljoin("https://www.epdk.gov.tr/", hit[0][1]))
        if rr is None or rr.content[:4] != b"%PDF": p("not pdf", pick, None if rr is None else rr.status_code); continue
        d = fitz.open(stream=rr.content, filetype="pdf"); p("-----", pick, "pages", len(d))
        shown = 0
        for i in range(len(d)):
            tx = d[i].get_text()
            if re.search(r"(?i)kurulu\s*g[üu]ç", tx) and re.search(r"(?i)(do[ğg]al\s*gaz)", tx) and re.search(r"(?i)(linyit)", tx) and re.search(r"(?i)(r[üu]zg[âa]r|hidro)", tx):
                p(f"  [page {i+1}]", re.sub(r"[ \t]+", " ", tx)[:3500].replace("\n", " | ")); shown += 1
                if shown >= 2: break
        if not shown: p("  no capacity table page found")
p("=== DAPEEP")
for u in ("https://www.dapeep.gr/wp-content/uploads/ELAPE/2020/07_NOV_2020_DELTIO_ELAPE_v1.0_01.02.2021_d.pdf", "https://www.dapeep.gr/", "https://www.dapeep.gr/en/", "https://www.dapeep.gr/category/deltia-elape/", "https://www.dapeep.gr/wp-content/uploads/ELAPE/2018/March_April__2018_DELTIO_ELAPE_v1.0_15.06.2018.pdf"):
    rr = get(u)
    if rr is None: continue
    p(u, rr.status_code, len(rr.content), rr.headers.get("content-type"))
    if rr.content[:4] == b"%PDF":
        d = fitz.open(stream=rr.content, filetype="pdf"); p("   pages", len(d))
        for i in range(len(d)):
            tx = d[i].get_text()
            if re.search(r"(?i)MW", tx) and re.search(r"(?i)(εγκατ|ισχύ)", tx): p(f"   [p{i+1}]", re.sub(r"\s+", " ", tx)[:1500]); break
    elif rr.status_code == 200: p("   hrefs", sorted(set(re.findall(r'href="([^"]+(?:ELAPE|elape|deltia|bulletin)[^"]*)"', rr.text)))[:30])
p("=== HROTE year pages")
for y in ("2026-godina", "2025-godina", "2024-godina", "2023-godina"):
    rr = get("https://www.hrote.hr/" + y)
    if rr is None: continue
    hs = sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?))"', rr.text)))
    p(y, rr.status_code, len(hs), hs[:20])
rr = get("https://www.hrote.hr/izvjestaji")
if rr is not None:
    p("izvjestaji hrefs w/ Mjes or 2024-26:", [h for h in re.findall(r'href="([^"]+)"', rr.text) if re.search(r"(?i)mjes|202[456]", h)][:30])
p("=== SODO layout")
rr = get("https://www.sodo.si/sl/distibucija-elektricne-energije/stanje-vlog-za-samooskrbo")
if rr is not None:
    u = re.findall(r'href="([^"]+\.xlsx?)"', rr.text)[0]
    x = get(u)
    import pandas as pd
    xl = pd.ExcelFile(io.BytesIO(x.content))
    for sn in (xl.sheet_names[0], xl.sheet_names[-1]):
        df = xl.parse(sn, header=None).dropna(how="all").dropna(axis=1, how="all")
        p("sheet", sn, df.shape)
        for _, row in df.iterrows(): p("   ", " | ".join(str(v)[:60] for v in row.tolist() if str(v) != "nan"))
p("=== Borzen quarterly")
u = "https://borzen.si/Portals/0/Poročila CP/2026/2026-6M-CP-objava_ANG_Q2.pdf"
rr = get(u)
if rr is not None and rr.content[:4] == b"%PDF":
    d = fitz.open(stream=rr.content, filetype="pdf"); p("pages", len(d))
    for i in range(min(len(d), 12)):
        tx = d[i].get_text()
        if re.search(r"(?i)(installed|MW)", tx): p(f"[p{i+1}]", re.sub(r"\s+", " ", tx)[:1800])
else: p("borzen pdf", None if rr is None else rr.status_code)
