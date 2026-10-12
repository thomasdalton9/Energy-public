"""
Probe 4 (SE Europe capacity). Output: results/se_probe4.txt
 - EPDK monthly electricity sector report PDFs: installed-capacity table text for several months (2026-07, 2025-01, 2023-01, 2021-01, 2019)
 - ADMIE file types mentioning installed/RES/capacity; ADMIE monthly energy report text sample
 - SODO self-supply xlsx, Borzen page hrefs, TEIAS YTBS statistics page, ERE/MEPSO/ZRRE/EMS report pages
"""
import os, re, sys, json, html, io
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en,tr;q=0.8,el;q=0.7,hr;q=0.6"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe4.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(10, 120), **kw)
        if r.status_code in (403, 503, 429, 412): raise RuntimeError(f"http {r.status_code}")
        return r
    except Exception as e:
        try:
            from curl_cffi import requests as cr
            return cr.get(url, impersonate="chrome120", timeout=120, **kw)
        except Exception as e2:
            p("ERR", url, str(e)[:60], "|", str(e2)[:60]); return None
def clean(s, n=600): return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s)))[:n]
import fitz
p("=== EPDK")
r = get("https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar")
items = []
if r is not None:
    t = r.text
    for m in re.finditer(r'data-id="(\d+)"[^>]*>\s*([^<]*Elektrik Piyasası[^<]*Sektör Rapor[^<]*)<.*?href="([^"]*DownloadDocument[^"]*)"', t, re.S):
        items.append((clean(m.group(2), 90).strip(), html.unescape(m.group(3))))
    p("items", len(items), items[:3], items[-3:])
want = ["2026 Yılı Elektrik Piyasası Temmuz", "2025 Yılı Elektrik Piyasası Ocak", "2023 Yılı Elektrik Piyasası Ocak", "2021 Yılı Elektrik Piyasası Ocak", "2019 Yılı Elektrik Piyasası Ocak"]
for w in want:
    hit = [x for x in items if x[0].startswith(w)]
    if not hit: p("no item", w); continue
    rr = get(urljoin("https://www.epdk.gov.tr/", hit[0][1]))
    if rr is None or rr.content[:4] != b"%PDF": p("not pdf", w, rr.status_code if rr is not None else None, rr.content[:100] if rr is not None else ""); continue
    d = fitz.open(stream=rr.content, filetype="pdf"); p("-----", hit[0][0], "pages", len(d), "bytes", len(rr.content))
    shown = 0
    for i in range(len(d)):
        tx = d[i].get_text()
        if re.search(r"(?i)kurulu\s*g[uü]c", tx) and re.search(r"(?i)(do[gğ]al\s*gaz|linyit|rüzgar|r[uü]zg[aâ]r)", tx) and re.search(r"\d{2,3}[\.,]\d{3}", tx):
            p(f"  [page {i+1}]", re.sub(r"[ \t]+", " ", tx)[:3200].replace("\n", " | ")); shown += 1
            if shown >= 3: break
p("=== ADMIE file types")
r = get("https://www.admie.gr/getFiletypeInfoEN")
if r is not None:
    try:
        for x in r.json():
            s = json.dumps(x)
            if re.search(r"(?i)install|capacity|RES|CHP|monthly|unit|registry|plant", s): p("  ", x["filetype"], "|", x["data_type"], "|", x["period_covered"], "|", x["publication_frequency"])
    except Exception as e: p("json fail", e)
r = get("https://www.admie.gr/sites/default/files/attached-files/type-file/2026/02/Energy_Report_202512_v2_en.pdf")
if r is not None and r.content[:4] == b"%PDF":
    d = fitz.open(stream=r.content, filetype="pdf"); p("ADMIE energy report pages", len(d))
    for i in range(len(d)):
        tx = d[i].get_text()
        if re.search(r"(?i)installed", tx): p(f"  [p{i+1}]", re.sub(r"\s+", " ", tx)[:1800])
p("=== SODO xlsx")
r = get("https://www.sodo.si/sl/distibucija-elektricne-energije/stanje-vlog-za-samooskrbo")
if r is not None:
    for u in re.findall(r'href="([^"]+\.xlsx?)"', r.text)[:2]:
        rr = get(u)
        if rr is None: continue
        import pandas as pd
        x = pd.ExcelFile(io.BytesIO(rr.content)); p("xlsx sheets", x.sheet_names, len(rr.content))
        for sn in x.sheet_names[:4]:
            df = x.parse(sn, header=None); p("  sheet", sn, df.shape); p(df.head(25).to_string()[:2500]); p(df.tail(6).to_string()[:900])
p("=== Borzen hrefs")
r = get("https://borzen.si/en-us/electricity-production-support/data-and-reports/data-and-reports/pid/1195/mcat/1193/acat/2/artdatemonth/7-2026")
if r is not None:
    hs = [h for h in re.findall(r'href="([^"]+)"', r.text) if re.search(r"(?i)download|link|\.pdf|\.xls|dnn|file|portals|poro|report", h)]
    p(sorted(set(hs))[:50])
    i = r.text.find("July"); p(clean(r.text[r.text.rfind("2026, July"):][:3000], 1500))
p("=== TEIAS YTBS")
for u in ("https://ytbsbilgi.teias.gov.tr/ytbsbilgi/frm_istatistikler.jsf", "https://ytbsbilgi.teias.gov.tr/ytbsbilgi/frm_kurulu_guc_raporlari.jsf", "https://ytbsbilgi.teias.gov.tr/ytbsbilgi/"):
    r = get(u)
    if r is not None:
        p(u, r.status_code, len(r.content), clean(r.text, 700)); p("  links:", sorted(set(re.findall(r'href="([^"#]+)"', r.text)))[:50])
p("=== Balkan report pages")
for u, pat in (("https://www.ere.gov.al/sq/tregu-i-energjise/raporte-dhe-te-dhena-te-tregut-te-energjise-elektrike", r"\.(pdf|xlsx?)"),
               ("https://www.mepso.com.mk/index.php/mk/doma/regulativa/izveshtai", r"\.(pdf|xlsx?)"),
               ("https://www.mepso.com.mk/index.php/mk/podruznici/oeps/informacii-i-izveshtai", r"\.(pdf|xlsx?)"),
               ("https://www.ero-ks.org/zrre/raportet-e-monitorimit", r"\.(pdf|xlsx?)"),
               ("https://ems.rs/dokumenta/", r"\.(pdf|xlsx?)"), ("https://www.derk.ba/ba/godisnji-izvjestaji-derk-a", r"\.(pdf|xlsx?)"),
               ("https://www.nosbih.ba/en/", r"\.(pdf|xlsx?)"), ("https://www.nosbih.ba/bs/", r".")):
    r = get(u)
    if r is None: continue
    hs = []
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S):
        h = urljoin(u, html.unescape(m.group(1)))
        if re.search(pat, h, re.I): hs.append(clean(m.group(2), 60) + " -> " + h[:150])
    p("==", u, r.status_code, len(hs)); [p("   ", x) for x in hs[:25]]
