"""
Probe 7 (CEE capacity group). Output: results/probe7.txt
 - CZ ERU quarterly/annual report xlsx attachments (sheet dump)
 - PL URE 'Instalacje OZE' quarterly xlsx: layout and sums by technology code, every release
 - PL PSE monthly reports page documents
 - Power BI cluster routing (SK/RO) with curl_cffi and plain requests
 - BG ESO operational pages
"""
import io, json, os, re
import requests
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe7.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=kw.pop("headers", H), timeout=(10, 120), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
p("=== CZ")
for u in ["https://eru.gov.cz/sites/default/files/obsah/prilohy/eruelektroq22026.xlsx", "https://eru.gov.cz/sites/default/files/obsah/prilohy/eruelektro2025_0.xlsx"]:
    x = get(u)
    if x is None or x.status_code != 200: p("fail", u, x.status_code if x is not None else None); continue
    xl = pd.ExcelFile(io.BytesIO(x.content)); p("XLSX", u, len(x.content), x.headers.get("last-modified"), xl.sheet_names)
    for sh in xl.sheet_names:
        dd = xl.parse(sh, header=None); p("   --", sh, dd.shape)
        if re.search(r"instal|výkon|vykon|kapac", sh, re.I) or dd.shape[0] < 45: p(dd.head(45).to_string()[:4500])
p("=== PL URE")
r = get("https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/8108,Instalacje-odnawialnych-zrodel-energii-stan-na-30-czerwca-2026-r.html")
rels = re.findall(r'href="(/download/9/\d+/[^"]+)"> <strong>([^<]+)</strong>', r.text) if r is not None else []
p(len(rels), "releases")
for href, lab in rels:
    x = get("https://www.ure.gov.pl" + href)
    if x is None or x.status_code != 200: p("fail", lab, href); continue
    try:
        xl = pd.ExcelFile(io.BytesIO(x.content)); d = xl.parse(xl.sheet_names[0], header=None)
    except Exception as e: p("xl err", lab, e); continue
    p("REL", sq(lab, 70), href, len(x.content), xl.sheet_names, d.shape)
    p("   head:", sq(d.head(4).to_string(), 700))
    # find MW and code columns
    hdr = None
    for i in range(min(8, len(d))):
        row = [str(v) for v in d.iloc[i].tolist()]
        if any("Moc" in v for v in row) and any("Rodzaj" in v for v in row): hdr = i; break
    if hdr is None: p("   no header"); continue
    row = [str(v) for v in d.iloc[hdr].tolist()]; ci = [k for k, v in enumerate(row) if "Moc" in v][0]; ck = [k for k, v in enumerate(row) if "Rodzaj" in v][0]
    dd = d.iloc[hdr + 1:]; mw = pd.to_numeric(dd[ci].astype(str).str.replace(",", ".").str.replace(r"[^0-9.]", "", regex=True), errors="coerce")
    g = mw.groupby(dd[ck].astype(str).str.strip()).agg(["count", "sum"]).round(1); p("   sums:", g.to_dict("index"))
p("=== PL PSE monthly")
for u in ["https://www.pse.pl/dane-systemowe/funkcjonowanie-kse/raporty-miesieczne-z-funkcjonowania-kse/raporty-miesieczne", "https://www.pse.pl/raporty-historyczne", "https://www.pse.pl/mapa-raportow"]:
    r = get(u)
    if r is None: continue
    p("==", u, r.status_code, len(r.content))
    for h, t in re.findall(r'href=["\']([^"\']*(?:documents|\.pdf|\.xls)[^"\']*)["\'][^>]*>\s*([^<]{0,100})<', r.text)[:40]: p("   ", sq(t, 80), "|", h[:170])
    for h, t in re.findall(r'href=["\']([^"\']+)["\'][^>]*>\s*([^<]*(?:moc|Moc|zainstal|struktur)[^<]*)<', r.text)[:20]: p("   MOC", sq(t, 80), "|", h[:170])
p("=== Power BI routing")
tenant = "6dfa2abc-8bb8-4557-855c-e532cacb5122"
for h in [{"User-Agent": "python-requests/2.31"}, {"User-Agent": H["User-Agent"], "Accept": "application/json", "Origin": "https://app.powerbi.com"}, {"User-Agent": H["User-Agent"]}]:
    r = get(f"https://api.powerbi.com/public/routing/cluster/{tenant}", headers=h); p("requests", h.get("User-Agent")[:20], r.status_code if r is not None else None, sq(r.text, 200) if r is not None else "")
try:
    from curl_cffi import requests as cr
    for imp in ["chrome", "chrome124", "safari17_0", "firefox"]:
        try:
            r = cr.get(f"https://api.powerbi.com/public/routing/cluster/{tenant}", impersonate=imp, timeout=60); p("curl_cffi", imp, r.status_code, sq(r.text, 200))
        except Exception as e: p("curl_cffi", imp, "ERR", str(e)[:100])
    r = cr.get("https://installedpower.sepsas.sk/", impersonate="chrome", timeout=60); p("sepsas curl_cffi", r.status_code, len(r.text))
except Exception as e: p("curl_cffi n/a", e)
for u in ["https://app.powerbi.com/", "https://wabi-west-europe-api.analysis.windows.net/", "https://api.powerbi.com/"]:
    r = get(u); p("GET", u, r.status_code if r is not None else None, sq(r.text, 150) if r is not None else "")
p("=== BG ESO")
for u in ["https://www.eso.bg/doc/?460", "https://www.eso.bg/doc/?39", "https://www.eso.bg/doc/?joining", "https://www.eso.bg/doc/?investments"]:
    r = get(u)
    if r is None: continue
    p("==", u, r.status_code, len(r.content))
    for h, t in re.findall(r'href=["\']([^"\']+)["\'][^>]*>\s*([^<]{3,120})<', r.text):
        if re.search(r"мощност|инсталир|fileObj|xls|капацитет|ВЕИ|РЕС|производ", h + t, re.I): p("   ", sq(t, 90), "|", h[:120])
