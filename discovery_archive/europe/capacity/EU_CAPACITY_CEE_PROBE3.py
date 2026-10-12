"""
Probe 3 (CEE capacity group): second-level pages and file inspection.
SK installedpower.sepsas.sk; RO anre.ro/puteri-instalate; PL URE OZE pages + raporty.pse.pl bundle;
CZ OTE / ERU / CEPS; HU MAVIR pages; BG ESO / DKER; CH BFE production-plant register zip.
Output: results/probe3.txt
"""
import io, json, os, re, time, zipfile
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe3.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=H, timeout=(10, 90), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
def grep(name, url, pat, ctx=120, maxn=12):
    r = get(url)
    if r is None: return None
    p(f"== {name} {url} -> {r.status_code} {len(r.content)} {r.headers.get('content-type')}")
    n = 0
    for m in re.finditer(pat, r.text, re.I):
        s = max(0, m.start() - ctx); p("    ..", sq(r.text[s:m.end() + ctx], 2 * ctx + 60)); n += 1
        if n >= maxn: break
    return r
IP = r"instal|kapacit|výkon|vykon|putere|teljes|beép|installed|мощност|capacity|xls|csv|json|api"
p("=== SK")
r = grep("SEPS installedpower", "https://installedpower.sepsas.sk/", IP)
if r is not None:
    p(sq(r.text, 1500))
    for s in re.findall(r'(?:src|href)=["\']([^"\']+)["\']', r.text): p("    asset", s)
for u in ["https://installedpower.sepsas.sk/api", "https://installedpower.sepsas.sk/swagger", "https://installedpower.sepsas.sk/data.json"]:
    r = get(u)
    if r is not None: p(r.status_code, u, sq(r.text, 200))
p("=== RO")
r = grep("ANRE puteri", "https://anre.ro/puteri-instalate", IP, maxn=30)
if r is not None: p(sq(r.text, 1500))
p("=== PL")
r = grep("URE moc zainstalowana", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/5753,Moc-zainstalowana-MW.html", r"xls|pdf|stan na|Moc", maxn=40)
r = grep("URE instalacje 30.06.2026", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/8108,Instalacje-odnawialnych-zrodel-energii-stan-na-30-czerwca-2026-r.html", r"xls|pdf|stan na|Moc|MW", maxn=40)
r = get("https://raporty.pse.pl/")
if r is not None:
    p("raporty.pse.pl", r.status_code, sq(r.text, 600))
    for s in re.findall(r'src=["\']([^"\']+\.js)["\']', r.text):
        js = get("https://raporty.pse.pl" + s if s.startswith("/") else s)
        if js is not None:
            p("   js", s, len(js.text))
            names = sorted(set(re.findall(r'["\']((?:his|kse|pdgsz|pk5l|mocy|moc|gen|zdarz|dane|pomoc|zap|rez|ogr)[a-z0-9\-]*)["\']', js.text)))
            p("   candidate names:", names[:120])
            for m in re.finditer(r"zainstal|installed", js.text, re.I): p("   ..", sq(js.text[max(0, m.start()-150):m.end()+150], 330))
p("=== PL API entity probes")
for e in ["his-wlk-cal", "mocy-zainst", "moc-zainstalowana", "inst-cap", "installed-capacity", "gen-jw", "gen-jg", "pdgsz", "his-gen", "kse-load", "poze-redoze", "poze-redoze-his"]:
    r = get(f"https://api.raporty.pse.pl/api/{e}?$top=1")
    if r is not None: p(r.status_code, e, sq(r.text, 260))
p("=== CZ")
grep("OTE monthly electricity", "https://www.ote-cr.cz/en/statistics/monthly-report-electricity", r"xls|pdf|instal|capacity", maxn=15)
for u in ["https://www.eru.gov.cz/elektrina-mesicni-zpravy-o-provozu-es-cr", "https://www.eru.gov.cz/mesicni-zpravy-o-provozu-es-cr", "https://www.eru.gov.cz/en/electricity-monthly-reports-on-the-operation-of-the-czech-power-system"]:
    r = get(u); 
    if r is not None: p(r.status_code, u, len(r.content))
r = get("https://www.eru.gov.cz/")
if r is not None:
    for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,80})<', r.text):
        if re.search(r"zpr[aá]v|statist|elektr|obnov|oze|výrob|vyrob", m.group(1) + m.group(2), re.I): p("    ERU link", m.group(2).strip(), m.group(1)[:120])
r = get("https://www.ceps.cz/en/data")
if r is not None:
    p("CEPS data", r.status_code, len(r.content))
    for m in re.finditer(r"instal|Instal", r.text): p("   ..", sq(r.text[max(0, m.start()-150):m.end()+200], 380))
    for m in re.finditer(r'href=["\']([^"\']*(?:data|Data)[^"\']*)["\']', r.text): p("   link", m.group(1)[:140])
p("=== HU")
for u in ["https://mavir.hu/web/mavir-en/hungarian-power-system-daily-weekly-monthly-annual-data", "https://mavir.hu/web/mavir-en/renewable-energy-sources", "https://mavir.hu/web/mavir-en/monthly-reports"]:
    r = get(u)
    if r is None: continue
    p("==", u, r.status_code, len(r.content))
    for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,100})<', r.text):
        if re.search(r"xls|csv|zip|document|instal|capacity|report|data", m.group(1) + m.group(2), re.I) and "javascript" not in m.group(1): p("    ", m.group(2).strip()[:70], "|", m.group(1)[:150])
p("=== BG")
grep("ESO generation_per_day", "https://www.eso.bg/doc/?generation_per_day", r"instal|мощност|xls|csv", maxn=10)
r = get("https://www.eso.bg/")
if r is not None:
    for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,100})<', r.text):
        if re.search(r"данни|data|doc|мощност|генерац|instal|отчет|report", m.group(1) + m.group(2), re.I): p("    ESO", m.group(2).strip()[:70], "|", m.group(1)[:130])
r = get("https://www.dker.bg/")
if r is not None:
    for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,100})<', r.text):
        p("    DKER", m.group(2).strip()[:70], "|", m.group(1)[:130])
p("=== CH register")
for u in ["https://data.geo.admin.ch/ch.bfe.elektrizitaetsproduktionsanlagen/csv/2056/ch.bfe.elektrizitaetsproduktionsanlagen.zip",
          "https://data.geo.admin.ch/ch.bfe.elektrizitaetsproduktionsanlagen/csv/2056/ch.bfe.elektrizitaetsproduktionsanlagen_2056.csv.zip"]:
    r = get(u)
    if r is None or r.status_code != 200: p("fail", u, r.status_code if r is not None else None); continue
    p("OK", u, len(r.content), r.headers.get("last-modified"), r.headers.get("etag"))
    z = zipfile.ZipFile(io.BytesIO(r.content)); p(z.namelist())
    import pandas as pd
    for n in z.namelist():
        if n.endswith(".csv"):
            d = pd.read_csv(z.open(n), nrows=5, encoding="utf-8"); p(n, list(d.columns)); p(d.T.to_string()[:3000])
            d = pd.read_csv(z.open(n), encoding="utf-8"); p("rows", len(d))
            for c in d.columns:
                if d[c].nunique() < 30: p("   ", c, d[c].value_counts(dropna=False).head(25).to_dict())
    break
r = get("https://www.bfe.admin.ch/elektrizitaetsproduktionsanlagen")
if r is not None: p("bfe page", r.status_code, sq(r.text, 300))
