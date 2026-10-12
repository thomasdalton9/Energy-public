"""
Probe 5 (CEE capacity group). Output: results/probe5.txt
 - HU MAVIR installed-capacity PDFs (monthly change by production type, 2000-2026) -> text dump
 - CZ ERU quarterly/annual electricity system report attachments -> text search for installed capacity
 - SK / RO Power BI public embeds (installedpower.sepsas.sk, anre.ro/puteri-instalate): key, schema, visual queries, data
 - PL PSE monthly KSE report page attachments; URE OZE quarterly xlsx totals by type
 - BG DKER general information PDF, ESO statistical book
"""
import base64, io, json, os, re, uuid
import requests
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe5.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=kw.pop("headers", H), timeout=(10, 120), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
import fitz
def pdf_text(content, maxpages=8):
    d = fitz.open(stream=content, filetype="pdf"); out = []
    for i, pg in enumerate(d):
        if i >= maxpages: break
        out.append(f"--- page {i+1}\n" + pg.get_text())
    return d.page_count, "\n".join(out)
p("=== HU")
r = get("https://mavir.hu/web/mavir-en/energy-mix-power-plants-installed-capacity-data")
pdfs = re.findall(r'href=["\'](/documents/10262/[^"\']+\.pdf[^"\']*)["\']', r.text) if r is not None else []
p(len(pdfs), "pdf links"); 
for u in dict.fromkeys(pdfs):
    full = "https://mavir.hu" + u.replace("&amp;", "&"); x = get(full)
    if x is None or x.status_code != 200: p("fail", full, x.status_code if x is not None else None); continue
    n, t = pdf_text(x.content, 3); p("PDF", full[:200], len(x.content), "pages", n, x.headers.get("last-modified")); p(t[:3500])
p("=== CZ")
for slug in ["ctvrtletni-zprava-o-provozu-elektrizacni-soustavy-cr-za-ii-ctvrtleti-2026", "rocni-zprava-o-provozu-elektrizacni-soustavy-cr-pro-rok-2025", "mesicni-spotreba-elektriny-2026"]:
    u = "https://eru.gov.cz/" + slug; r = get(u)
    if r is None: continue
    p("==", u, r.status_code, len(r.content))
    fl = re.findall(r'href=["\']([^"\']+\.(?:pdf|xlsx?|csv|zip)[^"\']*)["\']', r.text, re.I); p("   files", fl[:10])
    for f in fl[:2]:
        full = f if f.startswith("http") else "https://eru.gov.cz" + f
        x = get(full)
        if x is not None and x.status_code == 200 and (".pdf" in f.lower()):
            n, t = pdf_text(x.content, 60); p("   PDF", full, len(x.content), "pages", n)
            for m in re.finditer(r"instalovan|Instalovan|výkon", t): p("      ..", sq(t[max(0, m.start()-120):m.end()+250], 380)); 
            break
        elif x is not None and x.status_code == 200:
            try:
                xl = pd.ExcelFile(io.BytesIO(x.content)); p("   xlsx", full, xl.sheet_names[:10]); p(xl.parse(xl.sheet_names[0], header=None).head(25).to_string()[:2500])
            except Exception as e: p("   xl err", e)
p("=== Power BI")
def pbi(name, r_param=None, key=None, tenant=None):
    if r_param:
        j = json.loads(base64.b64decode(r_param + "=" * (-len(r_param) % 4))); key, tenant = j["k"], j["t"]; p(name, "decoded", j)
    cl = get(f"https://api.powerbi.com/public/routing/cluster/{tenant}"); p("cluster", cl.status_code if cl is not None else None, sq(cl.text, 200) if cl is not None else "")
    base = cl.json()["FixedClusterUri"].replace("-redirect", "-api")
    hdr = {**H, "X-PowerBI-ResourceKey": key, "ActivityId": str(uuid.uuid4()), "RequestId": str(uuid.uuid4()), "Content-Type": "application/json;charset=UTF-8"}
    me = get(f"{base}public/reports/{key}/modelsAndExploration?preferReadOnlySession=true", headers=hdr)
    p("modelsAndExploration", me.status_code if me is not None else None, len(me.content) if me is not None else 0)
    if me is None or me.status_code != 200: p(sq(me.text, 300) if me is not None else ""); return
    j = me.json(); json.dump(j, open(os.path.join(OUT, f"pbi_{name}_explore.json"), "w"))
    models = j["models"]; p("models", [(m["id"], m.get("name")) for m in models])
    secs = j["exploration"]["sections"]; p("sections", [(s["displayName"], len(s.get("visualContainers", []))) for s in secs])
    q = []
    for s in secs:
        for vc in s.get("visualContainers", []):
            try: cfg = json.loads(vc["config"])
            except Exception: continue
            sv = cfg.get("singleVisual", {})
            pq = sv.get("prototypeQuery")
            title = ""
            try: title = sv["vcObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
            except Exception: pass
            p("  visual", s["displayName"], sv.get("visualType"), title, "query" if pq else "-", (len(pq.get("Select", [])) if pq else 0))
            if pq: q.append((s["displayName"], sv.get("visualType"), title, pq))
    model_id = models[0]["id"]; rep_id = j.get("exploration", {}).get("report", {}).get("objectId") or j.get("exploration", {}).get("reportId")
    for sname, vt, title, pq in q[:25]:
        n = len(pq["Select"])
        body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": pq, "Binding": {"Primary": {"Groupings": [{"Projections": list(range(n))}]}, "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": 30000}}}, "Version": 1}, "ExecutionMetricsKind": 1}}]}, "QueryId": "", "ApplicationContext": {"DatasetId": models[0].get("dbName") or str(model_id), "Sources": [{"ReportId": rep_id}]}}], "cancelQueries": [], "modelId": model_id}
        rr = requests.post(f"{base}public/reports/querydata?synchronous=true", headers=hdr, data=json.dumps(body), timeout=120)
        p("  QUERY", sname, vt, title, rr.status_code, len(rr.content)); 
        if rr.status_code == 200:
            fn = os.path.join(OUT, f"pbi_{name}_{abs(hash(title+sname))%10**6}.json"); open(fn, "w").write(rr.text); p("     ", sq(rr.text, 900))
        else: p("     ", sq(rr.text, 300))
r = get("https://anre.ro/puteri-instalate")
m = re.search(r"powerbi\.com/view\?r=([A-Za-z0-9_\-=]+)", r.text) if r is not None else None
if m: pbi("ro_anre", r_param=m.group(1))
p("--- SK page")
r = get("https://installedpower.sepsas.sk/")
if r is not None:
    t = r.text
    for m in re.finditer(r"resourceKey\s*=|ResourceKey|\"k\"|tenantId\s*=|embedUrl|powerbi\.com/view|[?&]r=", t): p("   ..", sq(t[max(0, m.start()-200):m.end()+300], 520))
    p("tail", sq(t[-1500:], 1500))
p("=== PL")
r = get("https://www.pse.pl/dane-systemowe/funkcjonowanie-kse/raporty-miesieczne-z-funkcjonowania-kse/raporty-miesieczne")
if r is not None:
    p("PSE monthly", r.status_code, len(r.content))
    fl = re.findall(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,100})<', r.text)
    for h, t in fl:
        if re.search(r"documents|pdf|xls|raport", h + t, re.I) and "javascript" not in h: p("   ", sq(t, 80), "|", h[:170])
p("=== BG")
x = get("https://www.dker.bg/uploads/2026/obshta_inf_26.pdf")
if x is not None and x.status_code == 200:
    n, t = pdf_text(x.content, 30); p("DKER obshta_inf", len(x.content), n)
    for m in re.finditer(r"инсталирана|мощност", t): p("    ..", sq(t[max(0, m.start()-100):m.end()+250], 350))
