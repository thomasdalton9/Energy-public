"""
Probe 8 (CEE capacity group). Output: results/probe8.txt
 - CZ: ERU quarterly / annual report slugs 2021-2026 -> xlsx attachments, the 'INSTALOVANÝ VÝKON' sheet of each
 - PL: URE 'Instalacje OZE' xlsx releases: layout, sums by technology code
 - SK / RO: Power BI public API called directly on the wabi-*-api cluster hosts (the routing endpoint answers 403 from Actions)
"""
import base64, io, json, os, re, uuid
import requests
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe8.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=kw.pop("headers", H), timeout=(10, 120), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
p("=== CZ")
ROM = {1: "i", 2: "ii", 3: "iii", 4: "iv"}
slugs = []
for y in range(2021, 2027):
    slugs.append(f"rocni-zprava-o-provozu-elektrizacni-soustavy-cr-pro-rok-{y}")
    slugs.append(f"rocni-zprava-o-provozu-elektrizacni-soustavy-cr-za-rok-{y}")
    for q in range(1, 5): slugs.append(f"ctvrtletni-zprava-o-provozu-elektrizacni-soustavy-cr-za-{ROM[q]}-ctvrtleti-{y}")
for s in slugs:
    r = get("https://eru.gov.cz/" + s)
    if r is None or r.status_code != 200: p(s, r.status_code if r is not None else None); continue
    fl = sorted(set(re.findall(r'href=["\']([^"\']+\.xlsx?)["\']', r.text, re.I)))
    p(s, 200, fl)
    for f in fl[:1]:
        x = get("https://eru.gov.cz" + f if f.startswith("/") else f)
        if x is None or x.status_code != 200: continue
        try: xl = pd.ExcelFile(io.BytesIO(x.content))
        except Exception as e: p("   xl err", e); continue
        for sh in xl.sheet_names:
            d = xl.parse(sh, header=None, nrows=22)
            if d.shape[0] and "INSTALOVAN" in str(d.iloc[0, 0]).upper() and "KRAJ" in str(d.iloc[0, 0]).upper():
                p("   sheet", sh, "of", f); p(d.iloc[:16, :14].to_string()[:2600])
p("=== PL URE")
r = get("https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/8108,Instalacje-odnawialnych-zrodel-energii-stan-na-30-czerwca-2026-r.html")
rels = re.findall(r'href="(/download/9/\d+/[^"]+)">\s*<strong>\s*([^<]+?)\s*</strong>', r.text) if r is not None else []
p(len(rels), "releases")
for href, lab in rels:
    x = get("https://www.ure.gov.pl" + href)
    if x is None or x.status_code != 200: p("fail", lab, href); continue
    try:
        xl = pd.ExcelFile(io.BytesIO(x.content)); d = xl.parse(xl.sheet_names[0], header=None)
    except Exception as e: p("xl err", lab, e); continue
    p("REL", sq(lab, 70), href, len(x.content), xl.sheet_names, d.shape)
    hdr = None
    for i in range(min(10, len(d))):
        row = [str(v) for v in d.iloc[i].tolist()]
        if any("Moc" in v for v in row): hdr = i; break
    if hdr is None: p("   no header", sq(d.head(3).to_string(), 500)); continue
    row = [str(v) for v in d.iloc[hdr].tolist()]; p("   header", [sq(v, 40) for v in row])
    ci = [k for k, v in enumerate(row) if "Moc" in v][0]; ck = [k for k, v in enumerate(row) if "Rodzaj" in v or "OZE" in v or "Technolog" in v]
    p("   cols", ci, ck)
    dd = d.iloc[hdr + 1:]; mw = pd.to_numeric(dd[ci].astype(str).str.replace(",", ".").str.replace(r"[^0-9.]", "", regex=True), errors="coerce")
    for k in ck[:1]:
        g = mw.groupby(dd[k].astype(str).str.strip()).agg(["count", "sum"]).round(1); p("   sums by", k, g.to_dict("index"))
    p("   total MW", round(mw.sum(), 1))
p("=== Power BI direct")
def pbi(name, k, t, hosts):
    for host in hosts:
        base = f"https://{host}.analysis.windows.net/"
        hdr = {**H, "X-PowerBI-ResourceKey": k, "ActivityId": str(uuid.uuid4()), "RequestId": str(uuid.uuid4()), "Accept": "application/json"}
        me = get(f"{base}public/reports/{k}/modelsAndExploration?preferReadOnlySession=true", headers=hdr)
        p(name, host, me.status_code if me is not None else None, len(me.content) if me is not None else 0, sq(me.text, 160) if me is not None and me.status_code != 200 else "")
        if me is not None and me.status_code == 200: break
    else: return
    j = me.json(); json.dump(j, open(os.path.join(OUT, f"pbi_{name}_explore.json"), "w"))
    models = j["models"]; p("models", [(m["id"], m.get("dbName")) for m in models][:3], "keys", list(j.keys()), list(j["exploration"].keys()))
    secs = j["exploration"]["sections"]; qs = []
    for s in secs:
        for vc in s.get("visualContainers", []):
            try: cfg = json.loads(vc["config"])
            except Exception: continue
            sv = cfg.get("singleVisual", {}); pq = sv.get("prototypeQuery")
            title = ""
            try: title = sv["vcObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
            except Exception: pass
            p("  visual", s["displayName"], sv.get("visualType"), title, len(pq["Select"]) if pq else 0, [x.get("Name") for x in pq["Select"]][:8] if pq else "")
            if pq: qs.append((s["displayName"], sv.get("visualType"), title, pq))
    rep_id = j["exploration"].get("report", {}).get("objectId") or j["exploration"].get("reportId")
    for sname, vt, title, pq in qs[:20]:
        n = len(pq["Select"])
        body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": pq, "Binding": {"Primary": {"Groupings": [{"Projections": list(range(n))}]}, "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": 30000}}}, "Version": 1}, "ExecutionMetricsKind": 1}}]}, "QueryId": "", "ApplicationContext": {"DatasetId": models[0].get("dbName") or str(models[0]["id"]), "Sources": [{"ReportId": rep_id}]}}], "cancelQueries": [], "modelId": models[0]["id"]}
        try:
            rr = requests.post(f"{base}public/reports/querydata?synchronous=true", headers={**hdr, "Content-Type": "application/json;charset=UTF-8"}, data=json.dumps(body), timeout=120)
            p("  QUERY", sname, vt, title, rr.status_code, len(rr.content), sq(rr.text, 700))
        except Exception as e: p("  QUERY err", e)
hosts = ["wabi-west-europe-api", "wabi-north-europe-api", "wabi-uk-south-api", "wabi-germany-west-central-api", "wabi-france-central-api"]
pbi("sk_seps", "c4f2ea34-d9d6-4963-bfa1-b1b1bd28a2ad", "6dfa2abc-8bb8-4557-855c-e532cacb5122", hosts)
pbi("ro_anre", "26133983-3d01-4513-8973-6f75b7d00733", "840aa2a0-7b45-4f9f-abd4-c6e7908704ac", hosts)
