"""
Probe 9 (CEE capacity group). Output: results/probe9.txt
 - SK: SEPS Power BI tables (allocated / used / free capacity by DSO and technology) as parsed rows
 - CH: BFE register - does Total-Initial power equal later PlantDetail increments?
 - PL: why the older URE releases failed; PSE report map for installed capacity
"""
import io, json, os, re, uuid, zipfile
import requests
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe9.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=kw.pop("headers", H), timeout=(10, 180), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
p("=== SK Power BI")
K = "c4f2ea34-d9d6-4963-bfa1-b1b1bd28a2ad"; base = "https://wabi-west-europe-api.analysis.windows.net/"
hdr = {**H, "X-PowerBI-ResourceKey": K, "ActivityId": str(uuid.uuid4()), "RequestId": str(uuid.uuid4()), "Accept": "application/json"}
j = get(f"{base}public/reports/{K}/modelsAndExploration?preferReadOnlySession=true", headers=hdr).json()
models = j["models"]; rep_id = j["exploration"].get("report", {}).get("objectId") or j["exploration"].get("reportId")
p("sections", [s["displayName"] for s in j["exploration"]["sections"]])
def decode(res):
    out = []
    for r in res["results"]:
        ds = r["result"]["data"]; desc = ds["descriptor"]["Select"]; names = [d["Name"] for d in desc]
        dsr = ds["dsr"]["DS"][0]; vd = dsr["ValueDicts"] if "ValueDicts" in dsr else {}
        for ph in dsr.get("PH", []):
            for key, rows in ph.items():
                prev = None
                for row in rows:
                    c = row.get("C", []); R = row.get("R", 0); Z = row.get("�", 0)
                    out.append((names, row))
    return out
for s in j["exploration"]["sections"]:
    for vc in s.get("visualContainers", []):
        try: cfg = json.loads(vc["config"])
        except Exception: continue
        sv = cfg.get("singleVisual", {}); pq = sv.get("prototypeQuery")
        if not pq or sv.get("visualType") not in ("pivotTable", "card", "tableEx", "table"): continue
        title = ""
        try: title = sv["vcObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
        except Exception: pass
        n = len(pq["Select"])
        body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": pq, "Binding": {"Primary": {"Groupings": [{"Projections": list(range(n))}]}, "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": 30000}}}, "Version": 1}, "ExecutionMetricsKind": 1}}]}, "QueryId": "", "ApplicationContext": {"DatasetId": models[0]["dbName"], "Sources": [{"ReportId": rep_id}]}}], "cancelQueries": [], "modelId": models[0]["id"]}
        rr = requests.post(f"{base}public/reports/querydata?synchronous=true", headers={**hdr, "Content-Type": "application/json;charset=UTF-8"}, data=json.dumps(body), timeout=120)
        try:
            res = rr.json()["results"][0]["result"]["data"]
            names = [d["Name"] for d in res["descriptor"]["Select"]]
            dsr = res["dsr"]["DS"][0]
            p("TABLE", title, names); p("   ValueDicts", json.dumps(dsr.get("ValueDicts", {}), ensure_ascii=False)[:600])
            for ph in dsr.get("PH", []):
                for k, rows in ph.items(): p("   ", k, json.dumps(rows, ensure_ascii=False)[:1500])
        except Exception as e: p("parse err", title, e, sq(rr.text, 300))
p("=== CH register reconcile")
r = get("https://data.geo.admin.ch/ch.bfe.elektrizitaetsproduktionsanlagen/csv/2056/ch.bfe.elektrizitaetsproduktionsanlagen.zip")
z = zipfile.ZipFile(io.BytesIO(r.content))
pl = pd.read_csv(z.open("ElectricityProductionPlant.csv")); det = pd.read_csv(z.open("PlantDetail.csv"))
pl["start"] = pd.to_datetime(pl["BeginningOfOperation"], errors="coerce"); det["Date"] = pd.to_datetime(det["Date"], errors="coerce")
d2 = det.merge(pl[["xtf_id", "start", "InitialPower", "TotalPower", "SubCategory"]], left_on="ElectricityProductionPlantR", right_on="xtf_id", how="left", suffixes=("", "_pl"))
p("detail rows", len(det), "matched plants", d2["start"].notna().sum())
pv = pl[pl.SubCategory == "subcat_2"].copy()
ds = det.groupby("ElectricityProductionPlantR")["Power"].agg(["sum", "count"]); pv = pv.merge(ds, left_on="xtf_id", right_index=True, how="left")
p("PV plants with detail rows", int(pv["count"].notna().sum()), "of", len(pv), "; sum Total (MW)", round(pv.TotalPower.sum() / 1000), "sum Initial", round(pv.InitialPower.sum() / 1000), "sum detail", round(pv["sum"].sum() / 1000))
x = pv[pv["count"].notna()]
p("   among plants with detail: Total", round(x.TotalPower.sum() / 1000), "Initial", round(x.InitialPower.sum() / 1000), "detail", round(x["sum"].sum() / 1000), "| |detail-Total|<=1kW:", int(((x["sum"] - x.TotalPower).abs() <= 1).sum()), "of", len(x))
y = pv[pv.TotalPower != pv.InitialPower]
p("   plants with Total != Initial", len(y), "with detail", int(y["count"].notna().sum()), "; Total-Initial", round((y.TotalPower - y.InitialPower).sum() / 1000), "MW")
# detail rows after plant start
dd = d2[(d2.SubCategory == "subcat_2") & d2["start"].notna()]
late = dd[dd.Date > dd.start + pd.Timedelta(days=31)]
p("   detail rows dated >31d after plant start:", len(late), "sum MW", round(late.Power.sum() / 1000), "; rows total", len(dd))
p("   detail Date year hist:", dd.Date.dt.year.value_counts().sort_index().to_dict())
p("   plant start year hist (PV):", pv.start.dt.year.value_counts().sort_index().tail(12).to_dict())
p("   Initial+late detail vs Total (MW):", round((pv.InitialPower.sum() + late.Power.sum()) / 1000), "vs", round(pv.TotalPower.sum() / 1000))
for sub, nm in [("subcat_1", "hydro"), ("subcat_3", "wind"), ("subcat_4", "biomass"), ("subcat_8", "gas"), ("subcat_10", "waste"), ("subcat_6", "nuclear")]:
    s = pl[pl.SubCategory == sub]; p(nm, "n", len(s), "Total", round(s.TotalPower.sum() / 1000), "Initial", round(s.InitialPower.sum() / 1000), "Total!=Initial", int((s.TotalPower != s.InitialPower).sum()), "start years", s.start.dt.year.value_counts().sort_index().tail(8).to_dict())
p("   PlantCategory pumped (plantcat_6) start years:", pl[pl.PlantCategory == "plantcat_6"].start.dt.year.value_counts().sort_index().tail(12).to_dict())
p("=== PL older URE")
for u in ["https://www.ure.gov.pl/download/9/13999/MAPAna30-06-2023.xlsx", "https://www.ure.gov.pl/download/9/13797/MAPAna31-12-2022nastronezprzypisem.xlsx", "https://www.ure.gov.pl/download/9/12809/file.file", "https://www.ure.gov.pl/download/9/11882/MapaOZEnastroneURE31-12-2020.xlsx"]:
    r = get(u); p(u, r.status_code if r is not None else None, len(r.content) if r is not None else 0, r.headers.get("content-type") if r is not None else "")
    if r is not None and r.status_code == 200:
        try:
            xl = pd.ExcelFile(io.BytesIO(r.content)); d = xl.parse(xl.sheet_names[0], header=None); p("   ", xl.sheet_names, d.shape); p(d.head(4).to_string()[:700])
        except Exception as e: p("   xl err", e)
p("=== PL PSE report map")
r = get("https://www.pse.pl/mapa-raportow")
if r is not None:
    for h, t in re.findall(r'href=["\']([^"\']+)["\'][^>]*>\s*([^<]{3,120})<', r.text):
        if re.search(r"instal|struktur|zainstal|OZE|odnawial|fotowolt|źródeł|zrodel|miesi", h + t, re.I): p("   ", sq(t, 90), "|", h[:150])
