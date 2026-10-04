"""CEE probe 5: Amber Grid Power BI querydata replay (Lithuania gas consumption). Writes probe_cee_out/amber5.txt"""
import json, os
import requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "probe_cee_out"); os.makedirs(OUT, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
RID = "305e13f3-56da-4cce-813c-92a5ac0181e1"; MODEL = 1732958; DS = "dbe6a55e-3ce8-42b8-ace8-7a1df7e0cf58"; REPORT = 2182883
URL = "https://wabi-west-europe-d-primary-api.analysis.windows.net/public/reports/querydata?synchronous=true"
H = {"User-Agent": UA, "X-PowerBI-ResourceKey": RID, "Content-Type": "application/json;charset=UTF-8", "Accept": "application/json, text/plain, */*",
     "Origin": "https://app.powerbi.com", "Referer": "https://app.powerbi.com/"}
L = []
def col(src, prop): return {"Column": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}
def meas(src, prop): return {"Measure": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}
def q(select, frm, where=None, count=30000):
    pq = {"Version": 2, "From": frm, "Select": select}
    if where: pq["Where"] = where
    n = len(select)
    return {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": pq, "Binding": {"Primary": {"Groupings": [{"Projections": list(range(n))}]},
            "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": count}}}, "Version": 1}, "ExecutionMetricsKind": 1}}]},
            "QueryId": "", "ApplicationContext": {"DatasetId": DS, "Sources": [{"ReportId": str(RID)}]}}], "cancelQueries": [], "modelId": MODEL}
FROM = [{"Name": "d", "Entity": "Date", "Type": 0}, {"Name": "s", "Entity": "Σ Measures", "Type": 0}, {"Name": "v", "Entity": "VARTOTOJU_DUJU_SUNAUDOJIMAS", "Type": 0}]
PAV = [{"Condition": {"In": {"Expressions": [col("v", "pav")], "Values": [[{"Literal": {"Value": f"'{x}'"}}] for x in
       ("Domestic consumption", "Transmitted to directly connected cons.", "Transmitted to distribution systems")]}}}]

def parse(j):
    ds = j["results"][0]["result"]["data"]["dsr"]["DS"][0]
    dicts = ds.get("ValueDicts", {}); rows = []; prev = None; sch = None
    for r in ds["PH"][0]["DM0"]:
        if "S" in r: sch = r["S"]
        n = len(sch); R = r.get("R", 0); N = r.get("Ø", 0); C = list(r.get("C", [])); vals = []
        for i in range(n):
            if R >> i & 1: v = prev[i]
            elif N >> i & 1: v = None
            else:
                v = C.pop(0)
                dn = sch[i].get("DN")
                if dn is not None and isinstance(v, int) and dn in dicts: v = dicts[dn][v]
            vals.append(v)
        prev = vals; rows.append(vals)
    return sch, rows

def run(name, body):
    try:
        r = requests.post(URL, json=body, headers=H, timeout=120)
        L.append(f"== {name}: {r.status_code} len={len(r.text)}")
        if r.status_code != 200: L.append(r.text[:500]); return
        j = r.json()
        open(os.path.join(OUT, f"amber5_{name}.json"), "w").write(r.text[:300000])
        sch, rows = parse(j)
        L.append(f"schema={sch} rows={len(rows)}"); L.extend(str(x) for x in rows[:6]); L.append("..."); L.extend(str(x) for x in rows[-4:])
        return rows
    except Exception as e: L.append(f"{name}: {type(e).__name__} {e}")

sel = [col("d", "Date"), col("v", "pav"), meas("s", "SuvartojimoKiekis, GWh")]
run("daily_pav", q(sel, FROM, PAV))
run("daily_nopav", q([col("d", "Date"), meas("s", "SuvartojimoKiekis, GWh")], FROM, PAV))
run("raw", q([col("v", "prsk_data"), col("v", "pav"), col("v", "source"), col("v", "prsk_kiekis")], FROM[2:], PAV))
open(os.path.join(OUT, "amber5.txt"), "w").write("\n".join(L))
