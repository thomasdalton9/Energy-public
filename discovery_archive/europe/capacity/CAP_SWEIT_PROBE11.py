"""Probe 11: Spain REData widget generacion/potencia-instalada-generacion (renamed; works), Terna download-centre records (GET f=xlsx / POST).
Probe 10 findings: apidatos.ree.es/es/datos/generacion/potencia-instalada-generacion?start_date=..&end_date=..&time_trunc=month returns HTTP 200 JSON
('Estructura de potencia instalada', id gen4); the old name 'potencia-instalada' returns HTTP 500. Terna: the page POSTs JSON {filterDataset, filterYear,
filterMonth, orderByColumn, orderByDir, pageSize, pageIndex, db:'dati'} to the records endpoint; the xlsx button is a GET with f=xlsx."""
import io, json, re, subprocess, sys
import pandas as pd
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=(), out=None, post=None):
    cmd = ["curl", "-sL", "-m", "120", "-A", UA] + [x for h in hdr for x in ("-H", h)]
    if post is not None:
        cmd += ["-X", "POST", "--data", post]
    if out:
        cmd += ["-o", out]
    cmd += ["-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=150)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u[:300], "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

print("#### A. REData potencia-instalada-generacion")
H = ("Accept: application/json",)
for geo in ["", "&geo_trunc=electric_system&geo_limit=peninsular&geo_ids=8741"]:
    b = get("https://apidatos.ree.es/es/datos/generacion/potencia-instalada-generacion?start_date=2020-12-01T00:00&end_date=2026-09-30T23:59&time_trunc=month" + geo, show=False, hdr=H)
    try:
        j = json.loads(b)
    except Exception:
        print("not json", b[:200]); continue
    print("geo:", geo or "(none)", "| last-update", j["data"]["attributes"].get("last-update"), "| meta", json.dumps(j["data"].get("meta"))[:200])
    for it in j["included"]:
        a = it["attributes"]
        v = a.get("values", [])
        print(f"  {it['type']!r:40} id={it['id']} group={it.get('groupId')} title={a.get('title')!r} magnitude={a.get('magnitude')} n={len(v)} first={v[0] if v else None} last={v[-1] if v else None}")
    tot = [x for x in j["included"] if "otal" in str(x["attributes"].get("title"))]
    print("totals:", [(x["attributes"]["title"], x["attributes"]["values"][-1]) for x in tot])
    print("keys of an item:", list(j["included"][0]["attributes"].keys()), list(j.keys()))
    if not geo:
        print("sample dates:", [x["datetime"] for x in j["included"][0]["attributes"]["values"]][:4], "...", [x["datetime"] for x in j["included"][0]["attributes"]["values"]][-3:])

print("\n#### B. Terna records")
API = "https://dati.terna.it/api/sitecore/dati/downloadcenter/records"
for ds in ["CapacityRenewableSources", "CapacityThermal", "GenerationPlants"]:
    body = json.dumps({"filterDataset": ds, "filterYear": "2025", "orderByColumn": "Anno", "orderByDir": "desc", "pageSize": "10", "pageIndex": "0", "db": "dati"})
    get(API, 1800, hdr=("Content-Type: application/json", "Accept: application/json"), post=body)
    get(f"{API}?f=xlsx&filterDataset={ds}&filterYear=2025&pageSize=2000&pageIndex=0&db=dati", 300, out=f"/tmp/{ds}.xlsx")
    try:
        x = pd.ExcelFile(f"/tmp/{ds}.xlsx")
        for sh in x.sheet_names[:2]:
            d = x.parse(sh)
            print("XLSX", ds, sh, d.shape); print(d.head(8).to_string()[:1800])
    except Exception as e:
        print("xlsx ERR", ds, e)
