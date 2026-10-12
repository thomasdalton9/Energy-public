"""Probe 2 (12 Oct 2026): follow-ups to probe 1. Findings: see results note."""
import json, re
import requests
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
def get(url, **kw):
    try:
        return S.get(url, timeout=45, **kw)
    except Exception as e:  # noqa
        print(f"  ERR {url[:100]} {type(e).__name__}"); return None
def links(url, pat, n=40):
    r = get(url)
    if r is None: return
    print(f"[{r.status_code}] {url} len={len(r.text)}")
    seen = []
    for m in re.finditer(r'(?:href|src)=["\']([^"\']+)["\']', r.text):
        u = m.group(1)
        if re.search(pat, u, re.I) and u not in seen:
            seen.append(u)
    for u in seen[:n]: print("    ", u)
    for m in list(re.finditer(pat, re.sub(r"<[^>]+>", " ", r.text), re.I))[:6]:
        t = re.sub(r"<[^>]+>", " ", r.text)
        print("   txt:", re.sub(r"\s+", " ", t[max(0, m.start()-100):m.end()+160]))

print("=== DK")
try:
    r = get("https://api.energidataservice.dk/meta/dataset/CapacityPerMunicipality")
    j = r.json(); print(j.get("description"), "|", j.get("comment")); print([str(c.get("name"))+":"+str(c.get("description"))[:60] for c in j.get("columns", [])])
    r = get("https://api.energidataservice.dk/dataset/CapacityPerMunicipality?limit=1&sort=Month%20asc"); print(r.text[:400])
    r = get("https://api.energidataservice.dk/dataset/CapacityPerMunicipality?limit=0&aggregate=OffshoreWindCapacity%20sum&group_by=Month&sort=Month%20asc"); print(r.text[:300])
    for d in ["PowerSystemRightNow", "ProductionConsumptionSettlement", "Forecasts_5Min", "StorageCapacity"]:
        r = get(f"https://api.energidataservice.dk/meta/dataset/{d}"); 
        if r is not None and r.ok and r.text: print(d, r.json().get("description"))
    r = get("https://api.energidataservice.dk/meta/dataset"); 
    print([ (d.get("datasetName") if isinstance(d, dict) else d) for d in r.json()])
    
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== SE PxWeb tree")
try:
    base = "https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndighetens_statistikdatabas"
    def walk(path, depth=0):
        r = get(base + path)
        if r is None or not r.ok: print("  fail", path, getattr(r, "status_code", None)); return
        for it in r.json():
            t = it.get("text", ""); i = it["id"]
            if it["type"] == "l":
                if depth < 3: print("  " * depth, "[", path + "/" + i, "]", t); walk(path + "/" + i, depth + 1)
            else:
                if re.search(r"effekt|kapacitet|installer|anläggning|vind|sol", t, re.I): print("  " * depth, "T", path + "/" + i, t)
    walk("")
    
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== NO")
try:
    r = get("https://api.nve.no/web/Powerplant/GetHydroPowerPlantsInOperation")
    j = r.json(); print(len(j)); print(json.dumps(j[0], ensure_ascii=False)[:1500])
    import collections
    print(collections.Counter(x.get("VannKVType") for x in j))
    for u in ["GetWindPowerPlants", "GetWindPowerPlantsInOperation", "GetWindPowerplantsInOperation", "GetWindpowerPlantsInOperation", "GetWindPowerPlantsInOperation/", "GetHydroPowerPlants"]:
        r = get("https://api.nve.no/web/Powerplant/" + u); print(u, getattr(r, "status_code", None), (r.text[:120] if r is not None else ""))
    for u in ["https://api.nve.no/web/Powerplant/", "https://api.nve.no/web/swagger", "https://api.nve.no/doc/", "https://nve.no/umbraco/api/Powerplant/GetWindPowerPlantsInOperation", "https://gis3.nve.no/arcgis/rest/services/featureservice?f=json", "https://gis3.nve.no/map/rest/services/Mapservices?f=json"]:
        r = get(u); print(u, getattr(r, "status_code", None), (r.text[:900] if r is not None else ""))
    
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== FI")
try:
    links("https://tuulivoimayhdistys.fi/en/wind-power-in-finland/wind-power-statistics", r"xls|csv|pdf|statistic|capacity|MW", 25)
    for u in ["https://pxdata.stat.fi/PxWeb/api/v1/en/StatFin/ehk/", "https://pxdata.stat.fi/PxWeb/api/v1/en/StatFin/ehk", "https://pxdata.stat.fi/PxWeb/api/v1/en/StatFin/"]:
        r = get(u); print(u, getattr(r, "status_code", None), (r.text[:1500] if r is not None else ""))
    try:
        from curl_cffi import requests as cr
        for u in ["https://energiavirasto.fi/en/power-plant-register", "https://energiavirasto.fi/en/web/energiavirasto/voimalaitosrekisteri"]:
            x = cr.get(u, impersonate="chrome", timeout=40); print("cffi", u, x.status_code, re.findall(r'href="([^"]+\.xlsx?[^"]*)"', x.text)[:10])
    except Exception as e: print("cffi", e)
    for u in ["https://data.fingrid.fi/en/datasets/", "https://data.fingrid.fi/en/datasets/268"]:
        r = get(u); print(u, getattr(r, "status_code", None), len(r.text) if r is not None else "")
    
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== EE")
try:
    r = get("https://dashboard.elering.ee/v3/api-docs"); j = r.json()
    for p, v in j["paths"].items(): print("  ", p, [x.get("summary") or x.get("operationId") for x in v.values()][:2])
    
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== LV")
try:
    links("https://www.ast.lv/en/electricity-market-review", r"xls|csv|capacity|generation|installed|open-?data|api", 40)
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
print("=== LT")
try:
    links("https://www.litgrid.eu/index.php/system-services/power-system/installed-generation-capacities/", r"xls|csv|capacity|generation|installed|open-?data|api|statistic", 40)
    for u in ["https://www.litgrid.eu/", "https://www.litgrid.eu/index.php/system-services/power-system/generation/3078"]:
        links(u, r"xls|csv|capacity|capacit|installed|open-?data|api|statistic|generation", 30)
except Exception as e:
    print('SECTION ERROR', type(e).__name__, e)
