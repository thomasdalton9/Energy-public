"""Probe 1 (12 Oct 2026): which official sources give monthly installed capacity for DK SE NO FI EE LV LT.
Findings are printed to the log; see discovery_archive/results/europe_capacity/ for the written summary."""
import json, re, sys, time
import requests

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"


def get(url, **kw):
    try:
        r = S.get(url, timeout=40, **kw)
        return r
    except Exception as e:  # noqa
        print(f"  ERR {url[:110]} -> {type(e).__name__}: {str(e)[:100]}")
        return None


def show(url, n=600, **kw):
    r = get(url, **kw)
    if r is None:
        return None
    print(f"[{r.status_code}] {url[:130]}  ct={r.headers.get('content-type','')[:40]} len={len(r.content)}")
    print("   ", r.text[:n].replace("\n", " ") if 'text' in r.headers.get('content-type','') or 'json' in r.headers.get('content-type','') else "(binary)")
    return r


print("=== DENMARK Energi Data Service")
r = get("https://api.energidataservice.dk/meta/dataset")
if r is not None and r.ok:
    try:
        js = r.json()
        names = [d.get("datasetName") if isinstance(d, dict) else d for d in js] if isinstance(js, list) else list(js)[:5]
        print(len(names), "datasets")
        print([n for n in names if n and re.search(r"capac|master|stam|plant|install|wind|solar|unit|register", str(n), re.I)])
    except Exception as e:
        print("list parse", e, r.text[:300])
else:
    print("  list status", getattr(r, "status_code", None))
for ds in ["CapacityPerMunicipality", "MasterdataWindPower", "MasterdataSolarPower", "PowerPlantMasterData",
           "ProductionCapacityPerMunicipality", "Stamdataregister", "CapacityPerPostalCode"]:
    show(f"https://api.energidataservice.dk/dataset/{ds}?limit=3", 500)
    show(f"https://api.energidataservice.dk/meta/dataset/{ds}", 300)

print("=== SWEDEN")
show("https://pxexternal.energimyndigheten.se/api/v1/sv/", 1500)
show("https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndigheten/", 1500)
show("https://www.energimyndigheten.se/statistik/den-officiella-statistiken/statistikprodukter/vindkraftsstatistik/", 300)
show("https://api.scb.se/OV0104/v1/doris/sv/ssd/START/EN/EN0105", 1500)
show("https://api.scb.se/OV0104/v1/doris/en/ssd/START/EN/EN0105/", 1500)
show("https://vbk.lansstyrelsen.se/", 200)

print("=== NORWAY")
show("https://gis3.nve.no/map/rest/services/?f=json", 1500)
show("https://gis3.nve.no/arcgis/rest/services?f=json", 800)
show("https://api.nve.no/web/Powerplant/GetHydroPowerPlantsInOperation", 400)
show("https://api.nve.no/web/Powerplant/GetWindPowerPlantsInOperation", 400)
show("https://data.ssb.no/api/v0/en/table/", 300)

print("=== FINLAND")
show("https://energiavirasto.fi/en/power-plant-register", 400)
show("https://energiavirasto.fi/voimalaitosrekisteri", 400)
show("https://data.fingrid.fi/api/datasets?pageSize=500", 400)
show("https://api.fingrid.fi/v1/variable/", 200)
show("https://tuulivoimayhdistys.fi/en/wind-power-in-finland/wind-power-statistics", 300)

print("=== ESTONIA")
for p in ["api/system", "api/system/with-plan", "api/power-system", "api/installed-capacity", "api/production-capacity"]:
    show(f"https://dashboard.elering.ee/{p}", 300)
show("https://dashboard.elering.ee/assets/api-doc.html", 300)
show("https://dashboard.elering.ee/api/swagger.json", 300)
show("https://dashboard.elering.ee/v3/api-docs", 300)

print("=== LATVIA")
show("https://www.ast.lv/en/electricity-market-review", 300)
show("https://www.ast.lv/en/content/power-system-data", 300)
show("https://data.gov.lv/dati/api/3/action/package_search?q=installed%20capacity%20electricity&rows=10", 1500)
show("https://data.gov.lv/dati/api/3/action/package_search?q=elektrostacij&rows=10", 1500)

print("=== LITHUANIA")
show("https://energy.litgrid.eu/", 200)
show("https://www.litgrid.eu/index.php/system-services/power-system/installed-generation-capacities/", 300)
show("https://data.gov.lt/api/3/action/package_search?q=elektrin%C4%97&rows=10", 800)
show("https://opendata.gov.lt/api/3/action/package_search?q=installed%20capacity&rows=10", 800)

print("=== ENERGY-CHARTS fallback")
for c in ["dk", "se", "no", "fi", "ee", "lv", "lt"]:
    r = get(f"https://api.energy-charts.info/installed_power?country={c}&time_step=monthly&installation_decommission=false")
    if r is None:
        continue
    print(f"[{r.status_code}] {c} len={len(r.content)}")
    try:
        js = r.json()
        print("   keys", list(js)[:6], "n_time", len(js.get("time", [])), "first/last", (js.get("time") or [None])[0], (js.get("time") or [None])[-1])
        for s in js.get("production_types", []):
            d = s.get("data", [])
            print("     ", s.get("name"), len(d), d[-1:] if d else None)
    except Exception as e:
        print("   parse", e, r.text[:200])
