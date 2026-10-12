"""Probe 4 (12 Oct 2026): Ember monthly wind/solar capacity dataset, Elering production page, Statistics Estonia / Latvia
PxWeb capacity tables, Finnish wind/solar half-year pages, Norwegian wind register fields, Energinet dataset list, Lithuania links."""
import io, json, re
import requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
def get(url, **kw):
    try: return S.get(url, timeout=60, **kw)
    except Exception as e: print("  ERR", url[:100], type(e).__name__); return None
def text(r): return re.sub(r"\s+", " ", re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S|re.I)) 
def section(name):
    def deco(f):
        print(f"=== {name}")
        try: f()
        except Exception as e: print("SECTION ERROR", type(e).__name__, e)
    return deco

@section("EMBER")
def _():
    for u in ["https://ember-energy.org/data/monthly-wind-and-solar-capacity-data/"]:
        r = get(u); print(u, getattr(r, "status_code", None))
        if r is not None:
            print(sorted(set(re.findall(r'https?://[^"\'\s<>]+\.(?:csv|xlsx|zip|json)', r.text)))[:20])
            print(sorted(set(re.findall(r'https?://[^"\'\s<>]*googleapis[^"\'\s<>]*', r.text)))[:20])
            print(text(r)[:1500])

@section("ELERING page")
def _():
    for u in ["https://elering.ee/en/production-and-forecast", "https://elering.ee/en/node/2304"]:
        r = get(u); print(u, getattr(r, "status_code", None))
        if r is not None:
            t = text(r); i = t.find("apacity"); print(t[max(0, i-300):i+1500])
            print(sorted(set(re.findall(r'href="([^"]*(?:xls|csv|capacity|tootmis|api)[^"]*)"', r.text)))[:20])

@section("EE stat")
def _():
    for u in ["https://andmed.stat.ee/api/v1/en/stat/majandus/energia", "https://andmed.stat.ee/api/v1/en/stat/majandus", "https://andmed.stat.ee/api/v1/en/stat"]:
        r = get(u); print(u, getattr(r, "status_code", None), r.text[:600] if r is not None else "")
    r = get("https://andmed.stat.ee/api/v1/en/stat/majandus/energia/energia-tarbimine-tootmine")
    print(r.status_code if r is not None else None, r.text[:800] if r is not None else "")

@section("LV CSB")
def _():
    for u in ["https://data.stat.gov.lv/api/v1/en/OSP_PUB/START/NOZ/ENE", "https://data.stat.gov.lv/api/v1/en/OSP_PUB/START/NOZ", "https://data.stat.gov.lv/api/v1/en/OSP_PUB"]:
        r = get(u); print(u, getattr(r, "status_code", None), r.text[:1500] if r is not None else "")

@section("LT stat")
def _():
    for u in ["https://osp.stat.gov.lt/PXWeb/api/v1/en/Ekonomines_gamybos_apimtys/", "https://osp.stat.gov.lt/PXWeb/api/v1/en/"]:
        r = get(u); print(u, getattr(r, "status_code", None), r.text[:500] if r is not None else "")
    r = get("https://www.litgrid.eu/index.php/system/open-data/open-data-sets/")
    print(getattr(r, "status_code", None))
    r = get("https://www.litgrid.eu/")
    hrefs = sorted(set(re.findall(r'href="(https?://(?:www\.)?litgrid\.eu[^"]+)"', r.text)))
    print(len(hrefs)); print([h for h in hrefs if re.search(r"duomen|data|sistem|system|stat", h, re.I)][:40])

@section("FI pages")
def _():
    for u in ["https://suomenuusiutuvat.fi/en/wind-power-statistics-06-2026/", "https://suomenuusiutuvat.fi/en/solar-power-statistics-06-2026/"]:
        r = get(u); print(u, getattr(r, "status_code", None))
        if r is not None:
            print(sorted(set(re.findall(r'https://suomenuusiutuvat.fi/media/[^"\'\s<>]+', r.text)))[:10]); print(text(r)[1500:3200])
    r = get("https://suomenuusiutuvat.fi/en/category/statistics/")
    print(sorted(set(re.findall(r'href="(https://suomenuusiutuvat.fi/en/(?:wind|solar)[^"]*)"', r.text)))[:30])

@section("NO wind fields")
def _():
    import collections
    j = get("https://api.nve.no/web/WindPowerplant/GetWindPowerplants").json()
    print(len(j)); print(json.dumps(j[0], ensure_ascii=False)[:1200]); print(sorted({k for x in j for k in x}))
    print(collections.Counter(str(x.get("Status") or x.get("Driftsstatus"))[:30] for x in j).most_common(5))
    jo = get("https://api.nve.no/web/WindPowerplant/GetWindPowerplantsInOperation").json()
    print("in operation", len(jo), sum(x.get("InstallertEffekt_MW") or 0 for x in jo))
    print(collections.Counter(str(x.get("IdriftsettelseForsteByggetrinn"))[:4] for x in jo).most_common(30))
    r = get("https://www.nve.no/energi/energisystem/solkraft/"); print("solkraft", getattr(r, "status_code", None))
    if r is not None: print(sorted(set(re.findall(r'href="([^"]*(?:sol|xls|csv)[^"]*)"', r.text)))[:20])
    for u in ["https://api.nve.no/doc/vindkraftdatabase/"]:
        r = get(u); print(text(r)[:1500])

@section("DK dataset names")
def _():
    r = get("https://api.energidataservice.dk/meta/dataset")
    print([ (d.get("datasetName") if isinstance(d, dict) else d) for d in r.json()])
    r = get("https://api.energidataservice.dk/dataset/CapacityPerMunicipality?start=2026-09-01&end=2026-10-01&limit=2000")
    j = r.json(); import pandas as pd
    d = pd.DataFrame(j["records"]); print(d.shape); print(d[["OffshoreWindCapacity", "OnshoreWindCapacity", "SolarPowerCapacity", "CapacityGe100MW", "CapacityLt100MW"]].sum())
    r = get("https://api.energidataservice.dk/dataset/CapacityPerMunicipality?start=2016-01-01&end=2016-02-01&limit=2000"); d = pd.DataFrame(r.json()["records"]); print(d[["OffshoreWindCapacity", "OnshoreWindCapacity", "SolarPowerCapacity", "CapacityGe100MW", "CapacityLt100MW"]].sum())
