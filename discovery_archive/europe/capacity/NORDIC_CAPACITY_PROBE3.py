"""Probe 3 (12 Oct 2026): variables of Swedish PxWeb tables, Norwegian wind register, Finnish wind plant xlsx, Fingrid dataset page,
Latvian AST xlsx, Lithuanian Litgrid open-data links."""
import io, json, re
import requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
def get(url, **kw):
    try: return S.get(url, timeout=60, **kw)
    except Exception as e: print("  ERR", url[:100], type(e).__name__); return None
def section(name):
    def deco(f):
        print(f"=== {name}")
        try: f()
        except Exception as e: print("SECTION ERROR", type(e).__name__, e)
    return deco

@section("SE tables")
def _():
    base = "https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndighetens_statistikdatabas/"
    for t in ["Officiell_energistatistik/Vindkraftsstatistik/EN0105_5.px", "Officiell_energistatistik/Vindkraftsstatistik/EN0105_1.px",
              "Officiell_energistatistik/Natanslutna_solcellsanlaggningar/EN0123_1.px", "Energiindikatorer/4/4.1/EN_IND4-1A_Cont.px",
              "Energiindikatorer/3/3.2/EN_IND3-2A.px"]:
        r = get(base + t)
        if r is None or not r.ok: print(t, getattr(r, "status_code", None)); continue
        j = r.json(); print(t, j.get("title"))
        for v in j["variables"]:
            vals = v["valueTexts"]; print("   ", v["code"], v["text"], len(vals), vals[:3], "...", vals[-3:])
for _t in []: pass

@section("SE data sample")
def _():
    base = "https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndighetens_statistikdatabas/"
    for t in ["Officiell_energistatistik/Vindkraftsstatistik/EN0105_5.px", "Officiell_energistatistik/Natanslutna_solcellsanlaggningar/EN0123_1.px", "Energiindikatorer/4/4.1/EN_IND4-1A_Cont.px"]:
        meta = get(base + t).json()
        q = {"query": [{"code": v["code"], "selection": {"filter": "item", "values": v["values"][-3:]}} for v in meta["variables"] if v["code"].lower().startswith(("tid", "år", "ar", "period"))] +
                      [{"code": v["code"], "selection": {"filter": "all", "values": ["*"]}} for v in meta["variables"] if not v["code"].lower().startswith(("tid", "år", "ar", "period"))],
             "response": {"format": "csv"}}
        r = S.post(base + t, json=q, timeout=60); print(t, r.status_code); print(r.text[:1500])

@section("NO wind + hydro")
def _():
    r = get("https://api.nve.no/doc/")
    for m in sorted(set(re.findall(r'href="([^"]+)"', r.text)))[:60]: print("   ", m)
    for c in ["WindPowerplant", "Windpowerplant", "Vindkraft", "WindPower", "Powerplant"]:
        for u in ["GetWindPowerplantsInOperation", "GetWindPowerPlantsInOperation", "GetWindPowerplants", "GetAll", "GetWindPowerPlants"]:
            r = get(f"https://api.nve.no/web/{c}/{u}")
            if r is not None and r.status_code != 404: print(c, u, r.status_code, r.text[:200])
    j = get("https://api.nve.no/web/Powerplant/GetHydroPowerPlants").json()
    import collections
    print(len(j), collections.Counter(x.get("ErIDrift") for x in j), sorted({k for x in j for k in x})[:60])
    jj = get("https://api.nve.no/web/Powerplant/GetHydroPowerPlantsInOperation").json()
    tot = sum((x.get("MaksYtelse") or 0) for x in jj); print("in operation total MW", tot, "n", len(jj))
    print(collections.Counter(str(x.get("IDriftDato"))[:4] for x in jj).most_common(8))
    print(sum(1 for x in jj if not x.get("IDriftDato")), "without IDriftDato")
    print([ (x["Navn"], x.get("MaksYtelse"), x.get("IDriftDato"), x.get("VannKVType")) for x in jj if x.get("VannKVType") != "Kraftverk"][:6])
    for u in ["https://www.nve.no/energi/energisystem/vindkraft/", "https://www.nve.no/energi/energisystem/vindkraft/vindkraftverk-i-drift/", "https://www.nve.no/energi/energisystem/vindkraft/vindkraftverk-i-norge/"]:
        r = get(u); 
        if r is not None:
            print(u, r.status_code); print(sorted(set(re.findall(r'href="([^"]+(?:xls|csv|api|rest|vindkraft)[^"]*)"', r.text)))[:25])

@section("NO ArcGIS")
def _():
    for u in ["https://gis3.nve.no/map/rest/services/Mapservices/Vindkraft/MapServer?f=json", "https://gis3.nve.no/map/rest/services/Mapservices?f=json"]:
        r = get(u)
        if r is not None: print(u, r.status_code, re.findall(r'"name":"([^"]*(?:ind|Vann|Kraft|kraft|Sol)[^"]*)"', r.text)[:20])

@section("FI")
def _():
    r = get("https://pxdata.stat.fi/PxWeb/api/v1/en/StatFin/ehk/"); 
    for it in r.json(): 
        if re.search(r"capacit|power plant|wind|install", it["text"], re.I): print("  ", it["id"], it["text"])
    for u in ["https://suomenuusiutuvat.fi/media/julkinen-tuulivoima_vuositilastot_6_2025.xlsx", "https://suomenuusiutuvat.fi/media/tuulivoima-rakenteilla_2025_2027_28.8.2025.xlsx"]:
        r = get(u); print(u, getattr(r, "status_code", None), len(r.content) if r is not None else "")
        if r is not None and r.ok:
            import pandas as pd
            x = pd.ExcelFile(io.BytesIO(r.content)); print(x.sheet_names)
            for sh in x.sheet_names[:6]:
                d = x.parse(sh, header=None); print(sh, d.shape); print(d.head(8).to_string()[:1800]); print(d.tail(3).to_string()[:600])
    r = get("https://suomenuusiutuvat.fi/en/category/statistics/")
    print(sorted(set(re.findall(r'href="(https://suomenuusiutuvat.fi/en/[^"]*statist[^"]*)"', r.text)))[:20])
    r = get("https://data.fingrid.fi/en/datasets/268"); t = re.sub(r"<[^>]+>", " ", r.text); print(re.sub(r"\s+", " ", t)[:800])
    r = get("https://data.fingrid.fi/en/datasets/"); 
    names = re.findall(r'datasets/(\d+)"[^>]*>([^<]{3,90})<', r.text)
    for i, n in names:
        if re.search(r"capacit|install", n, re.I): print("  fingrid", i, n)

@section("LV")
def _():
    r = get("https://www.ast.lv/sites/default/files/editor/tirgus%20apskats/TA%202026/Market%20overview%20-%20for%20download.xlsx")
    import pandas as pd
    x = pd.ExcelFile(io.BytesIO(r.content)); print(x.sheet_names)
    for sh in x.sheet_names[:12]:
        d = x.parse(sh, header=None); print(sh, d.shape); print(d.head(4).to_string()[:500])
    for u in ["https://www.ast.lv/en", "https://www.ast.lv/en/content/electricity-generation", "https://www.ast.lv/en/electricity-generation"]:
        r = get(u)
        if r is not None:
            print(u, r.status_code); 
            print(sorted(set(re.findall(r'href="(/en/[^"]*(?:generat|capacity|power-plant|statistic|data)[^"]*)"', r.text)))[:30])

@section("LT")
def _():
    r = get("https://www.litgrid.eu/")
    print(sorted(set(re.findall(r'href="(https://www.litgrid.eu/[^"]*(?:duomen|data|statist|galia|gamyb|gener)[^"]*)"', r.text)))[:50])
    for u in ["https://www.litgrid.eu/index.php/sistema/atviri-duomenys/atviru-duomenu-rinkiniai/3041", "https://www.litgrid.eu/index.php/sistema/atviri-duomenys/3040"]:
        r = get(u); print(u, getattr(r, "status_code", None))
    for u in ["https://www.litgrid.eu/index.php/sistema/elektros-energetikos-sistema/gamyba/3213", "https://www.ena.lt/", "https://www.ena.lt/en/", "https://www.eso.lt/en/about-us/statistics"]:
        r = get(u); print(u, getattr(r, "status_code", None), len(r.text) if r is not None else "")
    r = get("https://www.litgrid.eu/index.php/sistema/atviri-duomenys/atviru-duomenu-rinkiniai/")
    if r is not None: print(r.status_code, sorted(set(re.findall(r'href="([^"]*(?:rinkin|duomen|csv|xls|json|api)[^"]*)"', r.text)))[:40])
