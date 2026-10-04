"""Biomethane probe 1: ODRE catalog search, Energinet datasets, CBS 86103NED, Eurostat siec, GNI. Prints only."""
import json, re, requests
UA = {"User-Agent": "Mozilla/5.0 Chrome/124.0"}
def g(url, **kw):
    try:
        r = requests.get(url, timeout=60, headers=UA, **kw)
        return r
    except Exception as e:
        print("ERR", url, type(e).__name__, e); return None

print("=== ODRE")
for q in ["biométhane", "biomethane", "injection", "gaz renouvelable", "biogaz"]:
    r = g("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets", params={"where": f'search("{q}")', "limit": 40, "select": "dataset_id,title,records_count"})
    if r is None: continue
    print(q, r.status_code)
    try:
        for d in r.json().get("results", []): print("  ", d["dataset_id"], "|", d.get("title"), "|", d.get("records_count"))
    except Exception as e: print(r.text[:300])

print("=== Energinet")
r = g("https://api.energidataservice.dk/meta/dataset")
if r is not None:
    print(r.status_code, r.text[:200])
    try:
        for d in r.json():
            s = json.dumps(d)
            if re.search(r"gas|bio", s, re.I): print("  ", d.get("datasetName") or d.get("name") or s[:150])
    except Exception as e: print(e)
r = g("https://api.energidataservice.dk/dataset/Gasflow", params={"start": "2025-01-01", "end": "2025-01-03", "limit": 3})
if r is not None: print(r.text[:800])

print("=== CBS")
for t in ["86103NED", "00377", "83140NED", "82610NED", "85007NED"]:
    r = g(f"https://opendata.cbs.nl/ODataApi/odata/{t}/DataProperties")
    if r is None or not r.ok: print(t, r and r.status_code); continue
    print(t)
    for p in r.json()["value"]:
        if p.get("Type") in ("Dimension","TimeDimension","TopicGroup","Topic") : print("  ", p.get("Key"), p.get("Type"), p.get("Title"), "|", p.get("Unit"))
r = g("https://opendata.cbs.nl/ODataApi/odata/86103NED/TypedDataSet", params={"$top": 3}); print(r.text[:600] if r is not None else "")
r = g("https://opendata.cbs.nl/ODataApi/odata/86103NED/TableInfos"); print(r.text[:800] if r is not None else "")
r = g("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$filter": "substringof('biogas',Title) or substringof('groen gas',Title) or substringof('Aardgasbalans',Title)", "$select": "Identifier,Title,Period,Updated"})
print(r.text[:2000] if r is not None else "")

print("=== Eurostat")
r = g("https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/dataflow/ESTAT/nrg_cb_gasm/latest")
r = g("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm", params={"geo": "FR", "sinceTimePeriod": "2025-01", "format": "JSON", "lang": "EN"})
if r is not None and r.ok:
    j = r.json(); print(j["dimension"]["siec"]["category"]["label"]); print(j["dimension"]["nrg_bal"]["category"]["label"].keys().__len__())
    print({k: v for k, v in j["dimension"]["nrg_bal"]["category"]["label"].items() if re.search("bio|inj|prod|indig", v, re.I)})
else: print(r and r.status_code, r and r.text[:200])
for ds in ["nrg_cb_gas", "nrg_bal_s", "nrg_cb_pem"]:
    pass

print("=== GNI")
for u in ["https://www.gasnetworks.ie/corporate/gas-regulation/transparency/", "https://www.gasnetworks.ie/business/renewable-gas/", "https://www.gasnetworks.ie/corporate/company/our-network/renewable-gas/"]:
    r = g(u); print(u, r and r.status_code, r and len(r.text))
    if r is not None and r.ok:
        for m in set(re.findall(r'href="([^"]+)"', r.text)):
            if re.search(r"biomethane|renewable|cush|mitchelstown|\.xls|\.csv|transparen", m, re.I): print("   ", m)
