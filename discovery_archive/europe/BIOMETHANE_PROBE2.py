"""Biomethane probe 2: ODRE biomethane dataset schemas + annual sums, CBS ProductieUitAndereBronnen, GNI datasets, THE keys, DE sources. Prints only."""
import io, json, re, requests
import pandas as pd
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
def g(url, **kw):
    kw.setdefault("headers", UA)
    try:
        return requests.get(url, timeout=60, **kw)
    except Exception as e:
        print("ERR", url, type(e).__name__, e); return None

print("=== ODRE")
B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
for ds in ["production-definitive-et-journaliere-de-biomethane-tout-reseau", "odre-prod-grdgrt-operateur-def", "prod-def-reg-jour-biom-reseau-grtgrd",
           "prod-nat-gaz-horaire-def", "igrm-nat", "production-annuelle-de-biomethane-par-site-raccorde-au-reseau-de-transport-et-de", "igra-nat"]:
    r = g(B + ds)
    if r is None or not r.ok: print(ds, r and r.status_code); continue
    j = r.json(); print("##", ds, j.get("metas", {}).get("default", {}).get("modified"), j.get("metas", {}).get("default", {}).get("data_processed"))
    print("  fields:", [(f["name"], f["type"]) for f in j["fields"]])
    r = g(B + ds + "/records", params={"limit": 3, "order_by": [f["name"] for f in j["fields"] if f["type"] in ("date", "datetime")][0] + " desc"} if any(f["type"] in ("date", "datetime") for f in j["fields"]) else {"limit": 3})
    if r is not None: print("  sample:", json.dumps(r.json().get("results"))[:900])

print("=== CBS")
r = g("https://opendata.cbs.nl/ODataApi/odata/86103NED/TypedDataSet", params={"$select": "Perioden,TotaalAanbod_1,WinningUitDeBodem_2,ProductieUitAndereBronnen_3", "$filter": "substringof('JJ',Perioden) or substringof('MM',Perioden)", "$top": 400})
if r is not None and r.ok:
    d = pd.DataFrame(r.json()["value"]); print(d.tail(14).to_string())
    m = d[d.Perioden.str.contains("MM")].copy(); m["y"] = m.Perioden.str[:4]; print(m.groupby("y").ProductieUitAndereBronnen_3.agg(["sum", "count"]).to_string())
    print(d[d.Perioden.str.contains("JJ")].to_string())
r = g("https://opendata.cbs.nl/ODataApi/odata/86103NED/DataProperties"); 
if r is not None:
    for p in r.json()["value"]:
        if p["Key"].startswith("ProductieUitAnd"): print(p.get("Description"))
r = g("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$filter": "substringof('iogas',Title) or substringof('roen gas',Title) or substringof('roengas',Title) or substringof('ernieuwbare energie',Title)", "$select": "Identifier,Title,Period"})
if r is not None: print(re.findall(r"<d:Identifier>(.*?)</d:Identifier><d:Title>(.*?)</d:Title><d:Period>(.*?)</d:Period>", r.text))

print("=== Vertogas / RVO")
for u in ["https://www.vertogas.nl/", "https://www.vertogas.nl/over-vertogas/cijfers", "https://www.vertogas.nl/nieuws"]:
    r = g(u); print(u, r and r.status_code, r and len(r.text))
    if r is not None and r.ok:
        print(" ", set(re.findall(r'href="([^"]*(?:cijfers|groen-gas|statist|xls|csv|data)[^"]*)"', r.text, re.I)))

print("=== GNI")
for ds in ["physicalflows", "gasconsumption", "renewablegas", "biomethane", "renewable-gas", "entryflows", "gasproduction"]:
    r = g(f"https://www.gasnetworks.ie/csv/{ds}", params={"frequency": "daily", "date": "2026-09-01", "date_end": "2026-09-04"}, headers={**UA, "Referer": "https://www.gasnetworks.ie/about/data-transparency/"})
    print(ds, r and r.status_code, r and r.text[:150].replace("\n", " | "))
    if r is not None and r.ok and "Value" in r.text[:100]:
        d = pd.read_csv(io.StringIO(r.text)); print(d.groupby(["Name", "Location"]).Value.sum().to_string())
r = g("https://www.gasnetworks.ie/about/data-transparency/entry-flows/physical-flows")
print("page", r and r.status_code, r and len(r.text))
if r is not None and r.ok:
    print(sorted(set(re.findall(r'(?:href|src|data-[a-z-]+)="([^"]+)"', r.text)))[:80])
for u in ["https://www.gasnetworks.ie/about/data-transparency/", "https://www.gasnetworks.ie/about/data-transparency/entry-flows/"]:
    r = g(u); print(u, r and r.status_code)
    if r is not None and r.ok: print(sorted(set(re.findall(r'href="(/about/data-transparency[^"]*)"', r.text))))
r = g("https://www.gasnetworks.ie/api/v1/physicalflows"); print("api", r and r.status_code, r and r.text[:600])

print("=== THE")
r = g("https://datenservice-api.tradinghub.eu/api/evoq/GetAggregierteVerbrauchsdatenTabelle", params={"DatumStart": "09-01-2026", "DatumEnde": "09-03-2026", "GasXType_Id": "all"}, headers={**UA, "Origin": "https://www.tradinghub.eu", "Referer": "https://www.tradinghub.eu/"})
print(r and r.status_code, r and r.text[:700])

print("=== DE other")
for u in ["https://www.biogaspartner.de/", "https://www.dena.de/themen-projekte/projekte/energiesysteme/biogaspartner/", "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Monitoringberichte/start.html",
          "https://www.destatis.de/DE/Themen/Branchen-Unternehmen/Energie/Erzeugung/Tabellen/gasbilanz.html"]:
    r = g(u); print(u, r and r.status_code, r and len(r.text))
    if r is not None and r.ok: print("  ", sorted(set(re.findall(r'href="([^"]*(?:biomethan|biogas|xlsx|xls|csv)[^"]*)"', r.text, re.I)))[:25])
