"""Probe 2: GISCO LAU (municipality) population for countries with < 4 Urban Audit cities and the Western Balkans; TurkStat; Wikidata."""
import json, os, requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "degree_days")
UA = {"User-Agent": "Mozilla/5.0"}
def get(url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=240, **kw)
        print(url[:140], r.status_code, len(r.content), flush=True)
        return r
    except Exception as e:
        print(url[:140], "ERR", e, flush=True)
base = "https://gisco-services.ec.europa.eu/distribution/v2/lau/"
r = get(base + "lau-2021.json")
if r is not None and r.ok:
    open(os.path.join(OUT, "lau_index.json"), "wb").write(r.content)
for f in ("geojson/LAU_RG_01M_2021_4326.geojson", "geojson/LAU_LB_2021_4326.geojson", "geojson/LAU_RG_01M_2023_4326.geojson", "geojson/LAU_LB_2023_4326.geojson", "csv/LAU_RG_01M_2021_4326.csv", "csv/LAU_RG_01M_2023_4326.csv"):
    r = get(base + f)
    if r is not None and r.ok:
        name = f.replace("/", "_")
        if len(r.content) < 25_000_000:
            open(os.path.join(OUT, name), "wb").write(r.content)
        else:
            try:
                js = r.json()
                rows = [dict(ft["properties"], lon=ft["geometry"]["coordinates"][0] if ft["geometry"]["type"] == "Point" else None) for ft in js["features"] if ft["properties"].get("CNTR_CODE") in
                        {"SI","CY","LU","MT","EE","XK","RS","BA","ME","MK","AL","TR","UK","IE","DK","LV","NO","IS"}]
                json.dump(rows, open(os.path.join(OUT, name + ".subset.json"), "w"))
                print("subset", len(rows))
            except Exception as e:
                print("big file not parsed", e)
sparql = """SELECT ?c ?cLabel ?pop ?pt ?time ?ref ?country WHERE { VALUES ?country { wd:Q403 wd:Q225 wd:Q236 wd:Q221 wd:Q222 wd:Q1246 }
 ?c wdt:P31/wdt:P279* wd:Q515 ; wdt:P17 ?country ; p:P1082 ?st . ?st ps:P1082 ?pop . OPTIONAL { ?st pq:P585 ?time } OPTIONAL { ?st pq:P459 ?m }
 OPTIONAL { ?c wdt:P625 ?pt } SERVICE wikibase:label { bd:serviceParam wikibase:language "en". } FILTER(?pop > 30000) }"""
r = get("https://query.wikidata.org/sparql", params={"query": sparql, "format": "json"})
if r is not None and r.ok:
    open(os.path.join(OUT, "wikidata_wb.json"), "wb").write(r.content)
for i, u in enumerate(["https://data.tuik.gov.tr/Bulten/Index?p=Adrese-Dayali-Nufus-Kayit-Sistemi-Sonuclari-2024-53783",
          "https://data.tuik.gov.tr/Bulten/Index?p=Address-Based-Population-Registration-System-Results-2024-53783",
          "https://biruni.tuik.gov.tr/medas/?kn=95&locale=tr",
          "https://www.stat.gov.rs/en-US/",
          "https://monstat.org/eng/", "https://www.instat.gov.al/en/", "https://askdata.rks-gov.net/", "https://www.stat.gov.mk/Default_en.aspx",
          "https://popis.gov.rs/", "https://ws.cso.ie/public/api.restful/PxStat.Data.Cube_API.ReadDataset/FY001/CSV/1.0/en"]):
    r = get(u)
    if r is not None and r.ok and r.headers.get("content-type", "").startswith("text"):
        open(os.path.join(OUT, f"probe2_page{i}.html"), "wb").write(r.content[:400000])
