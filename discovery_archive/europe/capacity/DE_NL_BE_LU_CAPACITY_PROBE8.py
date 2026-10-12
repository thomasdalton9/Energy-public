"""Probe 8 (12 Oct 2026): MaStR Katalog ids used by the unit files, Elia monthly availability, CBS 86266NED/82610NED MW by carrier,
Energy-Charts Germany monthly (FALLBACK reference) at 1 Jan."""
import io, re, time, json, zipfile, requests, collections
exec(open("discovery_archive/europe/capacity/DE_NL_BE_LU_CAPACITY_PROBE7.py").read().split('print("### MaStR")')[0])
html = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', html).group(1)
z = zipfile.ZipFile(io.BufferedReader(RangeFile(url), buffer_size=1 << 20))
info = {i.filename: i for i in z.infolist()}
with z.open(info["Katalogwerte.xml"]) as f: kw = dec(f.read())
with z.open(info["Katalogkategorien.xml"]) as f: kk = dec(f.read())
cats = {m[0]: m[1] for m in re.findall(r"<Id>(\d+)</Id><Name>(.*?)</Name>", kk)}
rows = [dict(re.findall(r"<(\w+)>(.*?)</\1>", e, re.S)) for e in re.findall(r"<Katalogwert>(.*?)</Katalogwert>", kw, re.S)]
want = set(range(2400, 2520)) | set(range(500, 560)) | set(range(680, 735)) | set(range(2440, 2480))
print("### Katalog (id: value [category])")
print("; ".join(f"{r['Id']}:{r['Wert']}[{cats.get(r['KatalogKategorieId'],'?')}]" for r in rows if int(r["Id"]) in want))
print("category names:", {k: v for k, v in cats.items() if any(x in v for x in ["Energie", "Techn", "Brennstoff", "Speicher", "Betrieb", "Wasser", "Biomasse", "Windkraft", "Solar"])})
print("### Elia")
B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/"
def agg(ds, **p):
    r = S.get(B + ds + "/records", params=p, timeout=120)
    try: return r.json()["results"]
    except Exception: return r.text[:300]
for ds in ["ods031", "ods032"]:
    res = agg(ds, select="date_format(datetime,'YYYY-MM') as m, count(*) as n", group_by="date_format(datetime,'YYYY-MM')", order_by="m", limit=100, where="datetime>='2020-09-01'")
    print(ds, "months with data:", len(res), [(r["m"], r["n"]) for r in res][:3], "...", [(r["m"], r["n"]) for r in res][-3:] if isinstance(res, list) else res)
    if isinstance(res, list):
        have = {r["m"] for r in res}; allm = [f"{y}-{m:02d}" for y in range(2020, 2027) for m in range(1, 13) if f"{y}-{m:02d}" >= "2020-09" and f"{y}-{m:02d}" <= "2026-10"]
        print(ds, "missing months:", [m for m in allm if m not in have])
for d in ["2024-05-31", "2024-06-30", "2025-12-31", "2026-09-30"]:
    res = agg("ods179", select="fueltypepublication, count(*) as n, sum(technicalpmax) as s", where=f"date>='{d}T00:00:00Z' and date<'{d}T23:59:59Z'", group_by="fueltypepublication", limit=20)
    print("ods179", d, res if isinstance(res, str) else [(r["fueltypepublication"], r["n"], r["s"]) for r in res])
res = agg("ods179", select="date, count(*) as n", where="date>='2024-05-20T00:00:00Z' and date<'2024-06-02T00:00:00Z'", group_by="date", order_by="date", limit=30)
print("ods179 first dates", res if isinstance(res, str) else [(r["date"], r["n"]) for r in res])
res = agg("ods036", select="date, fuelcode, sum(productionunitnominalpower) as s", group_by="date, fuelcode", order_by="date", limit=100)
print("ods036", res if isinstance(res, str) else [(r["date"][:10], r["fuelcode"], round(r["s"])) for r in res if r["date"] >= "2021"])
print("### CBS")
for tb, dim in [("86266NED", "Energiedragers"), ("82610NED", "BronTechniek"), ("86266NED", "CentraleDecentraleProductie")]:
    r = S.get(f"https://opendata.cbs.nl/ODataApi/odata/{tb}/{dim}", params={"$format": "json"}, timeout=60)
    print(tb, dim, [(x["Key"].strip(), x["Title"]) for x in r.json()["value"]])
r = S.get("https://opendata.cbs.nl/ODataApi/odata/86266NED/TypedDataSet", params={"$format": "json", "$filter": "Perioden eq '2025JJ00' or Perioden eq '2021JJ00'", "$select": "CentraleDecentraleProductie,Energiedragers,Perioden,ElektrischVermogenMW_7", "$top": 400}, timeout=60)
print("86266NED rows", len(r.json()["value"]) if r.ok else r.text[:200])
for x in r.json()["value"]: print("   ", x["CentraleDecentraleProductie"].strip(), x["Energiedragers"].strip(), x["Perioden"], x["ElektrischVermogenMW_7"])
r = S.get("https://opendata.cbs.nl/ODataApi/odata/82610NED/TypedDataSet", params={"$format": "json", "$filter": "Perioden eq '2025JJ00' or Perioden eq '2021JJ00'", "$select": "BronTechniek,Perioden,OpgesteldElektrischVermogenEindeJaar_8", "$top": 400}, timeout=60)
for x in r.json()["value"]: print("   82610", x["BronTechniek"].strip(), x["Perioden"], x["OpgesteldElektrischVermogenEindeJaar_8"])
print("### Energy-Charts DE monthly")
d = S.get("https://api.energy-charts.info/installed_power?country=de&time_step=monthly&installation_decommission=false", timeout=60).json()
idx = {t: i for i, t in enumerate(d["time"])}
for t in ["01.2021", "01.2022", "01.2023", "01.2024", "01.2025", "01.2026", "10.2026"]:
    print(t, {s["name"]: s["data"][idx[t]] for s in d["production_types"]})
