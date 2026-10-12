"""Probe 4 (12 Oct 2026): MaStR unit file names + one decoded record per unit type + Katalog values (UTF-16 XML, read by HTTP range),
Elia monthly aggregates of monitoredcapacity (ods031/ods032) and ods179 units by fuel, CBS OData, Luxembourg ILR dataset resources."""
import io, re, json, zipfile, requests, collections
exec(open("discovery_archive/europe/capacity/DE_NL_BE_LU_CAPACITY_PROBE3.py").read().split('print("### MaStR")')[0])
def dec(b): return b.decode("utf-16", "replace") if b[:2] in (b"\xff\xfe", b"\xfe\xff") else b.decode("utf-8", "replace")
print("### MaStR")
t = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', t).group(1)
z = zipfile.ZipFile(io.BufferedReader(RangeFile(url), buffer_size=1 << 20))
g = collections.defaultdict(lambda: [0, 0])
for i in z.infolist():
    n = re.sub(r"_\d+\.xml$", ".xml", i.filename); g[n][0] += 1; g[n][1] += i.file_size
for n, v in sorted(g.items()): print("  ", n, v)
want = ["EinheitenSolar", "EinheitenWind", "EinheitenGasErzeuger", "EinheitenStromSpeicher", "EinheitenBiomasse", "EinheitenVerbrennung", "EinheitenKernkraft", "EinheitenWasser", "EinheitenGeothermieGrubengasDruckentspannung"]
names = [i for i in z.infolist()]
for w in want:
    for i in names:
        if i.filename.startswith(w) and i.filename.endswith(".xml"):
            with z.open(i) as f:
                s = dec(f.read(6000))
            m = re.search(r"<Einheit\w+>.*?</Einheit\w+>", s, re.S)
            print("\n--", i.filename, i.file_size, "\n", (m.group(0) if m else s)[:2600]); break
for i in names:
    if "atalog" in i.filename:
        print("\n--", i.filename, i.file_size)
        with z.open(i) as f: s = dec(f.read())
        for k in ["Betriebsstatus", "Energietr", "Einheittyp", "Technologie"]:
            for m in re.finditer(r"<Katalogwert>(.*?)</Katalogwert>", s, re.S):
                pass
        ents = re.findall(r"<Id>(\d+)</Id>\s*<KatalogKategorieId>(\d+)</KatalogKategorieId>\s*<Wert>(.*?)</Wert>", s, re.S)
        print(len(ents), "entries; sample:", s[:800])
        wanted = {"35", "31", "37", "38", "2495", "2497", "2498", "2493", "2494", "2496", "2403", "2404", "2405", "2406", "2407", "2408", "2409", "2410"}
        print([e for e in ents if e[0] in wanted])
        break
print("### Elia")
B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/"
def q(ds, **p):
    r = S.get(B + ds + "/records", params=p, timeout=120); print("  ", ds, r.status_code, r.text[:1800]); return r
q("ods031", select="date_format(datetime,'YYYY-MM') as m, offshoreonshore, max(monitoredcapacity) as mx, min(monitoredcapacity) as mn", where="region='Belgium' and datetime>='2025-06-01'", group_by="date_format(datetime,'YYYY-MM'), offshoreonshore", order_by="m desc", limit=12)
q("ods031", select="region, gridconnectiontype, offshoreonshore, count(*) as n, max(monitoredcapacity) as mx", where="datetime>='2026-09-01'", group_by="region, gridconnectiontype, offshoreonshore", limit=40)
q("ods032", select="date_format(datetime,'YYYY-MM') as m, max(monitoredcapacity) as mx, min(monitoredcapacity) as mn", where="region='Belgium'", group_by="date_format(datetime,'YYYY-MM')", order_by="m asc", limit=3)
q("ods032", select="min(datetime) as a, max(datetime) as b, count(*) as n", limit=1)
q("ods031", select="min(datetime) as a, max(datetime) as b, count(*) as n", limit=1)
q("ods179", select="min(date) as a, max(date) as b, count(*) as n", limit=1)
q("ods179", select="fueltypepublication, count(*) as n, max(technicalpmax) as mx", group_by="fueltypepublication", limit=40)
q("ods179", select="date_format(date,'YYYY-MM') as m, fueltypepublication, count(distinct technicalunit) as u, sum(technicalpmax) as s", where="date>='2026-09-01'", group_by="date_format(date,'YYYY-MM'), fueltypepublication", limit=40)
q("ods036", select="fuelcode, count(*) as n, sum(productionunitnominalpower) as s, min(date) as a, max(date) as b", group_by="fuelcode", limit=40)
print("### CBS")
for u in ["https://opendata.cbs.nl/ODataApi/odata/85005NED/DataProperties", "https://opendata.cbs.nl/ODataApi/odata/85005NED/Perioden", "https://opendata.cbs.nl/ODataApi/odata/85005NED/TypedDataSet?$top=3",
          "https://opendata.cbs.nl/ODataApi/odata/82610NED/DataProperties", "https://opendata.cbs.nl/ODataApi/odata/82610NED/TypedDataSet?$top=2",
          "https://opendata.cbs.nl/ODataApi/odata/86266NED/DataProperties"]:
    r = S.get(u, params={"$format": "json"}, timeout=90)
    print(u, r.status_code, len(r.text))
    try:
        v = r.json()["value"]
        if "DataProperties" in u:
            for x in v: print("    ", x.get("Key"), "|", (x.get("Title") or "")[:60], "|", x.get("Unit"))
        elif "Perioden" in u: print("   ", [x["Key"] for x in v][-16:])
        else: print("   ", json.dumps(v)[:900])
    except Exception as e: print("   ", repr(e), r.text[:200])
print("### LU ILR")
r = S.get("https://data.public.lu/api/1/datasets/", params={"q": "ILR électricité", "page_size": 5}, timeout=60)
for d in r.json()["data"]:
    if "lectricit" in d["title"]:
        print(d["title"], d["page"])
        for rs in d["resources"]: print("   ", rs["title"][:70], rs["format"], rs["url"], rs.get("last_modified"))
