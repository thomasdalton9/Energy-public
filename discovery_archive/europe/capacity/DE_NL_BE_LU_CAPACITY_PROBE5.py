"""Probe 5 (12 Oct 2026): download the MaStR Gesamtdatenexport (3.2 GB, timed), dump compact records/Katalog values, time a parse,
Elia monitored capacity for specific dates, Luxembourg ILR xlsx, a few Dutch leads."""
import io, re, os, sys, time, json, zipfile, subprocess, requests, collections
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"
def dec(b): return b.decode("utf-16", "replace") if b[:2] in (b"\xff\xfe", b"\xfe\xff") else b.decode("utf-8", "replace")
def sq(s, n=1500): return re.sub(r"\s+", " ", s)[:n]
print("### MaStR download")
t0 = time.time()
html = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', html).group(1)
os.makedirs("/tmp/mastr", exist_ok=True); zp = "/tmp/mastr/export.zip"
for k in range(5):
    r = subprocess.run(["curl", "-sS", "-L", "--retry", "5", "--retry-delay", "5", "--retry-all-errors", "-C", "-", "-o", zp, url])
    if r.returncode == 0: break
print("downloaded", os.path.getsize(zp), "bytes in", round(time.time() - t0), "s", url)
z = zipfile.ZipFile(zp)
info = {i.filename: i for i in z.infolist()}
def head(name, n=7000):
    with z.open(info[name]) as f: return dec(f.read(n))
for nm in ["EinheitenWind.xml", "EinheitenVerbrennung.xml", "EinheitenKernkraft.xml", "EinheitenWasser.xml", "EinheitenBiomasse.xml", "EinheitenGasErzeuger.xml", "EinheitenSolar_1.xml", "EinheitenStromSpeicher_1.xml", "GeloeschteUndDeaktivierteEinheiten_1.xml", "Einheitentypen.xml"]:
    if nm in info:
        s = head(nm)
        m = re.search(r"<(Einheit\w+|Geloescht\w+|Deaktiv\w+)>\s*<(Einheit|Geloescht|Deaktiv)\w+>.*?</\2\w+>", s, re.S) or re.search(r"<(\w+)>\s*<(\w+)>.*?</\2>", s, re.S)
        print("\n--", nm, info[nm].file_size, "\n", sq(m.group(0) if m else s, 3200))
s = dec(open(zp, "rb").read(0) or b"") 
with z.open(info["Katalogwerte.xml"]) as f: kw = dec(f.read())
print("\n-- Katalogwerte head:", sq(kw, 700))
ents = re.findall(r"<Katalogwert>(.*?)</Katalogwert>", kw, re.S)
print(len(ents), "entries")
rows = []
for e in ents:
    g = dict(re.findall(r"<(\w+)>(.*?)</\1>", e, re.S)); rows.append(g)
cats = collections.Counter(r.get("KatalogKategorieId") for r in rows)
print("categories", dict(cats))
for cat in ["1", "4"]:
    print("cat", cat, [(r.get("Id"), r.get("Wert")) for r in rows if r.get("KatalogKategorieId") == cat])
# parse timing: EinheitenWind and EinheitenSolar_1 with simple regex
def parse(name, tag, fields):
    t1 = time.time(); n = 0; tot = collections.Counter(); dec_ = __import__("codecs").getincrementaldecoder("utf-16")()
    pat = re.compile(r"<(%s)>([^<]*)<" % "|".join(fields)); carry = ""
    with z.open(info[name]) as f:
        while True:
            b = f.read(1 << 26)
            if not b: break
            s = carry + dec_.decode(b); parts = s.split("</%s>" % tag); carry = parts.pop()
            for p in parts:
                n += 1
    return n, round(time.time() - t1, 1)
print("parse test Wind:", parse("EinheitenWind.xml", "EinheitWind", ["Inbetriebnahmedatum"]))
print("parse test Solar_1:", parse("EinheitenSolar_1.xml", "EinheitSolar", ["Inbetriebnahmedatum"]))
print("### Elia")
B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/"
for ds, extra in [("ods031", "region, gridconnectiontype, offshoreonshore"), ("ods032", "region")]:
    for d in ["2021-01-01", "2023-01-01", "2025-01-01", "2026-01-01", "2026-10-01"]:
        r = S.get(B + ds + "/records", params={"select": extra + ", max(monitoredcapacity) as mx", "where": f"datetime>='{d}T00:00:00Z' and datetime<'{d}T06:00:00Z'", "group_by": extra, "limit": 20}, timeout=120)
        try: print(ds, d, [(tuple(v for k, v in x.items() if k != "mx"), x["mx"]) for x in r.json()["results"]])
        except Exception: print(ds, d, r.status_code, r.text[:200])
print("### Luxembourg ILR xlsx")
try:
    u = "https://download.data.public.lu/resources/donnees-statistiques-du-secteur-delectricite-ilr/20260901-135330/donnees-statistiques-du-secteur-delectricite.xlsx"
    b = S.get(u, timeout=120).content
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows_ = list(ws.iter_rows(values_only=True))
        print("SHEET", ws.title, len(rows_))
        for rw in rows_[:14]: print("   ", [c for c in rw if c is not None][:14])
except Exception as e: print("ILR ERR", repr(e))
for q in ["energie renouvelable", "installations de production", "Creos production", "énergie solaire", "énergie éolienne", "Production d'électricité"]:
    r = S.get("https://data.public.lu/api/1/datasets/", params={"q": q, "page_size": 8}, timeout=60)
    print(q, [(d["title"][:60], d.get("frequency")) for d in r.json().get("data", [])])
print("### NL leads")
for u in ["https://www.netbeheernederland.nl/", "https://energieopwek.nl/", "https://api.energieopwek.nl/", "https://www.certiq.nl/", "https://www.tennet.org/english/operational_management/export_data.aspx",
          "https://opendata.cbs.nl/ODataApi/odata/85005NED/Perioden", "https://www.rvo.nl/onderwerpen/duurzaam-ondernemen/energie-en-milieu-innovaties/zon-pv", "https://klimaatmonitor.databank.nl/"]:
    try:
        r = S.get(u, timeout=40); print(u, r.status_code, len(r.text), sq(re.sub(r"<[^>]+>", " ", r.text), 150))
    except Exception as e: print(u, "ERR", repr(e)[:100])
