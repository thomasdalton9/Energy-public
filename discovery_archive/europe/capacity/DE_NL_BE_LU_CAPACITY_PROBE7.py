"""Probe 7 (12 Oct 2026): MaStR members by range read (no full download), Katalogwerte, Elia monitored capacity by date, ILR xlsx, NL leads."""
import io, re, os, sys, json, zipfile, requests, collections
import time
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"
def rget(u, tries=6, **kw):
    for k in range(tries):
        try:
            r = S.get(u, timeout=180, **kw)
            if r.status_code < 500: return r
        except Exception as e:
            print("   retry", k, repr(e)[:100])
        time.sleep(5 * (k + 1))
    raise RuntimeError("giving up " + u)
class RangeFile(io.RawIOBase):
    def __init__(self, url):
        r = rget(url, headers={"Range": "bytes=0-0"})
        print("RANGE0", r.status_code, r.headers.get("content-range"), r.headers.get("last-modified"))
        self.size = int(r.headers["content-range"].split("/")[1]); self.pos = 0; self.url = url
    def seekable(self): return True
    def readable(self): return True
    def tell(self): return self.pos
    def seek(self, off, whence=0):
        self.pos = off if whence == 0 else self.pos + off if whence == 1 else self.size + off
        return self.pos
    def read(self, n=-1):
        if n < 0: n = self.size - self.pos
        n = min(n, self.size - self.pos)
        if n <= 0: return b""
        d = rget(self.url, headers={"Range": f"bytes={self.pos}-{self.pos+n-1}"}).content
        self.pos += len(d); return d
    def readinto(self, b):
        d = self.read(len(b)); b[:len(d)] = d; return len(d)
def dec(b): return b.decode("utf-16", "replace") if b[:2] in (b"\xff\xfe", b"\xfe\xff") else b.decode("utf-8", "replace")
try:
    
def dec(b): return b.decode("utf-16", "replace") if b[:2] in (b"\xff\xfe", b"\xfe\xff") else b.decode("utf-8", "replace")
def sq(s, n=1500): return re.sub(r"\s+", " ", s)[:n]
print("### MaStR")
html = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', html).group(1)
z = zipfile.ZipFile(io.BufferedReader(RangeFile(url), buffer_size=1 << 20))
info = {i.filename: i for i in z.infolist()}
for n in ["EinheitenSolar_1.xml","EinheitenWind.xml","EinheitenStromSpeicher_1.xml","EinheitenVerbrennung.xml"]:
    i = info[n]; print("member", n, i.file_size, i.compress_size, i.compress_type, i.header_offset)
def head(name, n=9000):
    with z.open(info[name]) as f: return dec(f.read(n))
for nm in ["EinheitenWind.xml", "EinheitenVerbrennung.xml", "EinheitenKernkraft.xml", "EinheitenWasser.xml", "EinheitenBiomasse.xml", "EinheitenGasErzeuger.xml", "EinheitenSolar_1.xml", "EinheitenStromSpeicher_1.xml", "GeloeschteUndDeaktivierteEinheiten_1.xml", "Einheitentypen.xml"]:
    if nm in info:
        s = head(nm)
        m = re.search(r"<(\w+)>\s*<(\w+)>.*?</\2>", s[s.find("<", 40):], re.S)
        print("\n--", nm, info[nm].file_size, "\n", sq(s[s.find("?>")+2:], 2600))
with z.open(info["Katalogwerte.xml"]) as f: kw = dec(f.read())
print("\n-- Katalogwerte head:", sq(kw, 900))
ents = re.findall(r"<Katalogwert>(.*?)</Katalogwert>", kw, re.S)
rows = [dict(re.findall(r"<(\w+)>(.*?)</\1>", e, re.S)) for e in ents]
print(len(rows), "entries", rows[:2])
cats = collections.Counter(r.get("KatalogKategorieId") for r in rows)
print("categories", dict(cats))
for cat in ["1", "4"]:
    print("cat", cat, [(r.get("Id"), r.get("Wert")) for r in rows if r.get("KatalogKategorieId") == cat])
import codecs
def timeparse(name, tag):
    t1 = time.time(); n = 0; d_ = codecs.getincrementaldecoder("utf-16")(); carry = ""; nb = 0
    with z.open(info[name]) as f:
        while True:
            b = f.read(1 << 24)
            if not b: break
            nb += len(b); s = carry + d_.decode(b); parts = s.split("</%s>" % tag); carry = parts.pop(); n += len(parts)
    return n, nb, round(time.time() - t1, 1)
print("timeparse Wind", timeparse("EinheitenWind.xml", "EinheitWind"))
print("timeparse Solar_1 (range streamed)", timeparse("EinheitenSolar_1.xml", "EinheitSolar"))
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
