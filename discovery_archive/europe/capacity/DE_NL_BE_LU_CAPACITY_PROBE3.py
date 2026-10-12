"""Probe 3 (12 Oct 2026): MaStR Gesamtdatenexport zip structure by HTTP range reads (no full download), Elia installed/monitored
capacity datasets, CBS 85005NED/82610NED structure, Luxembourg open-data search."""
import io, re, json, zipfile, requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"

class RangeFile(io.RawIOBase):
    def __init__(self, url):
        self.url = url
        r = S.head(url, allow_redirects=True, timeout=60)
        print("HEAD", r.status_code, r.headers.get("content-length"), r.headers.get("last-modified"), r.headers.get("accept-ranges"))
        self.size = int(r.headers["content-length"]); self.pos = 0; self.url = r.url
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
        r = S.get(self.url, headers={"Range": f"bytes={self.pos}-{self.pos+n-1}"}, timeout=120)
        d = r.content; self.pos += len(d); return d
    def readinto(self, b):
        d = self.read(len(b)); b[:len(d)] = d; return len(d)

print("### MaStR")
t = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', t).group(1)
print(url)
try:
    z = zipfile.ZipFile(io.BufferedReader(RangeFile(url), buffer_size=1 << 20))
    names = z.infolist()
    print(len(names), "members")
    for i in names:
        print("  ", i.filename, i.file_size, i.compress_size)
    for pat in ["EinheitenSolar_1", "EinheitenWind_1", "EinheitenGasErzeuger", "EinheitenStromSpeicher_1", "EinheitenBiomasse_1", "EinheitenVerbrennung_1", "EinheitenKernkraft", "EinheitenWasser_1", "AnlagenEegSolar_1"]:
        for i in names:
            if pat in i.filename:
                with z.open(i) as f:
                    print("\n--", i.filename, "\n", f.read(2500).decode("utf-8", "replace"))
                break
except Exception as e:
    print("MaStR zip ERR", repr(e))

print("### Elia")
for ds in ["ods031", "ods032", "ods087", "ods036", "ods179", "ods201", "ods086"]:
    try:
        m = S.get(f"https://opendata.elia.be/api/explore/v2.1/catalog/datasets/{ds}", timeout=60).json()
        fs = [(f["name"], f["type"]) for f in m.get("fields", [])]
        print(ds, m["metas"]["default"].get("title"), "| records", m.get("metas", {}).get("default", {}).get("records_count"), "|", fs)
        r = S.get(f"https://opendata.elia.be/api/explore/v2.1/catalog/datasets/{ds}/records?limit=2&order_by=datetime%20desc", timeout=60)
        print("   ", r.status_code, r.text[:700])
    except Exception as e:
        print(ds, "ERR", repr(e))

print("### CBS")
for tb in ["85005NED", "82610NED"]:
    try:
        r = S.get(f"https://opendata.cbs.nl/ODataApi/OData/{tb}/DataProperties?$format=json&$select=Key,Title,Unit,Type", timeout=60)
        for x in r.json()["value"][:60]: print("  ", tb, x.get("Key"), "|", (x.get("Title") or "")[:70], "|", x.get("Unit"), "|", x.get("Type"))
        r = S.get(f"https://opendata.cbs.nl/ODataApi/OData/{tb}/Perioden?$format=json", timeout=60)
        print("  periods", [x["Key"] for x in r.json()["value"]][-20:])
    except Exception as e: print(tb, "ERR", repr(e))

print("### Luxembourg")
for q in ["Creos", "solaire", "photovolta", "eolien", "puissance installee", "ILR", "STATEC energie", "electricite"]:
    r = S.get("https://data.public.lu/api/1/datasets/", params={"q": q, "page_size": 12}, timeout=60)
    print(q, r.status_code, [(d["title"][:70], d.get("frequency"), (d.get("last_update") or "")[:10]) for d in r.json().get("data", [])])
