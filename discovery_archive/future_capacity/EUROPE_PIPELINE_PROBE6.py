"""
Probe 6 (compact): ERAA 2025 / 2023 data zips (member lists + heads), CfD AR5-AR7a result workbooks, DUKES 5.11, ODRE 'projets en
developpement' records, Commission NECP page links and the 3 reachable NECP PDFs, nuclear operator pages (link discovery).
"""
import io, re, sys, zipfile
import pandas as pd
import requests
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}


def get(u, **kw):
    r = requests.get(u, headers=UA, timeout=(20, 180), **kw)
    print("GET", r.status_code, len(r.content), u[:150]); return r


def sec(t): print("\n=========", t, flush=True)


def show(d, n=12, w=24):
    for i, row in d.head(n).iterrows():
        cells = ["" if pd.isna(v) else str(v).replace("\n", " ")[:w] for v in row.tolist()]
        if any(cells): print(i, "|".join(cells))


class RangeFile(io.RawIOBase):
    def __init__(self, url):
        r = requests.get(url, headers={**UA, "Range": "bytes=0-0"}, timeout=60)
        self.size = int(r.headers["content-range"].split("/")[1]); self.pos = 0; self.url = url
    def seekable(self): return True
    def readable(self): return True
    def tell(self): return self.pos
    def seek(self, off, whence=0):
        self.pos = off if whence == 0 else self.pos + off if whence == 1 else self.size + off; return self.pos
    def read(self, n=-1):
        if n < 0: n = self.size - self.pos
        n = min(n, self.size - self.pos)
        if n <= 0: return b""
        d = requests.get(self.url, headers={**UA, "Range": f"bytes={self.pos}-{self.pos + n - 1}"}, timeout=120).content
        self.pos += len(d); return d
    def readinto(self, b):
        d = self.read(len(b)); b[:len(d)] = d; return len(d)


B = "https://eepublicdownloads.blob.core.windows.net/public-cdn-container/clean-documents/sdc-documents/ERAA/"
sec("ERAA zips")
for f in ["ERAA_2025_CommonData.zip", "ERAA_2025_Dashboard_RawData.zip", "ERAA_2025_OtherData.zip"]:
    try:
        z = zipfile.ZipFile(io.BufferedReader(RangeFile(B + f), buffer_size=1 << 22))
        print("--", f)
        for i in z.infolist()[:60]: print("   ", i.filename, i.file_size)
    except Exception as e: print("ERR", f, repr(e)[:150])

sec("ERAA dashboard heads")
try:
    z = zipfile.ZipFile(io.BufferedReader(RangeFile(B + "ERAA_2025_Dashboard_RawData.zip"), buffer_size=1 << 22))
    for i in z.infolist():
        if re.search(r"(?i)capac|resource|install|thermal", i.filename) and i.file_size < 80e6:
            print("--", i.filename, i.file_size)
            data = z.read(i.filename)
            if i.filename.lower().endswith((".xlsx", ".xls")):
                x = pd.ExcelFile(io.BytesIO(data))
                for s in x.sheet_names[:6]:
                    d = x.parse(s, header=None); print(" sheet", s, d.shape); show(d, 8)
            else:
                print(data[:900].decode("utf-8", "replace"))
            break
except Exception as e: print("ERR", repr(e)[:200])

sec("CfD workbooks")
urls = ["https://assets.publishing.service.gov.uk/media/6516e8db6dfda6000d8e3904/ar5-successful-projects-list.xls",
        "https://assets.publishing.service.gov.uk/media/6968f9176aedc683a7540ca0/CfD_AR7_results_-_Pot_3_and_4.xlsx",
        "https://assets.publishing.service.gov.uk/media/698a0b4385bc7d6ba0fbc832/Contracts_for_Difference_AR7a.xlsx"]
import json
for ar in ["6", "7"]:
    c = get(f"https://www.gov.uk/api/content/government/publications/contracts-for-difference-cfd-allocation-round-{ar}-results").json()
    for a in c.get("details", {}).get("attachments", []):
        u = a.get("url", "")
        if re.search(r"\.xlsx?$", u): print("ATT", ar, u)
        if ar == "6" and re.search(r"\.xlsx?$", u): urls.append(u)
for u in urls:
    try:
        x = pd.ExcelFile(io.BytesIO(get(u).content))
        for s in x.sheet_names[:6]:
            d = x.parse(s, header=None); print(" sheet", s, d.shape); show(d, 18, 22)
    except Exception as e: print("ERR", u, repr(e)[:150])

sec("DUKES 5.11")
x = pd.ExcelFile(io.BytesIO(get("https://assets.publishing.service.gov.uk/media/6a6a365b0ddb7e4831c629f3/DUKES_5.11.xlsx").content))
print(x.sheet_names)
for s in x.sheet_names[:4]:
    d = x.parse(s, header=None); print(" sheet", s, d.shape); show(d, 10, 22)

sec("ODRE queue")
j = get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/suivi-projet-raccordement-enr/records?limit=100&select=filiere%2Creseau%2Csum(projets_en_developpement_au_31_12_2024_mw)%20as%20mw&group_by=filiere%2Creseau").json()
for r in j.get("results", []): print(r)

sec("NECP")
h = get("https://commission.europa.eu/energy-climate-change-environment/implementation-eu-countries/energy-and-climate-governance-and-reporting/national-energy-and-climate-plans_en").text
print("document/download links:", len(re.findall(r"document/download", h)), "country mentions:", len(re.findall(r"(?i)necp", h)))
for m in list(re.finditer(r'href="([^"]*(?:national-energy-and-climate-plans|necp)[^"]*)"', h, re.I))[:30]: print("  ", m.group(1)[:170])
pl = get("https://commission.europa.eu/document/download/d8800fed-c52a-45bb-a0ca-447c6e01720a_en?filename=PL_FINAL%20UPDATED%20NECP%202021-2030%20%28English%29.pdf")
try:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pl.content)) as p:
        print("pages", len(p.pages))
        for k, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            if re.search(r"(?i)installed capacit", t) and re.search(r"(?i)offshore", t) and re.search(r"GW|MW", t):
                print("--- page", k + 1); print(t[:1500]); break
except Exception as e: print("pdf err", repr(e)[:150])

sec("Nuclear pages")
for u, pat in [("https://www.seas.sk/", r"mochovce"), ("https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c", r"news|press|media|about|timeline"),
               ("https://www.sizewellc.com/", r"about|project|facts|timeline"), ("https://www.paks2.hu/", r"en|projekt"),
               ("https://www.axpo.com/ch/en/about-us/power-generation/nuclear-power.html", r"beznau")]:
    try:
        h = get(u).text
        ls = sorted(set(re.findall(r'href="([^"#]+)"', h)))
        print("  ", u, "->", [l[:90] for l in ls if re.search(pat, l, re.I)][:25])
    except Exception as e: print("ERR", u, repr(e)[:100])
