"""
One-off discovery, round 2, for the Australia + NZ master (prints only): folder listings for the
probes that missed in ANZ_DISCOVERY.py - NZ EMI hydro storage location, WEM SCADA path, AU hydro
storage (BoM SOS offerings) - plus column layouts of GBB flow/storage, STTM, MBIE monthly gas.
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
BLOB = "https://emidatasets.blob.core.windows.net/publicdata"


def section(t):
    print("\n" + "=" * 25, t, "=" * 25, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=90, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:35]:35}  {r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)
        return None


def links(r, pat, n=40):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    print(f"   ({len(hit)} matching of {len(found)} links)")
    return hit


def peek(content, n=3, name=""):
    try:
        d = pd.read_csv(io.BytesIO(content), nrows=300, low_memory=False)
        print(f"   {name} columns:", list(d.columns)[:40])
        print(d.head(n).to_string()[:2000])
        return d
    except Exception as e:  # noqa: BLE001
        print("   (not CSV)", type(e).__name__, str(e)[:120], "| first bytes:", content[:300])


def folders(prefix):
    r = get(BLOB, params={"restype": "container", "comp": "list", "prefix": prefix, "delimiter": "/"})
    if r is None or not r.ok:
        return [], []
    pre = re.findall(r"<BlobPrefix><Name>([^<]+)</Name>", r.text)
    blobs = re.findall(r"<Blob><Name>([^<]+)</Name>", r.text)
    print(f"   {prefix}: {len(pre)} folders, {len(blobs)} files")
    for p in pre:
        print("    dir ", p)
    for b in blobs[-10:]:
        print("    file", b)
    return pre, blobs


section("NZ EMI folder tree (2 levels) - find hydro storage")
top, _ = folders("Datasets/")
for p in top:
    sub, _ = folders(p)
    for s in sub:
        if re.search(r"hydro|storage|lake|inflow|wholesale", s, re.I):
            sub2, files = folders(s)
            for s2 in sub2:
                if re.search(r"hydro|storage|lake|inflow", s2, re.I):
                    _, f2 = folders(s2)
                    csvs = [x for x in f2 if x.lower().endswith(".csv")]
                    if csvs:
                        r = get(f"{BLOB}/{csvs[-1]}")
                        if r is not None and r.ok:
                            peek(r.content, name=csvs[-1])

section("NZ: Gas storage (Ahuroa) / GIC")
links(get("https://www.gasindustry.co.nz/about-the-gas-industry/gas-storage/"), r"\.(csv|xlsx?)|storage|ahuroa")
links(get("https://www.gasindustry.co.nz/about-the-gas-industry/production-consumption/"), r"\.(csv|xlsx?)")

section("NZ MBIE monthly gas webtable layout")
r = get("https://www.mbie.govt.nz/assets/Data-Files/Energy/monthly-gas-webtable-sept-2026.xlsx")
if r is not None and r.ok:
    for sh in ("Contents", "Monthly_PJ"):
        d = pd.read_excel(io.BytesIO(r.content), sheet_name=sh, header=None)
        print(f"   --- {sh} shape {d.shape}")
        print(d.head(14).to_string(max_colwidth=40)[:4000])
        print(d.tail(4).to_string(max_colwidth=25)[:2000])

section("AU WEM listings")
for u in ("https://data.wa.aemo.com.au/public/public-data/datafiles/",
          "https://data.wa.aemo.com.au/public/public-data/datafiles/facility-scada/",
          "https://data.wa.aemo.com.au/public/market-data/wemde/",
          "https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/",
          "https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/previous/",
          "https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/current/"):
    hit = links(get(u), r"scada|\.zip$|\.csv$|/$", 25)

section("AU GBB actual flow / storage layout")
r = get("https://nemweb.com.au/Reports/Current/GBB/GasBBActualFlowStorageLast31.CSV")
if r is not None and r.ok:
    d = peek(r.content, 5, "Last31")
    if d is not None:
        for c in d.columns:
            if d[c].dtype == object and d[c].nunique() < 40:
                print(f"   {c}: {sorted(d[c].dropna().astype(str).unique())[:40]}")
r = get("https://nemweb.com.au/Reports/Current/GBB/GasBBLNGShipments.CSV")
if r is not None and r.ok:
    peek(r.content, 3, "LNG shipments")
r = get("https://nemweb.com.au/Reports/Current/GBB/GasBBActualFlowStorage.zip")
if r is not None and r.ok:
    z = zipfile.ZipFile(io.BytesIO(r.content))
    print("   zip members:", [(i.filename, i.file_size) for i in z.infolist()][:10])

section("AU STTM Day01 contents")
r = get("https://nemweb.com.au/Reports/Current/STTM/Day01.zip")
if r is not None and r.ok:
    z = zipfile.ZipFile(io.BytesIO(r.content))
    names = z.namelist()
    print("   members:", names[:60])
    for n in names:
        if re.search(r"int651|ex_ante_market_price|int652|price", n, re.I):
            peek(z.read(n), 3, n)
            break

section("AU hydro storage: BoM SOS storage offerings / Hydro Tasmania")
r = get("http://www.bom.gov.au/waterdata/services?service=SOS&version=2.0&request=GetCapabilities")
if r is not None and r.ok:
    offs = sorted(set(re.findall(r"<sos:identifier>([^<]+)</sos:identifier>", r.text)
                      + re.findall(r"<swes:identifier>([^<]+)</swes:identifier>", r.text)))
    print(f"   {len(offs)} identifiers; storage-like:", [o for o in offs if re.search(r"stor|volume", o, re.I)][:20])
    props = sorted(set(re.findall(r"<swes:observableProperty>([^<]+)</swes:observableProperty>", r.text)))
    print("   observable properties:", props[:40])
for u in ("https://www.hydro.com.au/clean-energy/hydro-power/water-storage",
          "https://www.hydro.com.au/water/lake-levels",
          "https://www.hydro.com.au/",
          "https://www.snowyhydro.com.au/generation/live-data/",
          "https://www.bom.gov.au/water/dashboards/#/water-storages/summary/state",
          "http://www.bom.gov.au/waterdata/wiski-web-public/"):
    links(get(u), r"storage|lake|level|\.csv|\.json", 15)
