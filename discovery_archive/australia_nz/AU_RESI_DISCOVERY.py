"""
One-off discovery (prints only), run in GitHub Actions: where is Australian east-coast residential / commercial
(distribution-network) gas demand published?

 1. GBB: facility types and names in GasBBActualFlowStorageLast31.CSV and GasBBFacilities.CSV (is there a
    distribution / BDIST type? which facilities?)
 2. STTM: every report in the Current/STTM folder and DayNN.zip - header row, hub column, and any column whose name
    mentions demand / withdrawal / customer / consumption / allocation
 3. DWGM (VicGas): same for the Current/VicGas folder (INT310 price and withdrawals, INT138 etc.)
 4. GSOO / Australian Energy Statistics (DCCEEW) / Australian Energy Update: links found on the landing pages
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
NEM = "https://nemweb.com.au"
KEY = re.compile(r"demand|withdraw|customer|consum|alloc|quantity|qty", re.I)


def get(url):
    try:
        r = requests.get(url, headers=H, timeout=(10, 180))
        print(f"{r.status_code} {len(r.content):>10,} B  {url}", flush=True)
        return r if r.ok else None
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:120]}  {url}", flush=True)


def links(url, pat=r"."):
    r = get(url)
    if r is None:
        return []
    out = [u for u in re.findall(r'href="([^"]+)"', r.text, re.I) if re.search(pat, u, re.I)]
    print(f"   {len(out)} links; first {out[:5]} last {out[-5:]}", flush=True)
    return out


def peek(name, content):
    try:
        d = pd.read_csv(io.BytesIO(content), low_memory=False)
    except Exception as e:  # noqa: BLE001
        print(f"   {name}: not a plain CSV ({type(e).__name__})")
        return
    hit = [c for c in d.columns if KEY.search(str(c))]
    print(f"   {name}: {d.shape} cols={list(d.columns)[:30]}\n      demand-like: {hit}")
    if hit:
        print(d[list(d.columns[:4]) + hit[:3]].head(3).to_string()[:600])


def folder(url):
    fl = links(url, r"\.(csv|zip)$")
    seen = set()
    for u in fl:
        n = u.rsplit("/", 1)[-1]
        key = re.sub(r"_\d+\.csv$", "", n.lower())
        if n.lower().endswith(".csv") and key not in seen:
            seen.add(key)
            r = get(NEM + u if u.startswith("/") else u)
            if r is not None:
                peek(n, r.content)
    return fl


print("=== 1 GBB")
r = get(NEM + "/Reports/Current/GBB/GasBBActualFlowStorageLast31.CSV")
if r is not None:
    d = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    print(d.columns.tolist())
    print(d.groupby("FacilityType")["Demand"].agg(["count", "sum"]).to_string())
    for t in sorted(d["FacilityType"].dropna().unique()):
        print(f"   {t}: {sorted(d[d['FacilityType'] == t]['FacilityName'].unique())[:25]}")
r = get(NEM + "/Reports/Current/GBB/GasBBFacilities.CSV")
if r is not None:
    f = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    print(f.columns.tolist())
    tcol = next((c for c in f.columns if "type" in c.lower()), None)
    print(f[tcol].value_counts().to_string() if tcol else f.head())
links(NEM + "/Reports/Current/GBB/", r"\.(csv|zip)$")

print("=== 2 STTM")
fl = folder(NEM + "/Reports/Current/STTM/")
dz = [u for u in fl if re.search(r"day\d+\.zip", u, re.I)]
if dz:
    r = get(NEM + dz[0] if dz[0].startswith("/") else dz[0])
    if r is not None:
        z = zipfile.ZipFile(io.BytesIO(r.content))
        print("   Day zip members:", z.namelist()[:80])
        done = set()
        for n in z.namelist():
            k = re.sub(r"_\d{8}.*$", "", n.lower())
            if k not in done and n.lower().endswith(".csv"):
                done.add(k)
                peek(n, z.read(n))

print("=== 3 DWGM / VicGas")
folder(NEM + "/Reports/Current/VicGas/")

print("=== 4 annual sources")
for u in ("https://aemo.com.au/energy-systems/gas/gas-bulletin-board-gbb/data-gbb",
          "https://aemo.com.au/energy-systems/gas/gas-services-information/gas-statement-of-opportunities-gsoo",
          "https://www.energy.gov.au/data/australian-energy-statistics",
          "https://www.energy.gov.au/energy-data/australian-energy-statistics",
          "https://www.energy.gov.au/data/australian-energy-update"):
    links(u, r"\.xlsx?|table.?f|gas|gsoo|consumption")
