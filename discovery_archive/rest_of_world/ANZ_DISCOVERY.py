"""
One-off discovery for the Australia + New Zealand master: which endpoints answer, what format, what
columns. Prints only - writes nothing. Sources probed:
  NZ  Electricity Authority EMI datasets (blob listing): generation by plant (Generation_MD), hydro
      storage (Hydrological), generating plant list; MBIE gas statistics page (xlsx links)
  AU  AEMO WEM (Western Australia) facility SCADA + facilities list; NEM registration list (capacity);
      Gas Bulletin Board facility types (storage, LNG); STTM / DWGM gas price report folders;
      Hydro Tasmania and BoM water-storage pages (hydro storage)
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def section(t):
    print("\n" + "=" * 25, t, "=" * 25, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=90, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:40]:40}  {url}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)
        return None


def links(r, pat, n=25):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    return hit


def peek_csv(r, n=3):
    if r is None or not r.ok:
        return
    try:
        d = pd.read_csv(io.BytesIO(r.content), nrows=200)
        print("   columns:", list(d.columns)[:25])
        print(d.head(n).to_string()[:1500])
    except Exception as e:  # noqa: BLE001
        print("   (not CSV)", type(e).__name__, str(e)[:120], "| first bytes:", r.content[:200])


def blob_list(prefix, n=15):
    url = "https://emidatasets.blob.core.windows.net/publicdata"
    r = get(url, params={"restype": "container", "comp": "list", "prefix": prefix, "maxresults": 5000})
    if r is None or not r.ok:
        return []
    names = re.findall(r"<Name>([^<]+)</Name>", r.text)
    print(f"   {len(names)} blobs under {prefix}; last {n}:")
    for x in names[-n:]:
        print("   ", x)
    return names


section("NZ EMI: generation by plant (Generation_MD)")
g = blob_list("Datasets/Wholesale/Generation/Generation_MD/")
if g:
    peek_csv(get("https://emidatasets.blob.core.windows.net/publicdata/" + g[-1]))

section("NZ EMI: hydrological (storage)")
h = blob_list("Datasets/Wholesale/Hydrological/", 30)
csvs = [x for x in h if x.lower().endswith(".csv")]
for x in csvs[-3:]:
    peek_csv(get("https://emidatasets.blob.core.windows.net/publicdata/" + x))

section("NZ EMI: generation folder (plant lists)")
gen = blob_list("Datasets/Wholesale/Generation/", 40)
plant = [x for x in gen if re.search(r"plant|station|capacity", x, re.I)]
print("   plant-like:", plant[:15])
for x in plant[-2:]:
    peek_csv(get("https://emidatasets.blob.core.windows.net/publicdata/" + x))

section("NZ MBIE gas statistics")
r = get("https://www.mbie.govt.nz/building-and-energy/energy-and-natural-resources/energy-statistics-and-modelling/energy-statistics/gas-statistics/")
gl = links(r, r"\.xlsx?$|gas")
x = [u for u in gl if u.lower().endswith((".xlsx", ".xls"))]
if x:
    u = x[0] if x[0].startswith("http") else "https://www.mbie.govt.nz" + x[0]
    rr = get(u)
    if rr is not None and rr.ok:
        xl = pd.ExcelFile(io.BytesIO(rr.content))
        print("   sheets:", xl.sheet_names[:20])

section("AU WEM (Western Australia)")
r = get("https://data.wa.aemo.com.au/public/public-data/datafiles/facilities/facilities.csv")
peek_csv(r)
for m in ("2026-08", "2026-09"):
    peek_csv(get(f"https://data.wa.aemo.com.au/public/public-data/datafiles/facility-scada/facility-scada-{m}.csv"), 2)
links(get("https://data.wa.aemo.com.au/"), r"datafiles|scada|facilit")

section("AU NEM registration list (capacity)")
r = get("https://www.aemo.com.au/-/media/Files/Electricity/NEM/Participant_Information/NEM-Registration-and-Exemption-List.xls")
if r is not None and r.ok:
    try:
        xl = pd.ExcelFile(io.BytesIO(r.content), engine="openpyxl")
        print("   sheets:", xl.sheet_names)
        d = pd.read_excel(xl, "PU and Scheduled Loads")
        print("   columns:", list(d.columns))
        print(d.head(2).to_string()[:1200])
    except Exception as e:  # noqa: BLE001
        print("   read failed", e)

section("AU Gas Bulletin Board facilities (storage / LNG)")
r = get("https://nemweb.com.au/Reports/Current/GBB/GasBBFacilities.CSV")
if r is not None and r.ok:
    d = pd.read_csv(io.BytesIO(r.content))
    print("   columns:", list(d.columns))
    tcol = next((c for c in d.columns if "type" in c.lower()), None)
    if tcol:
        print(d[tcol].value_counts().to_string())
        ncol = next((c for c in d.columns if "name" in c.lower()), d.columns[0])
        print(d[d[tcol].astype(str).str.contains("STOR|LNG", case=False)][[ncol, tcol]].to_string()[:2000])
links(get("https://nemweb.com.au/Reports/Current/GBB/"), r"\.(zip|csv)$", 40)

section("AU gas prices: STTM / DWGM folders")
links(get("https://nemweb.com.au/Reports/Current/STTM/"), r"\.(zip|csv)$", 15)
links(get("https://nemweb.com.au/Reports/Current/VicGas/"), r"\.(zip|csv)$", 15)

section("AU hydro storage: Hydro Tasmania / BoM")
links(get("https://www.hydro.com.au/water/energy-storage-and-water-levels"), r"\.(csv|xlsx?|json)|storage|level", 20)
links(get("https://www.bom.gov.au/water/dashboards/"), r"storage|csv|json", 20)
r = get("http://www.bom.gov.au/waterdata/services?service=SOS&version=2.0&request=GetCapabilities")
print("   BoM SOS capabilities head:", (r.text[:300] if r is not None else ""))
