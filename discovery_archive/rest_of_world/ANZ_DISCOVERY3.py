"""
One-off discovery, round 3, for the Australia + NZ master (prints only): layouts of Hydro Tasmania energy in
storage history, WEM facility SCADA (old monthly CSV and WEMDE daily zip), MBIE monthly gas rows, MBIE
electricity capacity tables, STTM archive, and candidate NZ hydro storage sources (Transpower / EMI reports).
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
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:35]:35}  {r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)
        return None


def links(r, pat, n=30):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    print(f"   ({len(hit)} matching of {len(found)} links)")
    return hit


def peek_csv(content, n=3, name=""):
    try:
        d = pd.read_csv(io.BytesIO(content), nrows=300, low_memory=False)
        print(f"   {name} columns:", list(d.columns)[:40])
        print(d.head(n).to_string()[:2000])
        return d
    except Exception as e:  # noqa: BLE001
        print("   (not CSV)", type(e).__name__, str(e)[:120], "| first bytes:", content[:300])


def peek_xl(content, rows=12):
    for engine in (None, "xlrd", "openpyxl"):
        try:
            xl = pd.ExcelFile(io.BytesIO(content), engine=engine)
            break
        except Exception as e:  # noqa: BLE001
            print("   excel open failed", engine, type(e).__name__, str(e)[:100])
    else:
        return
    print("   sheets:", xl.sheet_names)
    for sh in xl.sheet_names[:6]:
        d = pd.read_excel(xl, sheet_name=sh, header=None)
        print(f"   --- {sh} shape {d.shape}")
        print(d.iloc[:rows, :10].to_string(max_colwidth=30)[:2500])
        print(d.iloc[-3:, :10].to_string(max_colwidth=20)[:800])


section("Hydro Tasmania energy in storage history")
r = get("https://www.hydro.com.au/docs/energyinstorage/download/EnergyInStorage-HistoricalData.xls")
if r is not None and r.ok:
    peek_xl(r.content)

section("Snowy Hydro lake levels")
links(get("https://www.snowyhydro.com.au/generation/live-data/lake-levels/"), r"\.(csv|json|xlsx?)|api|level", 20)

section("WEM old facility SCADA monthly CSV")
r = get("https://data.wa.aemo.com.au/public/public-data/datafiles/facility-scada/facility-scada-2023-08.csv")
if r is not None and r.ok:
    peek_csv(r.content, 3, "facility-scada-2023-08")
links(get("https://data.wa.aemo.com.au/public/public-data/datafiles/facility-scada/"), r"202[3-6]", 10)

section("WEMDE facility SCADA daily zip / current json")
r = get("https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/previous/")
z = links(r, r"FacilityScada_2026", 3) if r is not None else []
lst = sorted(set(re.findall(r'FacilityScada_(\d{8})\.zip', r.text))) if r is not None and r.ok else []
print("   range:", lst[:1], lst[-3:])
if lst:
    rr = get(f"https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/previous/FacilityScada_{lst[-1]}.zip")
    if rr is not None and rr.ok:
        zz = zipfile.ZipFile(io.BytesIO(rr.content))
        print("   members:", [(i.filename, i.file_size) for i in zz.infolist()][:5])
        first = zz.read(zz.namelist()[0])
        print("   first bytes:", first[:1500])
r = get("https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/schema/")
links(r, r"json|schema", 5)

section("MBIE monthly gas: row labels")
r = get("https://www.mbie.govt.nz/assets/Data-Files/Energy/monthly-gas-webtable-sept-2026.xlsx")
if r is not None and r.ok:
    d = pd.read_excel(io.BytesIO(r.content), sheet_name="Monthly_PJ", header=None)
    for i in range(min(len(d), 35)):
        row = d.iloc[i]
        print(f"   r{i}: {[str(x)[:30] for x in row.iloc[:4].tolist()]} ... {[str(x)[:20] for x in row.iloc[-3:].tolist()]}")

section("MBIE electricity statistics (capacity tables)")
r = get("https://www.mbie.govt.nz/building-and-energy/energy-and-natural-resources/energy-statistics-and-modelling/energy-statistics/electricity-statistics/")
x = links(r, r"\.xlsx?$")
for u in x[:3]:
    u = u if u.startswith("http") else "https://www.mbie.govt.nz" + u
    rr = get(u)
    if rr is not None and rr.ok:
        xl = pd.ExcelFile(io.BytesIO(rr.content))
        print("   sheets:", xl.sheet_names)
        for sh in xl.sheet_names:
            if re.search(r"capac", sh, re.I):
                dd = pd.read_excel(xl, sheet_name=sh, header=None)
                print(dd.iloc[:25, :8].to_string(max_colwidth=30)[:3000])

section("NZ hydro storage candidates")
for u in ("https://www.transpower.co.nz/system-operator/security-supply/hydro-storage-and-risk-curves",
          "https://www.transpower.co.nz/system-operator/security-supply",
          "https://www.emi.ea.govt.nz/Wholesale/Reports/W_HS_C?_si=v|3",
          "https://www.emi.ea.govt.nz/Wholesale/Reports/W_HS_C?DateFrom=20250101&DateTo=20260930&_rsdr=ALL&_si=v|3&_export=csv",
          "https://www.nzx.com/services/market-data-information/hydro-storage"):
    r = get(u)
    links(r, r"\.(csv|xlsx?)|storage|hydro", 20)
    if r is not None and r.ok and "csv" in r.headers.get("content-type", ""):
        peek_csv(r.content, 5, u)

section("STTM archive")
links(get("https://nemweb.com.au/Reports/Archive/STTM/"), r"\.zip$|/$", 20)
links(get("https://nemweb.com.au/Reports/Current/STTM/"), r"int651|\.csv$", 20)
