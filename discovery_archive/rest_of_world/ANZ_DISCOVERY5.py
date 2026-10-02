"""
One-off discovery, round 5 (prints only): Victorian DWGM prices (nemweb VicGas reports), Wallumbilla gas supply
hub prices (nemweb GSH), WA Gas Bulletin Board (gbbwa.aemo.com.au / data.wa.aemo.com.au), Clean Energy Regulator
small-scale (rooftop solar) installation data.
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=90, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:35]:35}  {r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)


def links(r, pat, n=40):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    print(f"   ({len(hit)} of {len(found)})")
    return hit


def peek(content, name="", n=3):
    try:
        d = pd.read_csv(io.BytesIO(content), nrows=200, low_memory=False)
        print(f"   {name} columns: {list(d.columns)[:30]}")
        print(d.head(n).to_string()[:1500])
    except Exception as e:  # noqa: BLE001
        print("   (not CSV)", type(e).__name__, content[:300])


print("=== VicGas (DWGM)")
lk = links(get("https://nemweb.com.au/Reports/Current/VicGas/"), r"int0?41|int199|int310|price|\.csv$", 60)
for u in [x for x in lk if re.search(r"int0?41|price", x, re.I)][:3]:
    r = get("https://nemweb.com.au" + u if u.startswith("/") else u)
    if r is not None and r.ok:
        peek(r.content, u)
r = get("https://nemweb.com.au/Reports/Current/VicGas/PublicRpts01.zip")
if r is not None and r.ok:
    z = zipfile.ZipFile(io.BytesIO(r.content))
    names = z.namelist()
    print("   PublicRpts01 members:", [n for n in names if re.search(r"int0?41|price", n, re.I)][:20], len(names))
    for n in names:
        if re.search(r"int041", n, re.I):
            peek(z.read(n), n)
            break

print("=== GSH (Wallumbilla)")
lk = links(get("https://nemweb.com.au/Reports/Current/GSH/"), r"\.(zip|csv)$|/$", 60)
for sub in [x for x in lk if x.endswith("/") and not x.rstrip("/").endswith("Current")][:6]:
    links(get("https://nemweb.com.au" + sub), r"\.(zip|csv)$", 10)
for u in [x for x in lk if re.search(r"price|benchmark|trade", x, re.I) and x.lower().endswith((".csv", ".zip"))][:3]:
    r = get("https://nemweb.com.au" + u)
    if r is not None and r.ok:
        if u.lower().endswith(".zip"):
            z = zipfile.ZipFile(io.BytesIO(r.content))
            print("   members:", z.namelist()[:10])
            peek(z.read(z.namelist()[0]), z.namelist()[0])
        else:
            peek(r.content, u)

print("=== WA GBB")
for u in ("https://gbbwa.aemo.com.au/", "https://gbbwa.aemo.com.au/api/v1/report/actualFlow/current",
          "https://gbbwa.aemo.com.au/api/v1/report/actualFlow/2026-09",
          "https://data.wa.aemo.com.au/public/infrastructure/", "https://data.wa.aemo.com.au/public/public-data/datafiles/",
          "https://data.wa.aemo.com.au/public/infrastructure/gbb/"):
    r = get(u)
    if r is not None and r.ok:
        if "json" in r.headers.get("content-type", ""):
            print("   json head:", r.text[:800])
        else:
            links(r, r"gbb|gas|flow|storage|api|\.csv", 25)

print("=== Clean Energy Regulator small-scale installations")
for u in ("https://cer.gov.au/markets/reports-and-data/small-scale-installation-postcode-data",
          "https://www.cleanenergyregulator.gov.au/RET/Forms-and-resources/Postcode-data-for-small-scale-installations",
          "https://cer.gov.au/markets/reports-and-data/small-scale-installation-postcode-data/small-scale-installation-postcode-data-2026"):
    r = get(u)
    lk = links(r, r"\.(csv|xlsx?)(\?|$)|postcode|installation", 25)
    files = [x for x in lk if re.search(r"\.(csv|xlsx?)(\?|$)", x, re.I)]
    for f in files[:2]:
        f = f if f.startswith("http") else re.match(r"https://[^/]+", r.url).group(0) + f
        rr = get(f)
        if rr is not None and rr.ok:
            if f.lower().split("?")[0].endswith(".csv"):
                peek(rr.content, f)
            else:
                xl = pd.ExcelFile(io.BytesIO(rr.content))
                print("   sheets:", xl.sheet_names)
                print(pd.read_excel(xl, xl.sheet_names[0], header=None).iloc[:8, :10].to_string()[:1500])
