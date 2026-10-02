"""
One-off discovery, round 6 (prints only): DWGM INT310 price-and-withdrawals history file, GSH folder links (raw),
WA GBB API by gas day, CER small-scale solar installation / capacity document downloads.
"""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=120, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:35]:35}  {r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)


def show_csv(content, n=3):
    d = pd.read_csv(io.BytesIO(content), low_memory=False)
    print("   shape", d.shape, "columns:", list(d.columns)[:25])
    print(d.head(n).to_string()[:1500])
    print(d.tail(2).to_string()[:800])
    return d


print("=== DWGM INT310 / INT041 current")
for f in ("int310_v1_price_and_withdrawals_rpt_1.csv", "int041_v4_market_and_reference_prices_1.csv",
          "int042_v4_weighted_average_daily_prices_1.csv"):
    r = get("https://nemweb.com.au/Reports/CURRENT/VicGas/" + f)
    if r is not None and r.ok:
        show_csv(r.content)

print("=== GSH raw links")
r = get("https://nemweb.com.au/Reports/Current/GSH/")
if r is not None:
    print(re.findall(r'href="([^"]+)"', r.text, re.I))
    for sub in re.findall(r'href="([^"]+)"', r.text, re.I):
        if re.search(r"GSH/.+", sub) and not sub.lower().endswith((".zip", ".csv")):
            rr = get("https://nemweb.com.au" + sub if sub.startswith("/") else sub)
            if rr is not None and rr.ok:
                fl = re.findall(r'href="([^"]+\.(?:zip|csv|CSV|ZIP))"', rr.text)
                print("   ", sub, len(fl), fl[-3:])

print("=== WA GBB by gas day")
for u in ("https://gbbwa.aemo.com.au/api/v1/report/actualFlow/2026-09-29",
          "https://gbbwa.aemo.com.au/api/v1/report/actualFlow/2021-01-04",
          "https://gbbwa.aemo.com.au/api/v1/report/actualFlow?gasDay=2026-09-29",
          "https://gbbwa.aemo.com.au/api/v1/report/actualFlow/2026-09-29/csv",
          "https://gbbwa.aemo.com.au/api/v1/report/actualFlow/current/csv",
          "https://gbbwa.aemo.com.au/api/v1/report/endUserConsumption/current",
          "https://gbbwa.aemo.com.au/api/v1/report/capacityOutlook/current"):
    r = get(u)
    if r is not None and r.ok:
        print("   head:", r.text[:400].replace("\n", " "))

print("=== CER documents")
for doc in ("sgu-solar-installations-2011-to-present-and-totals", "sres-postcode-data-capacity-2011-to-present-and-totals",
            "sgu-battery-installations-2011-to-present-and-totals"):
    r = get(f"https://cer.gov.au/document/{doc}")
    if r is None or not r.ok:
        continue
    if "html" not in r.headers.get("content-type", ""):
        files = [r.url]
        content = {r.url: r.content}
    else:
        files = [u for u in re.findall(r'href="([^"]+)"', r.text) if re.search(r"\.(csv|xlsx?)(\?|$)|/download|files/", u, re.I)]
        print("   file links:", files[:6])
        content = {}
    for f in files[:1]:
        f = f if f.startswith("http") else "https://cer.gov.au" + f
        c = content.get(f) or (get(f).content if get(f) is not None else b"")
        try:
            if c[:2] == b"PK":
                xl = pd.ExcelFile(io.BytesIO(c))
                print("   sheets:", xl.sheet_names)
                print(pd.read_excel(xl, xl.sheet_names[0], header=None).iloc[:8, :14].to_string()[:2000])
            else:
                show_csv(c)
        except Exception as e:  # noqa: BLE001
            print("   parse failed", type(e).__name__, c[:200])
