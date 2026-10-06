"""Probe 7: CME Brent Last Day (424) settlements: which trade dates exist; EIA AEO2026 Brent series; STEO BREPUUS."""
import os, time, json, datetime as dt
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/"}
def g(u):
    try: return cr.get(u, impersonate="chrome", timeout=30, headers=H)
    except Exception as e:
        class R: status_code = -1; text = type(e).__name__
        return R()
def cme(d):
    r = g(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/424/FUT?tradeDate={d:%m/%d/%Y}&strategy=DEFAULT&pageSize=500")
    try:
        j = r.json(); s = [x for x in j.get("settlements", []) if x["month"] != "Total"]
        return len(s), (s[0]["month"], s[0]["settle"], s[-1]["month"]) if s else None, j.get("reportType")
    except Exception: return r.status_code, r.text[:60], None
print("=== CME daily Sep 1 - Oct 6 2026")
d = dt.date(2026, 9, 1)
while d <= dt.date(2026, 10, 6):
    if d.weekday() < 5: print(d, cme(d), flush=True)
    d += dt.timedelta(days=1); time.sleep(0.5)
print("=== CME history samples")
for d in [dt.date(2026, 8, 31), dt.date(2026, 7, 31), dt.date(2026, 6, 30), dt.date(2026, 3, 31), dt.date(2025, 12, 30), dt.date(2025, 6, 30), dt.date(2024, 12, 31), dt.date(2023, 6, 30), dt.date(2021, 6, 30), dt.date(2019, 6, 28)]:
    print(d, cme(d), flush=True); time.sleep(0.5)
key = os.environ.get("EIA_API_KEY", "")
def eia(path, qs):
    for i in range(4):
        r = g(f"https://api.eia.gov/v2/{path}?api_key={key}{qs}")
        if r.status_code != 429: return r
        time.sleep(20)
    return r
print("=== AEO2026 table 12 Brent")
r = eia("aeo/2026/data", "&frequency=annual&data[0]=value&facets[tableId][]=12&length=5000")
try:
    rows = r.json()["response"]["data"]; print("rows", len(rows), "total", r.json()["response"]["total"])
    seen = {}
    for x in rows:
        nm = x.get("seriesName", "")
        if "brent" in nm.lower() or "world oil" in nm.lower() or "imported" in nm.lower():
            seen.setdefault((x["seriesId"], x["scenario"], nm, x.get("unit")), []).append((x["period"], x["value"]))
    for k, v in seen.items(): print(k, sorted(v)[:2], sorted(v)[-2:], len(v))
    print("sample row", rows[0])
except Exception as e: print(r.status_code, r.text[:400], e)
print("=== STEO BREPUUS")
r = eia("steo/data", "&frequency=monthly&data[0]=value&facets[seriesId][]=BREPUUS&sort[0][column]=period&sort[0][direction]=desc&length=40")
try:
    j = r.json()["response"]; print([(x["period"], x["value"]) for x in j["data"]][:40], j["data"][0])
except Exception as e: print(r.status_code, r.text[:300])
r = eia("steo", "")
try: print(r.json()["response"].get("description"), r.json()["response"].get("endPeriod"))
except Exception: print(r.status_code)
