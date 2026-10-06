"""Probe 6: CME Brent Last Day product id + settlements history/depth; TradingView scanner ICEEUR BRN contracts; EIA AEO Brent series."""
import os, re, json, time
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/"}
def g(u, **k):
    try:
        return cr.get(u, impersonate="chrome", timeout=30, headers=H, **k)
    except Exception as e:
        class R: status_code = -1; text = type(e).__name__; 
        return R()
print("=== CME product id")
p = g("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html").text
print(sorted(set(re.findall(r'productId["\': =]+(\d+)', p)))[:20], re.findall(r'.{60}/CmeWS/mvc.{80}', p)[:3])
for pid in list(range(420, 432)) + [1111, 5201]:
    r = g(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{pid}/FUT?tradeDate=10/05/2026&strategy=DEFAULT&pageSize=500")
    try:
        j = r.json(); print(pid, r.status_code, j.get("dsHeader"), len(j.get("settlements", [])), [s["month"] for s in j.get("settlements", [])][-3:])
    except Exception:
        print(pid, r.status_code, r.text[:80])
print("=== CME BZ detail")
bz = None
for pid in range(420, 432):
    r = g(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{pid}/FUT?tradeDate=10/05/2026&strategy=DEFAULT&pageSize=500")
    try:
        if "Brent" in (r.json().get("dsHeader") or ""): bz = pid
    except Exception: pass
print("BZ id", bz)
if bz:
    for td in ("10/05/2026", "10/02/2026", "09/30/2026", "12/31/2025", "06/30/2024", "12/29/2023", "12/30/2022"):
        r = g(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{bz}/FUT?tradeDate={td}&strategy=DEFAULT&pageSize=500")
        try:
            j = r.json(); s = j.get("settlements", [])
            print(td, r.status_code, len(s), [(x["month"], x["settle"]) for x in s[:2]], [(x["month"], x["settle"]) for x in s[-3:]])
        except Exception:
            print(td, r.status_code, r.text[:100])
        time.sleep(1)
    r = g(f"https://www.cmegroup.com/CmeWS/mvc/quotes/v2/{bz}?quoteCodes=null&_=1")
    try:
        j = r.json(); q = j.get("quotes", [])
        print("quotes", len(q), [(x.get("quoteCode"), x.get("expirationMonth"), x.get("last")) for x in q[-4:]])
    except Exception:
        print(r.status_code, r.text[:100])
print("=== TradingView")
def tv(rt, cols):
    r = cr.post("https://scanner.tradingview.com/futures/scan", impersonate="chrome", timeout=30,
                json={"filter": [{"left": "name", "operation": "match", "right": rt}], "columns": cols, "range": [0, 500], "sort": {"sortBy": "name", "sortOrder": "asc"}})
    return r.status_code, r.json() if r.status_code == 200 else r.text[:200]
st, j = tv("BRN", ["name", "close", "update_mode", "expiration", "currency", "pricescale", "volume", "change"])
print(st, j["totalCount"] if st == 200 else j)
if st == 200:
    for d in j["data"]:
        if d["s"].startswith("ICEEUR:BRN"): print(d["s"], d["d"])
st, j = tv("BZ", ["name", "close", "expiration"])
if st == 200:
    print("BZ rows", [(d["s"], d["d"][1]) for d in j["data"] if d["s"].startswith("NYMEX:BZ")][-12:])
print("=== EIA AEO")
key = os.environ.get("EIA_API_KEY", "")
def eia(path, qs=""):
    for i in range(4):
        r = g(f"https://api.eia.gov/v2/{path}?api_key={key}{qs}")
        if r.status_code != 429: return r
        time.sleep(20)
    return r
r = eia("aeo/2026/facet/seriesId")
try:
    js = r.json()["response"]["facets"]
    print("series count", len(js), [f for f in js if "brent" in (f["id"] + f.get("name", "")).lower() or "world" in f.get("name", "").lower()][:20])
except Exception: print(r.status_code, r.text[:300])
r = eia("aeo/2026/facet/tableId")
try: print([f for f in r.json()["response"]["facets"] if "rice" in f.get("name", "")])
except Exception: print(r.status_code, r.text[:200])
r = eia("aeo/2026/facet/scenario")
try: print(r.json()["response"]["facets"])
except Exception: print(r.status_code, r.text[:200])
