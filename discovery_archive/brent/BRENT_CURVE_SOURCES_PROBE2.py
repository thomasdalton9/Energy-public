"""Probe 2: CME product id/endpoints from the BZ settlements page; Yahoo per-contract coverage incl. expired."""
import re, json, datetime as dt
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html"}
g = lambda u, **k: cr.get(u, impersonate="chrome", timeout=40, headers=H, **k)
p = g("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html").text
for m in sorted(set(re.findall(r"[^\"'\s]*(?:ettlements|productId|product_id)[^\"'\s]{0,120}", p)))[:40]: print("PAGE", m[:200])
print("tradeDate" in p, [m for m in re.findall(r"CmeWS[^\"'\s]{0,150}", p)][:15])
for pid in (424, 425, 5412, 1018):
    for u in (f"https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/Settlements/{pid}/FUT?strategy=DEFAULT&tradeDate=10022026&pageSize=500",
              f"https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/{pid}/FUT?strategy=DEFAULT&tradeDate=10/02/2026&pageSize=500&isProtected&_t=1"):
        r = g(u); print("CME", pid, r.status_code, r.text[:250].replace("\n", " "))
r = g("https://www.cmegroup.com/CmeWS/mvc/ProductSlate/V2/List?matchType=contains&q=brent&pageSize=50")
for pr in r.json().get("products", []):
    print("PROD", pr.get("id"), pr.get("name"), pr.get("globex"), pr.get("cleared"))
# Yahoo
def y(t, rng="max"):
    r = g(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}?range={rng}&interval=1d")
    if r.status_code != 200: return t, r.status_code, None
    res = r.json()["chart"]["result"][0]; ts = res.get("timestamp") or []
    c = res["indicators"]["quote"][0]["close"]
    ok = [(a, b) for a, b in zip(ts, c) if b is not None]
    return t, 200, (len(ok), dt.date.fromtimestamp(ok[0][0]) if ok else None, dt.date.fromtimestamp(ok[-1][0]) if ok else None, ok[-1][1] if ok else None, res["meta"].get("expireDate"))
codes = "FGHJKMNQUVXZ"
for t in ["BZV26.NYM", "BZX26.NYM", "BZZ26.NYM", "BZF27.NYM", "BZZ27.NYM", "BZZ28.NYM", "BZZ29.NYM", "BZZ30.NYM", "BZZ31.NYM", "BZH26.NYM", "BZG26.NYM", "BZZ25.NYM", "BZH25.NYM", "BZK24.NYM", "BZU26.NYM", "BZQ26.NYM", "BZ=F", "CLZ26.NYM", "CLZ30.NYM", "CLZ25.NYM", "CLV26.NYM", "CLX26.NYM"]:
    print("YH", y(t))
