"""Probe free sources for the Brent forward curve (settlements by delivery month). Prints status + snippet."""
import json, os, sys, time
import requests
try:
    from curl_cffi import requests as cr
except Exception:
    cr = None

def show(name, fn):
    try:
        r = fn()
        print(f"[{name}] {r.status_code} len={len(r.content)} ct={r.headers.get('content-type')}\n    {r.text[:400]!r}", flush=True)
        return r
    except Exception as e:
        print(f"[{name}] ERR {type(e).__name__}: {str(e)[:200]}", flush=True)

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html"}
def get(url, imp=True, **kw):
    if imp and cr:
        return cr.get(url, impersonate="chrome", timeout=40, headers=kw.pop("headers", None), **kw)
    return requests.get(url, headers=kw.pop("headers", H), timeout=40, **kw)

print("curl_cffi", bool(cr))
# CME product ids
for imp in (True, False):
    for pid in (425, 424, 5412, 2321):
        show(f"CME pid{pid} imp={imp}", lambda: get(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/{pid}/FUT?strategy=DEFAULT&tradeDate=10022026&pageSize=500", imp, headers=H))
show("CME product search", lambda: get("https://www.cmegroup.com/CmeWS/mvc/ProductSlate/V2/List?matchType=contains&q=brent&pageSize=20", True, headers=H))
show("CME BZ page", lambda: get("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html", True, headers=H))
# Yahoo
for t in ("BZV26.NYM", "BZZ26.NYM", "BZK27.NYM", "BZ=F", "CLV26.NYM", "CLK27.NYM", "BRN=F"):
    for host in ("query1", "query2"):
        show(f"Yahoo {t} {host}", lambda: get(f"https://{host}.finance.yahoo.com/v8/finance/chart/{t}?range=max&interval=1d", True, headers=H))
# Stooq
for s in ("cb.f", "lcoz26.f", "cl.f"):
    show(f"Stooq {s}", lambda: get(f"https://stooq.com/q/d/l/?s={s}&i=d", True))
# Nasdaq Data Link
for c in ("CHRIS/ICE_B1", "CHRIS/ICE_B2", "CHRIS/CME_BZ1"):
    show(f"NDL {c}", lambda: get(f"https://data.nasdaq.com/api/v3/datasets/{c}.json?rows=3", True))
# EIA
key = os.environ.get("EIA_API_KEY")
show("EIA API fut", lambda: get(f"https://api.eia.gov/v2/petroleum/pri/fut/data/?api_key={key}&frequency=daily&data[0]=value&sort[0][column]=period&sort[0][direction]=desc&length=8", False))
show("EIA RCLC4 xls", lambda: get("https://www.eia.gov/dnav/pet/hist_xls/RCLC4d.xls", False))
# ICE
for u in ("https://www.ice.com/marketdata/api/productguide/charting/data/historical?productId=254&hubId=403&marketId=0",
          "https://www.ice.com/products/219/Brent-Crude-Futures/data?marketId=5&span=1",
          "https://www.ice.com/marketdata/reports/10/index.html",
          "https://www.ice.com/api/productguide/spec/219/expiry/csv"):
    show("ICE " + u[-60:], lambda: get(u, True, headers=H))
# Barchart / Investing
show("Barchart", lambda: get("https://www.barchart.com/futures/quotes/CB*0/futures-prices", True))
