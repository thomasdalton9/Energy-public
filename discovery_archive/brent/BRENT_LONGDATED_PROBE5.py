"""Probe 5: how far out does any real Brent/WTI futures source go? Yahoo BZ/CL/BRN tickers Jan 2028-Dec 2035, other quote pages, EIA AEO."""
import os, re, json, datetime as dt
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*"}
def g(u, **k):
    try:
        return cr.get(u, impersonate="chrome", timeout=30, headers=k.pop("headers", H), **k)
    except Exception as e:
        class R: status_code = -1; text = type(e).__name__ + str(e)[:100]; headers = {}
        return R()
CODES = "FGHJKMNQUVXZ"
print("=== YAHOO per-contract")
for root, sufs in (("BZ", [".NYM"]), ("CL", [".NYM"]), ("BRN", [".NYM", ".IPE", ".ICE"]), ("B", [".NYM", ".IPE"])):
    for y in range(2028, 2036):
        found = []
        for m in range(12):
            for suf in sufs:
                t = f"{root}{CODES[m]}{str(y)[2:]}{suf}"
                r = g(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}?range=max&interval=1d")
                if r.status_code == 200:
                    try:
                        res = r.json()["chart"]["result"][0]
                        ts = [a for a, b in zip(res["timestamp"], res["indicators"]["quote"][0]["close"]) if b is not None]
                        found.append(f"{CODES[m]}{suf}:{len(ts)}rows {dt.date.fromtimestamp(ts[0])}..{dt.date.fromtimestamp(ts[-1])} last={res['indicators']['quote'][0]['close'][-1]}")
                    except Exception as e:
                        found.append(f"{CODES[m]}{suf}:parse {e}")
                elif r.status_code not in (404,):
                    found.append(f"{CODES[m]}{suf}:HTTP{r.status_code}")
        print(root, y, found or "none", flush=True)
print("=== OTHER SOURCES")
urls = [
 "https://www.cmegroup.com/CmeWS/mvc/Quotes/Future/1/G",
 "https://www.cmegroup.com/CmeWS/mvc/quotes/v2/425?quoteCodes=null&_=1",
 "https://www.cmegroup.com/CmeWS/mvc/quotes/v2/267?quoteCodes=null&_=1",
 "https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/425/FUT?tradeDate=09/30/2026&strategy=DEFAULT&pageSize=500",
 "https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/267/FUT?tradeDate=09/30/2026&strategy=DEFAULT&pageSize=500",
 "https://www.theice.com/marketdata/DelayedMarkets.shtml?getContractsAsJson=&productId=219&hubId=403",
 "https://www.theice.com/marketdata/DelayedMarkets.shtml?getContractsAsJson=&productId=254&hubId=403",
 "https://www.theice.com/products/219/Brent-Crude-Futures/data?marketId=5490784",
 "https://www.investing.com/commodities/brent-oil-futures",
 "https://www.marketwatch.com/investing/future/brn00",
 "https://www.wsj.com/market-data/quotes/futures/UK/IFEU/BRN00",
 "https://www.barchart.com/futures/quotes/CB*0/futures-prices",
 "https://www.barchart.com/proxies/core-api/v1/quotes/get?symbols=CBZ33,CBZ30,CBZ29&fields=symbol,lastPrice,tradeTime",
 "https://www.tradingview.com/symbols/ICEEUR-BRN1!/",
 "https://scanner.tradingview.com/futures/scan",
 "https://stooq.com/q/l/?s=lcz30.f&f=sd2t2ohlcv&h&e=csv",
 "https://stooq.com/q/l/?s=lcz27.f&f=sd2t2ohlcv&h&e=csv",
 "https://www.ft.com/markets/commodities",
 "https://markets.ft.com/data/commodities/tearsheet/summary?c=Brent+Crude+Oil",
 "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.quotes.html",
 "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html",
]
for u in urls:
    r = g(u)
    print(r.status_code, u, "|", (r.text or "")[:200].replace("\n", " "), flush=True)
# TradingView scanner POST
try:
    r = cr.post("https://scanner.tradingview.com/futures/scan", impersonate="chrome", timeout=30,
                json={"filter": [{"left": "name", "operation": "match", "right": "BRN"}], "columns": ["name", "close"], "range": [0, 200]})
    print("TV", r.status_code, r.text[:600])
except Exception as e:
    print("TV err", e)
# CME settlements via page-set cookies
try:
    s = cr.Session()
    s.get("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html", impersonate="chrome", timeout=30)
    for pid in (425, 267, 6):
        r = s.get(f"https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{pid}/FUT?tradeDate=10/02/2026&strategy=DEFAULT&pageSize=500",
                  impersonate="chrome", timeout=30, headers={**H, "Referer": "https://www.cmegroup.com/"})
        print("CMEsess", pid, r.status_code, r.text[:300])
except Exception as e:
    print("CME err", e)
print("=== EIA AEO")
key = os.environ.get("EIA_API_KEY", "")
def eia(path, **p):
    p["api_key"] = key
    r = g("https://api.eia.gov/v2/" + path + "?" + "&".join(f"{k}={v}" for k, v in p.items()))
    return r
for yr in (2025, 2026, 2027):
    r = eia(f"aeo/{yr}")
    print("AEO", yr, r.status_code, r.text[:500])
