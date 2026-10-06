"""Probe 4: CME JS endpoint strings; Yahoo per-contract row density by year; interval=1d, period1/period2 query."""
import re, collections, datetime as dt
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html"}
g = lambda u, **k: cr.get(u, impersonate="chrome", timeout=40, headers=H, **k)
p = g("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html").text
m = re.findall(r'/etc\.clientlibs/[^"\']*product-settlements[^"\']*\.js', p)[0]
js = g("https://www.cmegroup.com" + m).text
for mm in re.finditer(r"(ajax|url|getJSON|fetch|settlements|tradeDate)", js):
    s = js[max(0, mm.start()-60): mm.end()+160].replace("\n", " ")
    print("JS>", s)
    if mm.start() > 20000: break
for t in ("BZZ27.NYM", "BZZ26.NYM", "CLZ27.NYM"):
    r = g(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}?range=max&interval=1d")
    res = r.json()["chart"]["result"][0]
    ts, c = res["timestamp"], res["indicators"]["quote"][0]["close"]
    cnt = collections.Counter(dt.date.fromtimestamp(a).year for a, b in zip(ts, c) if b is not None)
    print("DENS", t, dict(sorted(cnt.items())), "nulls", sum(b is None for b in c))
    print("FIRST", [(dt.date.fromtimestamp(a).isoformat(), round(b, 2)) for a, b in zip(ts, c) if b is not None][:6])
