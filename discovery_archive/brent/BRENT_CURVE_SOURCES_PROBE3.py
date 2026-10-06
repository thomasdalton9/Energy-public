"""Probe 3: CME settlements JS endpoint; Yahoo coverage scan of every BZ/CL monthly contract."""
import re, datetime as dt
from curl_cffi import requests as cr
H = {"Accept": "application/json, text/plain, */*", "Referer": "https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html"}
g = lambda u, **k: cr.get(u, impersonate="chrome", timeout=40, headers=H, **k)
p = g("https://www.cmegroup.com/markets/energy/crude-oil/brent-crude-oil-last-day.settlements.html").text
for m in re.findall(r'/etc\.clientlibs/[^"\']*product-settlements[^"\']*\.js', p): 
    js = g("https://www.cmegroup.com" + m).text
    print("JS", m, len(js))
    for x in sorted(set(re.findall(r"[^\"'`\s(]{0,60}(?:CmeWS|mvc/)[^\"'`\s)]{0,140}", js))): print("  EP", x)
for x in sorted(set(re.findall(r"CmeWS[^\"'<\s]{0,120}|/services/[^\"'<\s]{0,100}", p)))[:20]: print("PG", x)
for tag in ("BZ", "CL"):
    for yr in range(26, 33):
        row = []
        for mc in "FGHJKMNQUVXZ":
            t = f"{tag}{mc}{yr}.NYM"
            r = g(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}?range=max&interval=1d")
            if r.status_code != 200: row.append(f"{mc}:-"); continue
            res = r.json()["chart"]["result"][0]; ts = res.get("timestamp") or []
            c = [b for b in res["indicators"]["quote"][0]["close"] if b is not None]
            row.append(f"{mc}:{len(c)}@{dt.date.fromtimestamp(ts[0]).isoformat()}" if ts else f"{mc}:0")
        print("YH", tag, yr, " ".join(row))
