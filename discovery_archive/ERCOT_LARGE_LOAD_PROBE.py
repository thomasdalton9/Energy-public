"""One-off probe: does ERCOT publish readable large-load / crypto / data-centre OBSERVED load numbers (GW) from GitHub Actions?
Fetches a few ERCOT pages, prints status, links that look like large-load / crypto / forecast files, and sentences with GW/MW near
'large load', 'crypto', 'data cent'. For the Load breakout override block in americas/TEXAS_DEMAND_REGRESSION.py."""
import re
import requests

URLS = [
    "https://www.ercot.com/services/rq/large-load-integration",
    "https://www.ercot.com/gridinfo/load/forecast",
    "https://www.ercot.com/gridinfo/load",
    "https://www.ercot.com/news/mediakit/factsheets",
    "https://www.ercot.com/gridmktinfo/dashboards",
    "https://www.ercot.com/mktinfo/loads",
]
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
KEY = re.compile(r"large.?load|crypto|data.?cent|bitcoin|LLIS|CLR|controllable", re.I)
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=40)
        print(f"\n== {u}\n   status {r.status_code}, {len(r.content)} bytes, {r.headers.get('content-type')}")
        links = sorted(set(re.findall(r'href="([^"]+)"', r.text)))
        keep = [l for l in links if (KEY.search(l) or re.search(r"\.(pdf|xlsx|xls|csv|zip)$", l, re.I)) and re.search(r"load|crypto|forecast|large|ldc|cdr", l, re.I)]
        for l in keep[:25]:
            print("   link:", l[:200])
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        hits = [s.strip() for s in re.split(r"(?<=[.])\s", t) if KEY.search(s) and re.search(r"\b\d[\d,.]*\s?(GW|MW|gigawatt)", s)]
        for s in hits[:10]:
            print("   *", s[:300])
    except Exception as e:  # noqa: BLE001
        print(f"\n== {u}\n   FAILED {type(e).__name__}: {e}")
