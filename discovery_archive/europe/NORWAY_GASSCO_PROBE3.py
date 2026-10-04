"""Probe 3: gassco.eu realTimeGraphData with year/month/day (history?), deliveryNumbers by year/month (how far back),
umm.gassco.no after accepting the disclaimer (requests session). Prints only."""
import re
import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
G = "https://gassco.eu/wp-json/gassco/v1/"


def g(path, **params):
    try:
        r = requests.get(G + path, params=params, timeout=40, headers=H)
        print(f"{path} {params} -> {r.status_code} {len(r.text)} {r.text[:700]!r}")
    except Exception as e:  # noqa: BLE001
        print(path, params, type(e).__name__)


for p in ({"year": 2026, "month": 9, "day": 15}, {"year": 2026, "month": 9, "day": 28}, {"year": 2025, "month": 1, "day": 15},
          {"year": 2022, "month": 6, "day": 1}, {"year": 2026, "month": 9}, {"year": 2026}, {"year": 2020, "month": 1, "day": 1}):
    g("realTimeGraphData", **p)
for y in (2026, 2025, 2024, 2020, 2015, 2010, 2005, 2000):
    g("deliveryNumbers", year=y)
g("deliveryNumbers", year=2025, month=9)
g("deliveryNumbersSetup", year=2025)
g("deliveryNumbersSetup", year=2026)

s = requests.Session()
s.headers.update(H)
for u in ("https://umm.gassco.no/disclaimer/acceptDisclaimer", "https://umm.gassco.no/", "https://umm.gassco.no/index", "https://umm.gassco.no/umm",
          "https://umm.gassco.no/history", "https://umm.gassco.no/flow", "https://umm.gassco.no/flows", "https://umm.gassco.no/data", "https://umm.gassco.no/download"):
    try:
        r = s.get(u, timeout=40, allow_redirects=True)
        print("UMM", u, r.status_code, r.url, len(r.text), s.cookies.get_dict())
        if r.ok and len(r.text) > 600:
            body = re.sub(r"\s+", " ", r.text)
            print("   ", body[:600])
            for m in sorted(set(re.findall(r'(?:href|src|action)="([^"]+)"', r.text))):
                if not re.search(r"\.(css|png|gif|ico|woff)", m):
                    print("     ref", m[:200])
            for m in re.findall(r"(?:url|\.get|\.post|\.ajax|fetch)\s*[:(]\s*['\"]([^'\"]+)['\"]", r.text)[:30]:
                print("     js", m)
    except Exception as e:  # noqa: BLE001
        print("UMM", u, type(e).__name__)
