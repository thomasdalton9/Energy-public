"""One-off probe: can GitHub Actions reach the ERCOT large-load interconnection queue / long-term load forecast pages and the IEA
'Energy and AI' page? Prints status, size and any sentence with GW / TWh so the Texas data-centre assumptions can be checked."""
import re
import requests

URLS = [
    "https://www.ercot.com/gridinfo/load/forecast",
    "https://www.ercot.com/services/rq/large-load-integration",
    "https://www.ercot.com/files/docs/2025/11/18/Large-Load-Interconnection-Status-Update.pdf",
    "https://www.iea.org/reports/energy-and-ai/energy-demand-from-ai",
    "https://www.iea.org/reports/energy-and-ai/executive-summary",
]
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=40)
        t = re.sub(r"<[^>]+>", " ", r.text) if "html" in r.headers.get("content-type", "") else ""
        t = re.sub(r"\s+", " ", t)
        print(f"\n== {u}\n   status {r.status_code}, {len(r.content)} bytes, {r.headers.get('content-type')}")
        hits = [s.strip() for s in re.split(r"(?<=[.])\s", t) if re.search(r"\b(GW|TWh|gigawatt)\b", s) and re.search(r"data cent|large load|AI", s, re.I)]
        for s in hits[:12]:
            print("   *", s[:300])
    except Exception as e:  # noqa: BLE001
        print(f"\n== {u}\n   FAILED {type(e).__name__}: {e}")
