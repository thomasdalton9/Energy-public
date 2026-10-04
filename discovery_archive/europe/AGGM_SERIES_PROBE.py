"""
Probe: list AGGM data monitor time-series names (vis-service /ts/attributes) matching production / storage / border flow terms, then fetch
2023 and 2024 annual sums (GWh) of every match. Looks for Austrian domestic production, storage and border flows to close the Austria balance.
Prints only.
"""
import re
import sys

import requests

UA = "Mozilla/5.0"
B = "https://platform.aggm.at/vis-service/api/"
t = requests.get(B + "ts/attributes", headers={"User-Agent": UA}, timeout=90).text
names = sorted(set(re.findall(r'"name":"([^"]+)"', t)))
print("total names", len(names), flush=True)
pat = re.compile(r"produ|f[oö]rder|inland|speicher|storage|haidach|baumgarten|grenz|border|fluss|flow|einspeis|ausspeis|entry|exit|physik|saldo|import|export|VHP|abgabe|verbrauch|netz|verlust", re.I)
sel = [n for n in names if pat.search(n)]
print("matching", len(sel), flush=True)
for n in sel:
    print("  ", n)
sys.stdout.flush()
for yr in (2023, 2024):
    print("== sums", yr, flush=True)
    for i in range(0, len(sel), 20):
        chunk = sel[i:i + 20]
        body = {"rangeType": "individual", "from": f"{yr}-01-01T06:00:00", "to": f"{yr + 1}-01-01T06:00:00", "granularity": "day", "timeseries": chunk}
        try:
            r = requests.post(B + "ts/values", json=body, headers={"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}, timeout=120)
            for cd in r.json()["timeSeriesData"]["chartData"]:
                ys = [p["y"] for p in cd["dataSet"] if p.get("y") is not None]
                if ys:
                    print(f"   {cd['header']['name']} n={len(ys)} sumGWh={sum(ys) / 1e6:.0f} unit={cd['header'].get('unit')}", flush=True)
        except Exception as e:  # noqa: BLE001
            print("   chunk failed", type(e).__name__, str(e)[:100], flush=True)
