"""
Probe: the Austria-Slovakia gas border (Baumgarten). Part 1: every ENTSOG point/direction of Eustream (SK-TSO-0001), of AT TSOs and of any 'Baumgarten' label,
annual TWh for Physical Flow / Allocation / Renomination. Part 2: AGGM data-monitor series names for border points (Entry/Exit/Grenz/Baumgarten ...) with 2025 sums.
Part 3: reachability of Eustream transparency pages (browser headers) and the links found on them. Prints only.
"""
import re
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
BR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}


def get(path, params, tries=3):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(8 * (i + 1))
        except requests.RequestException:
            time.sleep(8 * (i + 1))
    return {}


print("== PART 1 ENTSOG", flush=True)
ics = get("interconnections", {"limit": -1}).get("interconnections", [])
cand = {}
for i in ics:
    lab = str(i.get("pointLabel") or "")
    for opk, d, adjc, adjo, adjl in ((i.get("toOperatorKey"), "entry", i.get("fromCountryKey"), i.get("fromOperatorKey"), i.get("fromOperatorLabel")),
                                    (i.get("fromOperatorKey"), "exit", i.get("toCountryKey"), i.get("toOperatorKey"), i.get("toOperatorLabel"))):
        if not opk:
            continue
        if opk.startswith("SK-") or re.search("baumgarten|lanzhot|veltrusy|velke kapusany|mosonmagyar|vel.ke zlievce|budince", lab, re.I) or \
                (opk.startswith("AT-") and adjc in ("SK", "CZ", "HU")):
            cand[(opk, i["pointKey"], d)] = (lab, adjc, adjo, adjl)
print("candidates", len(cand), flush=True)
for (opk, pk, d), (lab, adjc, adjo, adjl) in sorted(cand.items()):
    for ind in ("Physical Flow", "Allocation", "Renomination"):
        tot = []
        for y in (2022, 2023, 2024, 2025):
            rows = get("operationalData", {"indicator": ind, "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                           "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
            tot.append(sum(float(r["value"]) for r in rows if r.get("value") not in (None, "")) / 1e9)
        if any(abs(t) > 0.05 for t in tot):
            print(" | ".join([ind[:5], opk, d, lab[:40], pk, str(adjc), str(adjo)] + [f"{t:.1f}" for t in tot]), flush=True)

print("== PART 2 AGGM", flush=True)
try:
    B = "https://platform.aggm.at/vis-service/api/"
    t = requests.get(B + "ts/attributes", headers={"User-Agent": "Mozilla/5.0"}, timeout=90).text
    names = sorted(set(re.findall(r'"name":"([^"]+)"', t)))
    pat = re.compile(r"entry|exit|grenz|baumgarten|oberkappel|ueberackern|[uü]berackern|arnoldstein|murfeld|mosonmagyar|reintal|laa|wallbach|slowak|ungarn|deutschland|italien|sloweni|tschech|VHP|nominier|allok|punkt|MGP|GCA|TAG|BOG|HAG|SOL|PRI|Netto", re.I)
    sel = [n for n in names if pat.search(n)]
    print("names", len(names), "selected", len(sel), flush=True)
    for yr in (2025,):
        for i in range(0, len(sel), 25):
            chunk = sel[i:i + 25]
            body = {"rangeType": "individual", "from": f"{yr}-01-01T06:00:00", "to": f"{yr + 1}-01-01T06:00:00", "granularity": "day", "timeseries": chunk}
            try:
                r = requests.post(B + "ts/values", json=body, headers={"User-Agent": "Mozilla/5.0", "Content-Type": "application/json", "Accept": "application/json"}, timeout=120)
                for cd in r.json()["timeSeriesData"]["chartData"]:
                    ys = [p["y"] for p in cd["dataSet"] if p.get("y") is not None]
                    if ys and sum(abs(v) for v in ys) > 0:
                        print(f"   {cd['header']['name']} n={len(ys)} sumGWh={sum(ys) / 1e6:.0f} unit={cd['header'].get('unit')}", flush=True)
            except Exception as e:  # noqa: BLE001
                print("   chunk failed", type(e).__name__, str(e)[:100], flush=True)
except Exception as e:  # noqa: BLE001
    print("AGGM failed", type(e).__name__, str(e)[:150], flush=True)

print("== PART 3 Eustream / Slovak sources", flush=True)
urls = ["https://www.eustream.sk/en/transparency", "https://www.eustream.sk/en/transparency/flow-data", "https://www.eustream.sk/en/transparency/data-and-publications",
        "https://www.eustream.sk/en/transparency/business-operational-data/", "https://data.eustream.sk/", "https://transparency.eustream.sk/en/",
        "https://www.eustream.sk/en/transparency/transparency-data", "https://www.eustream.sk/en",
        "https://datacube.statistics.sk/", "https://www.spp-distribucia.sk/en/", "https://www.mhsr.sk/", "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm?format=JSON&geo=SK&lang=en&siec=G3000&nrg_bal=IC_OBS&unit=TJ_GCV&time=2025M01"]
for u in urls:
    try:
        r = requests.get(u, headers=BR, timeout=40, allow_redirects=True)
        print(r.status_code, u, r.headers.get("content-type", "")[:40], len(r.content), flush=True)
        if r.ok and "html" in r.headers.get("content-type", ""):
            links = sorted(set(re.findall(r'href="([^"]+)"', r.text)))
            keep = [l for l in links if re.search(r"xls|csv|api|flow|data|transpar|download|baumgarten|nominat|allocat", l, re.I)]
            for l in keep[:40]:
                print("     ", l[:140], flush=True)
    except Exception as e:  # noqa: BLE001
        print("ERR", u, type(e).__name__, str(e)[:100], flush=True)
sys.exit(0)
