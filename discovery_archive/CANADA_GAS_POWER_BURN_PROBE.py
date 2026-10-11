"""
Canada gas for electric power: does StatCan 25-10-0086-01 (successor of 25-10-0055-01) or any other raw source carry
it?  The saved canada_gas.xlsx has residential, commercial, industrial (transmission/distribution split), plant
deliveries and pipeline fuel but no electric-power item, so power burn is probably inside 'Industrial consumption'.
Prints every member of every dimension in 25-10-0086-01 and the old 25-10-0055-01 (flagging electric/thermal/power),
then probes CER (Canada Energy Regulator) and NRCan pages for a gas-for-power series. Runs in GitHub Actions only
(the editing sandbox blocks these hosts).
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
FLAG = re.compile(r"electric|thermal|power|generat|utilit", re.I)


def out(*a):
    print(*a, flush=True)


for pid in ("25100086", "25100055"):
    out(f"=========== table {pid}")
    try:
        r = requests.get(f"https://www150.statcan.gc.ca/n1/tbl/csv/{pid}-eng.zip", headers=H, timeout=(10, 300))
        z = zipfile.ZipFile(io.BytesIO(r.content))
        d = pd.read_csv(z.open(next(n for n in z.namelist() if n.endswith(".csv") and "MetaData" not in n)),
                        dtype=str, encoding="utf-8-sig")
        out("columns", list(d.columns), "rows", len(d), "period", d["REF_DATE"].min(), d["REF_DATE"].max())
        for c in d.columns[d.columns.get_loc("DGUID") + 1:d.columns.get_loc("UOM")]:
            for v in d[c].dropna().unique():
                out(f"  {c}: {v}" + ("   <-- FLAG" if FLAG.search(v) else ""))
    except Exception as e:  # noqa: BLE001
        out("  failed", type(e).__name__, str(e)[:200])

for u in ("https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/natural-gas/index.html",
          "https://www.cer-rec.gc.ca/en/data-analysis/canada-energy-future/index.html",
          "https://open.canada.ca/data/en/dataset?q=natural+gas+electricity+generation+consumption",
          "https://natural-resources.canada.ca/energy-facts/natural-gas-facts/20066"):
    try:
        r = requests.get(u, headers=H, timeout=(10, 60))
        out(f"GET {u} -> {r.status_code} {len(r.content)}B")
        for m in sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx?|zip))"', r.text)))[:40]:
            out("    file:", m)
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} -> ERR {type(e).__name__}")
