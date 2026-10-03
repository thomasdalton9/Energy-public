"""One-off probe 2: EirGrid 'System and Renewable Data Reports' spreadsheets - sheet names, headers, date ranges, so we can
see whether they carry all-island generation by fuel (and how far back). Also checks the dashboard generationactual /
solar / interconnection series over a multi-day span. Prints only."""
import io
import re
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "*/*"}
page = requests.get("https://www.eirgrid.ie/grid/system-and-renewable-data-reports", headers=H, timeout=90).text
links = sorted(set(re.findall(r'href="([^"]+\.(?:xlsx|xls|csv)[^"]*)"', page)))
print("all links:", links, flush=True)
for u in links:
    try:
        r = requests.get(u, headers=H, timeout=180)
        print(f"\n=== {u}: {r.status_code} {len(r.content)} bytes", flush=True)
        x = pd.ExcelFile(io.BytesIO(r.content))
        print("sheets:", x.sheet_names[:40])
        for n in x.sheet_names[:12]:
            d = x.parse(n, header=None, nrows=14)
            print(f"--- {n} {d.shape}")
            print(d.iloc[:14, :10].to_string(max_colwidth=28)[:1800])
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u}: {type(e).__name__}: {e}")
sys.exit(0)
