"""
Round 5: Chile history. cne.cl keeps only the current Capacidad_Instalada_Generacion.xlsx
(round 2 found no earlier uploads at their old wp-content URLs), so the monthly
series before the first run is rebuilt from today's plant list and misses the
2021-2025 coal retirements. This asks the Internet Archive (Wayback CDX) whether
it holds earlier copies of CNE's capacity workbook, and opens the oldest and one
mid-period copy to compare the coal total.

    python3 POWER_CAPACITY_PROBE5.py

Found (Oct 2026): the Wayback CDX holds no copies of CNE's capacity workbooks
(empty result for all three patterns), so Chile's months before the first run
stay rebuilt from the current plant list.
"""
import io
import time

import pandas as pd
import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
CDX = "http://web.archive.org/cdx/search/cdx"
rows = []
for pattern in ["cne.cl/wp-content/uploads/*Capacidad*", "www.cne.cl/wp-content/uploads/*Capacidad*",
                "cne.cl/wp-content/uploads/*capacidad*"]:
    for attempt in range(4):
        try:
            r = requests.get(CDX, params={"url": pattern, "output": "json", "limit": 2000}, headers=UA,
                             timeout=(20, 120))
            print(pattern, r.status_code, len(r.content), flush=True)
            if r.ok:
                j = r.json()
                rows += j[1:] if j else []
                break
        except Exception as e:  # noqa: BLE001
            print(pattern, "attempt", attempt + 1, type(e).__name__, flush=True)
        time.sleep(15)
seen = {}
for row in rows:
    ts, orig, mime, status = row[1], row[2], row[3], row[4]
    seen[(ts, orig)] = (mime, status)
for (ts, orig), (mime, status) in sorted(seen.items()):
    print(ts, status, mime, orig)
xl = [(ts, orig) for (ts, orig), (mime, status) in sorted(seen.items())
      if status == "200" and orig.lower().endswith((".xlsx", ".xls"))]
for ts, orig in ([xl[0], xl[len(xl) // 2]] if len(xl) > 1 else xl):
    u = f"http://web.archive.org/web/{ts}id_/{orig}"
    try:
        r = requests.get(u, headers=UA, timeout=(20, 300))
        print("\nGET", u, r.status_code, len(r.content), r.content[:4], flush=True)
        x = pd.ExcelFile(io.BytesIO(r.content))
        print(x.sheet_names)
        for s in x.sheet_names[:3]:
            d = x.parse(s, header=None)
            print(s, d.shape)
            print(d.dropna(how="all").head(12).to_string(max_colwidth=25)[:3000])
    except Exception as e:  # noqa: BLE001
        print("open failed", u, e)
