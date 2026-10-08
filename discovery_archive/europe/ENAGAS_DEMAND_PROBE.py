"""Probe (8 Oct 2026): what does Enagas' demand-history JSON return for recent query dates? The Spain pull got
'no rows from 2026-09-08' although the workbook already holds days to 22 Sep. Prints the keys, the number of
'actual' rows and the first/last fecha_demanda for several query dates. Read-only."""
import json
from datetime import date, timedelta

import requests

URL = ("https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/demanda/historico/jcr:content/responsiveGrid/"
       "container_copy_19796/realdemand_copy_copy.realdemand.json")
H = {"Accept": "application/json, text/javascript, */*; q=0.01", "User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest",
     "Referer": "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/"}
today = date.today()
for q in [today, today - timedelta(days=3), today - timedelta(days=7), date(2026, 9, 30), date(2026, 9, 22), date(2026, 8, 31)]:
    try:
        r = requests.get(URL, params={"date": q.strftime("%d/%m/%Y")}, headers=H, timeout=(15, 60))
        print(f"\nquery {q}: HTTP {r.status_code}, {len(r.content)} bytes, content-type {r.headers.get('content-type')}")
        try:
            j = r.json()
        except Exception as e:  # noqa: BLE001
            print("  not JSON:", r.text[:300]); continue
        print("  keys:", list(j.keys()) if isinstance(j, dict) else type(j))
        for k, v in (j.items() if isinstance(j, dict) else []):
            if isinstance(v, list):
                f = [e.get("fecha_demanda") for e in v if isinstance(e, dict) and e.get("fecha_demanda")]
                print(f"  {k}: {len(v)} rows", (f[0], f[-1]) if f else (v[:1] if v else ""))
            else:
                print(f"  {k}: {str(v)[:120]}")
    except Exception as e:  # noqa: BLE001
        print(f"query {q}: FAILED {type(e).__name__}: {e}")
