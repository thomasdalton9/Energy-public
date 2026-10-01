"""
Singapore discovery, round 7: is there a MONTHLY natural gas import series?
SingStat T010002 'Merchandise Trade Volume by Commodity and Market, Monthly' (Enterprise Singapore trade
data) may carry natural gas (SITC 343 / HS 2711) import volumes. Probe its rows for gas, plus the monthly
merchandise trade tables on data.gov.sg. Also prints SES tidy T3.2 / T3.7 / T2.1 category names.
"""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json, */*"}
T = (10, 180)
SS = "https://tablebuilder.singstat.gov.sg/api/table"


def out(*a):
    print(*a, flush=True)


def singstat():
    for tid in ["T010002", "T010001"]:
        for params in [{"search": "gas", "limit": 200}, {"search": "Natural Gas", "limit": 200},
                       {"limit": 50}]:
            try:
                r = requests.get(f"{SS}/tabledata/{tid}", params=params, headers=H, timeout=T)
                out(f"GET {r.url} -> {r.status_code} {len(r.content)}")
                d = r.json().get("Data") or {}
            except Exception as e:
                out(f"  {tid} {params}: {e}")
                continue
            out(f"  {d.get('title')} | {d.get('frequency')} | rows {d.get('recordsTotal', '')} {len(d.get('row', []))}")
            for row in d.get("row", [])[:60]:
                cols = row.get("columns") or []
                if params.get("search") or len(d.get("row", [])) < 60:
                    out(f"   {row.get('seriesNo')} {row.get('rowText')!r} [{row.get('uoM')}] n={len(cols)} "
                        f"first={cols[:1]} last={cols[-2:]}")
            if not params.get("search"):
                out(f"   first rowTexts: {[x.get('rowText') for x in d.get('row', [])][:50]}")
    try:
        r = requests.get(f"{SS}/metadata/T010002", headers=H, timeout=T)
        out(f"metadata T010002: {r.text[:3000]}")
    except Exception as e:
        out(f"metadata: {e}")


def ses():
    url = ("https://www.ema.gov.sg/content/dam/corporate/resources/singapore-energy-statistics/excel/"
           "SES_tidy.xlsx.coredownload.xlsx")
    r = requests.get(url, headers=H, timeout=T)
    xl = pd.ExcelFile(io.BytesIO(r.content))
    for s in ["T2.1", "T3.2", "T3.7", "T1.1", "T2.2"]:
        df = pd.read_excel(xl, sheet_name=s)
        out(f"\n{s} columns {list(df.columns)}")
        for c in df.columns[1:-1]:
            out(f"  {c}: {list(dict.fromkeys(map(str, df[c].dropna())))}")
        out(df.tail(12).to_string()[:2500])


if __name__ == "__main__":
    for f in (singstat, ses):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__}: {e!r}")
