"""Probe 4: PSE open data 2025: annual energy of his-wlk-cal fields, kse-load, przeplywy-mocy by section, his-gen-pal by ENTSO-E fuel alias (Jan-Mar), to compare
with ENTSO-E (generation, load, flows). PSE history starts 2024-06-14, so 2022-23 cannot be covered."""
import re
import sys
import time
import requests
import pandas as pd

UA = "Mozilla/5.0"
B = "https://api.raporty.pse.pl/api/"


def pull(ep, d0, d1):
    rows, url, params = [], B + ep, {"$filter": f"business_date ge '{d0}' and business_date le '{d1}'", "$first": 100000}
    n = 0
    while url:
        for i in range(4):
            try:
                r = requests.get(url, params=params, headers={"User-Agent": UA}, timeout=120)
                if r.ok:
                    break
            except requests.RequestException:
                pass
            time.sleep(5)
        j = r.json()
        v = j.get("value", [])
        rows += v
        n += 1
        url, params = j.get("nextLink"), None
    print(f"{ep} {d0}..{d1}: {len(rows)} rows in {n} pages", flush=True)
    return pd.DataFrame(rows)


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", ".").str.replace(r"\s", "", regex=True), errors="coerce")


for ep in ("his-wlk-cal", "kse-load", "przeplywy-mocy"):
    for (a, b) in (("2025-01-01", "2025-12-31"), ("2024-07-01", "2024-12-31")):
        d = pull(ep, a, b)
        if d.empty:
            continue
        print(ep, "columns", d.columns.tolist(), flush=True)
        if ep == "przeplywy-mocy":
            d["value"] = num(d["value"])
            print((d.groupby("section_code")["value"].sum() * 0.25 / 1e6).round(3).to_dict(), "TWh", flush=True)
        else:
            for c in d.columns:
                if c in ("dtime", "period", "dtime_utc", "period_utc", "business_date", "publication_ts", "publication_ts_utc"):
                    continue
                x = num(d[c])
                print(f"  {c}: TWh={x.sum() * 0.25 / 1e6:.3f} nonnull={x.notna().sum()}/{len(x)}", flush=True)
d = pull("his-gen-pal", "2025-01-01", "2025-03-31")
if not d.empty:
    d["v"] = num(d["value"])
    print((d.groupby("alias_entsoe")["v"].sum() * 0.25 / 1e6).round(3).to_dict(), "TWh Jan-Mar 2025", flush=True)
