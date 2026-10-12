"""
Probe (Oct 2026): which JODI Gas balance rows exist for Jordan and Morocco, and over what months.
Uses the jodidb.org viewer code of middle_east/EGYPT_GAS.py (Viewer.country_handle, parse_jodi); one download of every
balance row, every month, per country. Prints per row the first/last month with a value, number of months with a value,
the months with no value inside that span, and calendar-year sums in mcm. Run from GitHub Actions
(discovery_archive/workflows/jordan_morocco_jodi_probe.yml); output copied to discovery_archive/results/jordan_gas/.
"""
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "middle_east"))
import EGYPT_GAS as eg  # noqa: E402

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 400)

v = eg.Viewer()
for country in ("Jordan", "Morocco"):
    print(f"\n================ {country}")
    try:
        h, label = v.country_handle(country)
        print(f"handle {h} label {label!r}")
        body = v.download({"Country": [h], "Product": [0], "BALANCE": None, "Unit": [0], "TIME": None})
        d = eg.parse_jodi(body, country)
    except Exception as e:  # noqa: BLE001
        print(f"FAILED {type(e).__name__}: {e}")
        continue
    for c in d.columns:
        s = d[c].dropna()
        if not len(s):
            print(f"{c:55s} no values")
            continue
        span = pd.date_range(s.index.min(), s.index.max(), freq="MS")
        gaps = [m.strftime("%b%Y") for m in span if m not in s.index]
        nz = (s != 0).sum()
        print(f"{c:55s} {s.index.min():%b %Y} .. {s.index.max():%b %Y}  n={len(s)} nonzero={nz} gaps={len(gaps)} "
              f"{gaps[:30]}")
    y = d.groupby(d.index.year).sum(min_count=1).round(1)
    n = d.notna().groupby(d.index.year).sum()
    print("\nmonths with a value per year:\n", n.to_string())
    print("\nannual sums mcm:\n", y.to_string())
    print("\nlast 18 months:\n", d.tail(18).to_string())
