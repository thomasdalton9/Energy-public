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
import csv, io  # noqa: E402
listing = v.download({"Country": None, "Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})
names = [r[0] for r in csv.reader(io.StringIO(listing)) if len(r) > 2 and r[0] not in ("", "Country")]
print(f"JODI gas viewer country list ({len(names)}):", names)
print("Jordan-like:", [n for n in names if "ord" in n.lower() or "jor" in n.lower()])
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
    d.columns = [c.replace("_JODI_fallback_mcm", "") for c in d.columns]
    y = d.groupby(d.index.year).sum(min_count=1).round(1)
    n = d.notna().groupby(d.index.year).sum()
    print("\nmonths with a value per year:\n", n.to_string())
    print("\nannual sums mcm:\n", y.to_string())
    print("\nlast 18 months:\n", d.tail(18).to_string())


# Pass 2: is Jordan in JODI's static world gas file (GAS_world_NewFormat.zip, stale but complete list of reporters)?
import requests, zipfile  # noqa: E402
for url in ("https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv_beta.zip",
            "https://www.jodidata.org/_resources/files/downloads/gas-data/GAS_world_NewFormat.zip"):
    try:
        r = requests.get(url, headers=eg.UA, timeout=300)
        print(f"\n{url}: HTTP {r.status_code} {len(r.content)} bytes")
        if r.status_code != 200:
            continue
        z = zipfile.ZipFile(io.BytesIO(r.content))
        for n in z.namelist():
            txt = z.read(n).decode("utf-8", "replace")
            lines = txt.splitlines()
            print(n, len(lines), "lines; header:", lines[0][:200])
            areas = sorted({ln.split(",")[0] for ln in lines[1:]})
            print("REF_AREA codes:", len(areas), areas)
            jo = [ln for ln in lines if ln.startswith("JO,") or ",JO," in ln[:10]]
            ma = [ln for ln in lines if ln.startswith("MA,")]
            print("Jordan (JO) rows:", len(jo), jo[:5])
            print("Morocco (MA) rows:", len(ma), "months", sorted({ln.split(',')[1] for ln in ma})[:1],
                  sorted({ln.split(',')[1] for ln in ma})[-1:])
        break
    except Exception as e:  # noqa: BLE001
        print(url, type(e).__name__, e)
