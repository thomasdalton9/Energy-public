"""
One-off discovery for the North America master (capacity + Canada gas successor table):
  1. StatCan: active cubes whose title mentions natural gas / generating capacity / electric power
  2. EIA API v2 electricity/operating-generator-capacity: route metadata (facets, data columns), one month's
     sample rows and the row count
  3. EIA-860M monthly generator workbook URLs (current + archive) and the Operating sheet's header
  4. Ember yearly release: capacity rows for Mexico / Canada / United States
"""
import io
import os

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def section(t):
    print("\n" + "=" * 20, t, "=" * 20, flush=True)


section("StatCan cubes")
try:
    cubes = requests.get("https://www150.statcan.gc.ca/t1/wds/rest/getAllCubesListLite", headers=H, timeout=120).json()
    for c in cubes:
        t = c.get("cubeTitleEn", "")
        if any(k in t.lower() for k in ("natural gas", "generating capacity", "electric power", "electricity")):
            print(c.get("productId"), "|", t, "|", c.get("cubeStartDate"), "->", c.get("cubeEndDate"),
                  "| archived" if str(c.get("archived")) == "1" else "", "| freq", c.get("frequencyCode"))
except Exception as e:  # noqa: BLE001
    print("StatCan list failed:", e)

section("EIA operating-generator-capacity")
key = os.environ.get("EIA_API_KEY")
base = "https://api.eia.gov/v2/electricity/operating-generator-capacity/"
try:
    meta = requests.get(base, params={"api_key": key}, timeout=60).json()["response"]
    print("frequency:", meta.get("frequency"))
    print("facets:", [f.get("id") for f in meta.get("facets", [])])
    print("data:", list(meta.get("data", {}).keys()))
    print("start/end:", meta.get("startPeriod"), meta.get("endPeriod"))
    for fid in ("energy_source_code", "status", "technology"):
        try:
            fv = requests.get(base + f"facet/{fid}", params={"api_key": key}, timeout=60).json()["response"]
            print(fid, [(x.get("id"), x.get("name")) for x in fv.get("facets", [])][:80])
        except Exception as e:  # noqa: BLE001
            print(fid, "facet failed", e)
    end = meta.get("endPeriod")
    r = requests.get(base + "data/", params={"api_key": key, "frequency": "monthly", "data[0]": "net-summer-capacity-mw",
                                             "data[1]": "nameplate-capacity-mw", "start": end, "end": end,
                                             "length": 5}, timeout=120).json()["response"]
    print("rows in latest month:", r.get("total"))
    for row in r.get("data", []):
        print(row)
except Exception as e:  # noqa: BLE001
    print("EIA API failed:", type(e).__name__, e)

section("EIA-860M workbooks")
for url in ["https://www.eia.gov/electricity/data/eia860m/xls/july_generator2026.xlsx",
            "https://www.eia.gov/electricity/data/eia860m/xls/august_generator2026.xlsx",
            "https://www.eia.gov/electricity/data/eia860m/archive/xls/january_generator2021.xlsx",
            "https://www.eia.gov/electricity/data/eia860m/archive/xls/december_generator2024.xlsx"]:
    try:
        r = requests.get(url, headers=H, timeout=120)
        print(url, r.status_code, len(r.content))
        if r.ok:
            x = pd.ExcelFile(io.BytesIO(r.content))
            print("  sheets:", x.sheet_names)
            d = pd.read_excel(x, sheet_name=x.sheet_names[0], header=None, nrows=4)
            print(d.to_string()[:1500])
    except Exception as e:  # noqa: BLE001
        print(url, "failed", e)

section("Ember yearly capacity")
try:
    u = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv"
    d = pd.read_csv(io.BytesIO(requests.get(u, timeout=300).content))
    print(d.columns.tolist())
    print(d["Category"].unique())
    c = d[(d["Area"].isin(["Mexico", "Canada", "United States of America", "United States"])) & (d["Category"] == "Capacity")]
    print(c[["Area", "Year", "Subcategory", "Variable", "Unit", "Value"]].query("Year >= 2023").to_string()[:5000])
except Exception as e:  # noqa: BLE001
    print("Ember failed", e)
