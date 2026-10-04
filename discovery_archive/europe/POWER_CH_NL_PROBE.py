"""
Probe: raw sources to fix the Swiss and Dutch power balances.
CH: BFE energiedashboard open data (Swissgrid production by type, national consumption) via opendata.swiss CKAN + direct ogd CSVs.
NL: CBS StatLine tables on renewable / total electricity production (solar), monthly if available.
Prints dataset resources, CSV heads and annual totals (TWh) for 2022-2025. Prints only.
"""
import io
import json
import re
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0"}


def j(url, **kw):
    r = requests.get(url, headers=H, timeout=90, **kw)
    return r


print("=== CH: opendata.swiss packages")
for pid in ("BFE-DS-0093@bundesamt-fur-energie-bfe", "BFE-DS-0096@bundesamt-fur-energie-bfe"):
    for url in (f"https://ckan.opendata.swiss/api/3/action/package_show?id={pid}", f"https://ckan.opendata.swiss/perma/{pid}"):
        try:
            r = j(url)
            print(pid, url[:60], r.status_code, r.headers.get("content-type", "")[:40])
            if "json" in r.headers.get("content-type", ""):
                res = r.json().get("result", {})
                print(" title:", res.get("title"), "| resources:")
                for x in res.get("resources", []):
                    print("   ", x.get("format"), x.get("url"))
                break
        except Exception as e:  # noqa: BLE001
            print(pid, "FAILED", type(e).__name__, str(e)[:120])
for url in ("https://www.uvek-gis.admin.ch/BFE/ogd/103/ogd103_stromverbrauch_swissgrid_lv_und_endv.csv",
            "https://www.uvek-gis.admin.ch/BFE/ogd/104/ogd104_stromproduktion_swissgrid.csv",
            "https://www.uvek-gis.admin.ch/BFE/ogd/102/ogd102_stromproduktion_swissgrid.csv"):
    try:
        r = j(url)
        print("\nCSV", url, r.status_code, len(r.content))
        if r.ok:
            t = r.text
            print(t[:600])
            df = pd.read_csv(io.StringIO(t))
            print(df.columns.tolist(), len(df), df.iloc[0, 0], df.iloc[-1, 0])
            dcol = df.columns[0]
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
            num = df.select_dtypes("number")
            print("annual sums (GWh if kWh/1e6):")
            print((num.groupby(df[dcol].dt.year).sum() / 1e6).round(1).loc[2022:2025].T.to_string())
    except Exception as e:  # noqa: BLE001
        print("CSV FAILED", url, type(e).__name__, str(e)[:150])

print("\n=== NL: CBS catalog")
for q in ("elektriciteit", "zonne", "hernieuwbare"):
    try:
        r = j("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$format": "json", "$filter": f"substringof('{q}',Title) eq true", "$select": "Identifier,Title,Period,Frequency,Status"})
        for t in r.json().get("value", []):
            if t.get("Status") != "Gediscontinueerd" and t.get("Frequency") in ("Permaand", "Perjaar", "Perkwartaal"):
                print(" ", t["Identifier"], "|", t["Title"][:90], "|", t.get("Period"), "|", t.get("Frequency"))
    except Exception as e:  # noqa: BLE001
        print("CBS catalog FAILED", q, type(e).__name__, str(e)[:120])
sys.exit(0)
