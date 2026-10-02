"""
Taiwan natural gas supply and consumption, monthly, from the Bureau of Energy's E-STAT open API (no key):
  https://ea01.moeaea.gov.tw/a0303/02/api/v1/zone/monthly/6/1   (table 6-01, thousand m3 = 10^3 m3)
Tables returned: natural gas total, indigenous, imported LNG (monthly and quarterly). Each is a list of records
{ "<title>": <row label>, "Column2": ..., ... "ColumnN": <period 'YYYY/MM'> }; the first rows are unit / header rows, then
one row per month, then year-on-year rows. The API only serves the latest ~3 years, so months accumulate in the stored
workbook (history store): stored months are kept, months in the response replace them.

Output sheets:
  Monthly      - date, natural gas supply / production / imports / uses by sector, kcm (10^3 m3) per month, plus mcm/d columns
  LNG monthly  - the imported-LNG table (same period handling), columns by position (see 'Headers')
  Headers      - header text found for every column of each table (so column meanings can be checked)
  Units        - units and notes

    python3 asia/TAIWAN_GAS.py --out "output/Data and Chart Outputs/taiwan_gas.xlsx"
"""
import argparse
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

LANDING = "https://ea01.moeaea.gov.tw/a0303/02/en/database/api/"
URL = "https://ea01.moeaea.gov.tw/a0303/02/api/v1/zone/monthly/6/1"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0 Safari/537.36",
     "Accept": "application/json, */*", "Accept-Language": "en-GB,en;q=0.9,zh-TW;q=0.8", "Referer": LANDING}
PERIOD = re.compile(r"^\d{4}/\d{2}$")
# natural gas total table: column position -> name (period is the last column); checked against 'Headers'
NG_COLUMNS = {2: "Supply", 3: "Indigenous_production", 4: "LNG_imports", 5: "Transformation_input", 6: "Refinery_input",
              7: "Power_and_cogen", 8: "Final_consumption", 9: "Energy_own_use", 10: "Industrial", 11: "Transport",
              12: "Agriculture", 13: "Services", 14: "Residential", 15: "Non_energy"}


def records_for(payload, name_contains, exclude="(季)"):
    for key, val in payload.items():
        if name_contains in key and exclude not in key and isinstance(val, list):
            return val
    raise KeyError(f"no list containing {name_contains!r}; keys {list(payload)}")


def split_table(records):
    """-> (monthly frame indexed by month start with ColumnN columns, header dict ColumnN -> text)."""
    if not records:
        return pd.DataFrame(), {}
    title = next(iter(records[0]))
    ncols = max(int(k[6:]) for r in records for k in r if k.startswith("Column") and k[6:].isdigit())
    pcol = f"Column{ncols}"
    rows, headers = [], {}
    for r in records:
        p = str(r.get(pcol, "")).strip()
        if PERIOD.match(p):
            d = {k: v for k, v in r.items() if k.startswith("Column") and k != pcol}
            d["date"] = pd.to_datetime(p + "/01", format="%Y/%m/%d")
            rows.append(d)
        elif not rows:   # header / unit rows come before the first month
            for k, v in r.items():
                if isinstance(v, str) and v.strip():
                    headers.setdefault(k if k != title else "Column1", []).append(v.strip().replace("\n", " "))
    df = pd.DataFrame(rows)
    if df.empty:
        return df, headers
    df = df.set_index("date").sort_index()
    return df.apply(pd.to_numeric, errors="coerce"), {k: " | ".join(dict.fromkeys(v)) for k, v in headers.items()}


def merge_history(old, new):
    if old is None or old.empty:
        return new
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def read_old(path, sheet):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=sheet)
        d["date"] = pd.to_datetime(d["date"])
        return d.set_index("date")
    except Exception as e:  # noqa: BLE001
        print(f"could not read stored sheet {sheet!r} ({type(e).__name__})")
        return pd.DataFrame()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/taiwan_gas.xlsx")
    args = ap.parse_args()
    r = requests.get(URL, headers=H, timeout=(15, 180))
    print(f"HTTP {r.status_code}, {len(r.content):,} bytes")
    r.raise_for_status()
    payload = r.json()
    ng, ng_head = split_table(records_for(payload, "天然氣"))
    lng, lng_head = split_table(records_for(payload, "進口"))
    if ng.empty:
        sys.exit("no monthly rows parsed from the natural-gas table")
    print(f"natural gas: {len(ng)} months {ng.index.min():%Y-%m} to {ng.index.max():%Y-%m}; LNG table {len(lng)} months")
    monthly = ng[[f"Column{i}" for i in NG_COLUMNS if f"Column{i}" in ng]].rename(
        columns={f"Column{i}": f"{n}_kcm" for i, n in NG_COLUMNS.items()})
    old = read_old(args.out, "Monthly")
    if not old.empty:
        old = old[[c for c in old.columns if c in monthly.columns]]   # keep the kcm columns; mcm/d is recomputed below
    monthly = merge_history(old, monthly)
    days = monthly.index.days_in_month
    for c in ["Supply", "Indigenous_production", "LNG_imports", "Power_and_cogen", "Final_consumption", "Industrial",
              "Services", "Residential"]:
        if f"{c}_kcm" in monthly:
            monthly[f"{c}_mcm_d"] = (monthly[f"{c}_kcm"] / 1000.0 / days).round(3)
    lng_all = merge_history(read_old(args.out, "LNG monthly"), lng)
    monthly.index.name = lng_all.index.name = "date"
    heads = pd.DataFrame({"table": ["natural gas"] * len(ng_head) + ["imported LNG"] * len(lng_head),
                          "column": list(ng_head) + list(lng_head),
                          "header text": list(ng_head.values()) + list(lng_head.values())})
    lines = ["UNITS", "kcm = thousand cubic metres (10^3 m3) per month; *_mcm_d = million m3 per day (kcm / 1000 / days in month).",
             "Column meanings: see 'Headers' (header text exactly as the API returns it).", "",
             "SOURCE", "Taiwan Bureau of Energy, Ministry of Economic Affairs: E-STAT open API, table 6-01",
             f"{URL}", "", "COVERAGE",
             "The API serves roughly the latest three years; earlier months accumulate in this workbook from the first run."]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Monthly": monthly, "LNG monthly": lng_all, "Headers": heads.set_index("table")},
                              lines, {"UNITS", "SOURCE", "COVERAGE"})
    print(f"Saved {args.out}: {len(monthly)} months {monthly.index.min():%Y-%m} to {monthly.index.max():%Y-%m}")


if __name__ == "__main__":
    main()
