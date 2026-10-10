"""
Pull the yen per US dollar exchange rate (daily) and keep it as a growing archive
(output/Data and Chart Outputs/japan_fx_usd_daily.xlsx, sheet 'Data': date index + JPY_per_USD).

Source (primary): Board of Governors of the Federal Reserve System, H.10 release, "Historical Rates for the Japanese
Yen" (https://www.federalreserve.gov/releases/h10/hist/dat00_ja.htm): noon buying rates in New York for cable
transfers, yen per US dollar, one row per business day since 2000 - the series that FRED republishes as DEXJPUS.
FRED itself timed out from GitHub Actions (China twin, probe 7-8 Oct 2026), the Fed page
answers. Days the Fed marks ND (no data: US holidays) are not rows. Nothing is filled here: the master applies
"the last published rate on or before the price date".

Cross-check column JPY_per_USD_ECB_cross = ECB euro reference rate JPY per EUR / USD per EUR
(https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv, 2.15 pm CET fixing). Validation only; not used by the master.

Incremental: the committed workbook is the history store. Each run downloads the one Fed page (it holds the whole
history), keeps stored rows older than a 30-day revision window and takes the page's rows from there on.
History kept from 2016-01-01 (the JEPX price series used start in 2016).

    python3 japan/JAPAN_FX_USD.py --out "output/Data and Chart Outputs/japan_fx_usd_daily.xlsx"
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

H10 = "https://www.federalreserve.gov/releases/h10/hist/dat00_ja.htm"
ECB = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
START = pd.Timestamp("2016-01-01")
REVISION_DAYS = 30
COL, COL_ECB = "JPY_per_USD", "JPY_per_USD_ECB_cross"
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "japan_fx_usd_daily.xlsx")
ROW = re.compile(r"<th[^>]*>\s*(\d{1,2}-[A-Za-z]{3}-\d{2})\s*</th>\s*<td[^>]*>\s*([0-9.]+|ND|NA)\s*</td>", re.S)


def get(url):
    last = None
    for i in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=(10, 90))
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"  attempt {i + 1}/3 {url}: {type(e).__name__}", file=sys.stderr)
            time.sleep(5 * (i + 1))
    raise last


def fed_h10():
    html = get(H10).text
    rows = {}
    for d, v in ROW.findall(html):
        if v in ("ND", "NA"):
            continue
        rows[pd.to_datetime(d, format="%d-%b-%y")] = float(v)
    s = pd.Series(rows, name=COL).sort_index()
    if s.empty:
        raise RuntimeError("no rows parsed from the H.10 page")
    return s[s.index >= START]


def ecb_cross():
    df = pd.read_csv(io.StringIO(get(ECB).text), usecols=["Date", "USD", "JPY"], na_values=["N/A"])
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.dropna().set_index("Date").sort_index()
    s = (df["JPY"] / df["USD"]).rename(COL_ECB)
    return s[s.index >= START]


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    old = None
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Data", index_col=0)
            old.index = pd.to_datetime(old.index)
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
    try:
        fed = fed_h10()
    except Exception as e:  # noqa: BLE001
        if old is None:
            raise
        print(f"  Fed H.10 failed ({type(e).__name__}: {e}); keeping the stored history", file=sys.stderr)
        fed = old[COL].dropna()
    try:
        ecb = ecb_cross()
    except Exception as e:  # noqa: BLE001
        print(f"  ECB failed ({type(e).__name__}); cross-check column not refreshed", file=sys.stderr)
        ecb = old[COL_ECB].dropna() if old is not None and COL_ECB in old else pd.Series(dtype=float, name=COL_ECB)
    if old is not None and COL in old:
        cut = fed.index.max() - pd.Timedelta(days=REVISION_DAYS)
        keep = old[COL].dropna()
        fed = pd.concat([keep[keep.index < cut], fed[fed.index >= cut]]).sort_index()
    data = pd.concat([fed.rename(COL), ecb.rename(COL_ECB)], axis=1).sort_index()
    data = data[data[COL].notna() | data[COL_ECB].notna()]
    both = data.dropna()
    diff = ((both[COL_ECB] / both[COL] - 1) * 100).abs()
    data.index.name = "date"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": data}, NOTES, TITLES)
    print(f"Saved {len(data)} rows ({data.index.min():%Y-%m-%d}..{data.index.max():%Y-%m-%d}); latest H.10 "
          f"{fed.index.max():%Y-%m-%d} = {fed.iloc[-1]}; ECB cross vs H.10 on shared days: mean abs diff "
          f"{diff.mean():.2f}%, max {diff.max():.2f}% over {len(both)} days -> {args.out}")


NOTES = [
    "UNITS",
    "JPY_per_USD: yen (yen) per one US dollar. JPY_per_USD_ECB_cross: the same from the ECB euro reference rates "
    "(JPY per EUR divided by USD per EUR), validation only.",
    "",
    "SOURCE",
    "Board of Governors of the Federal Reserve System, H.10 Foreign Exchange Rates, Historical Rates for the Japanese "
    "Yen (noon buying rates in New York for cable transfers, the series FRED publishes as DEXJPUS): "
    "https://www.federalreserve.gov/releases/h10/hist/dat00_ja.htm . "
    "ECB euro foreign exchange reference rates: https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv .",
    "",
    "DATES",
    "One row per day the source publishes (business days). Days the Fed marks ND (US holidays) have no row and nothing is "
    "filled here; users of the series apply the last published rate on or before the date they need.",
    "",
    "UPDATES",
    "Incremental: the stored history before the last 30 days is kept; the Fed page (whole history in one file) "
    "is read each run and its last 30 days replace the stored ones. FRED was tried first and timed out from GitHub Actions.",
]
TITLES = {"UNITS", "SOURCE", "DATES", "UPDATES"}

if __name__ == "__main__":
    main()
