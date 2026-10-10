"""
Pull the weekly LNG stock of Japan's large power utilities from ANRE (Agency for Natural Resources and Energy, METI) and keep it
as a growing archive (output/Data and Chart Outputs/japan_lng_stock_weekly.xlsx).

Source: "発電用LNGの在庫状況" - a one-page PDF ANRE refreshes every week
(https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf): the week-end
stock of the large power companies (10,000 tonnes, excluding the unpumpable dead stock) for the weeks of the current
season, plus the previous year's and the 2021-25 average month-end stock. The PDF keeps no history beyond the
current season, so the committed workbook is the history store (a week-end already stored is replaced by the latest
value ANRE shows for it). Output in thousand tonnes (kt = 10,000 t x 10).

Sheet 'Weekly': date (week-end Sunday), Stock_kt. Sheet 'Month-end reference': month-end, Prior_year_kt, Average_2021_2025_kt
as printed in the PDF (the comparison basis changes yearly, so each value keeps the print date).

    python3 japan/JAPAN_LNG_STOCK.py --out "output/Data and Chart Outputs/japan_lng_stock_weekly.xlsx"
"""
import argparse
import os
import re
import sys
import time

import pandas as pd
import pymupdf
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "japan_lng_stock_weekly.xlsx")
ZEN = str.maketrans("０１２３４５６７８９", "0123456789")


def parse(text):
    """Token walk over the PDF text: 'M/D 時点' followed by a number is a week-end stock; 'N 月末' followed by two numbers is a
    month-end reference (prior year, 2021-25 average)."""
    t = text.translate(ZEN)
    m = re.search(r"令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", t)
    if not m:
        raise ValueError("publication date not found")
    pub = pd.Timestamp(2018 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
    toks = [x.strip() for x in t.splitlines() if x.strip()]
    weekly, refs = {}, {}
    num = re.compile(r"^\d{2,4}$")
    i = 0
    while i < len(toks):
        w = re.match(r"^(\d{1,2})/(\d{1,2})\s*時点$", toks[i])
        mo = re.match(r"^(\d{1,2})\s*月末$", toks[i])
        if w:
            mth, day = int(w.group(1)), int(w.group(2))
            year = pub.year if mth <= pub.month + 6 else pub.year - 1
            if i + 1 < len(toks) and num.match(toks[i + 1]):
                weekly[pd.Timestamp(year, mth, day)] = float(toks[i + 1]) * 10.0   # 10,000 t -> kt
        elif mo:
            mth = int(mo.group(1))
            if i + 2 < len(toks) and num.match(toks[i + 1]) and num.match(toks[i + 2]):
                year = pub.year if mth <= pub.month + 6 else pub.year - 1
                refs[pd.Timestamp(year, mth, 1) + pd.offsets.MonthEnd(0)] = (float(toks[i + 1]) * 10.0, float(toks[i + 2]) * 10.0)
        i += 1
    return pub, weekly, refs


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    last = None
    for i in range(3):
        try:
            r = requests.get(URL, headers=UA, timeout=(10, 90))
            r.raise_for_status()
            break
        except requests.RequestException as e:
            last = e
            time.sleep(5 * (i + 1))
    else:
        raise last
    doc = pymupdf.open(stream=r.content, filetype="pdf")
    pub, weekly, refs = parse("\n".join(p.get_text() for p in doc))
    if not weekly:
        raise SystemExit("no weekly values parsed")
    for d, v in weekly.items():
        if not 500 <= v <= 6000:
            raise SystemExit(f"stock {v} kt on {d:%Y-%m-%d} outside 500-6000 kt")
    old_w = old_r = None
    if os.path.exists(args.out):
        try:
            old_w = pd.read_excel(args.out, sheet_name="Weekly", index_col=0, parse_dates=True)
            old_r = pd.read_excel(args.out, sheet_name="Month-end reference", index_col=0, parse_dates=True)
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
    new_w = pd.DataFrame({"Stock_kt": pd.Series(weekly)})
    new_r = pd.DataFrame.from_dict({k: {"Prior_year_kt": a, "Average_2021_2025_kt": b, "Printed": pub} for k, (a, b) in refs.items()},
                                   orient="index")
    w = pd.concat([old_w, new_w]) if old_w is not None else new_w
    w = w[~w.index.duplicated(keep="last")].sort_index()
    rr = pd.concat([old_r, new_r]) if old_r is not None and len(new_r) else (new_r if len(new_r) else old_r)
    if rr is not None:
        rr = rr[~rr.index.duplicated(keep="last")].sort_index()
    w.index.name = "date"
    if rr is not None:
        rr.index.name = "month_end"
    sheets = {"Weekly": w}
    if rr is not None:
        sheets["Month-end reference"] = rr
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES, TITLES)
    print(f"Saved {len(w)} weeks ({w.index.min():%Y-%m-%d}..{w.index.max():%Y-%m-%d}); latest {w['Stock_kt'].iloc[-1]:.0f} kt "
          f"(PDF dated {pub:%Y-%m-%d}) -> {args.out}")


NOTES = [
    "UNITS",
    "Stock_kt: week-end LNG stock of the large power companies, thousand tonnes (ANRE prints 10,000 tonnes; x10), excluding the "
    "dead stock that cannot be pumped out. Reference sheet: previous-year and 2021-25 average month-end stock as printed.",
    "",
    "SOURCE",
    "Agency for Natural Resources and Energy, METI: 発電用LNGの在庫状況 (weekly PDF), "
    "https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf .",
    "",
    "UPDATES",
    "The PDF shows only the current season, so this workbook is the history: each run adds the weeks shown and replaces a "
    "stored week with ANRE's latest value for it. History starts at the first run (the PDF's earliest week then).",
]
TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
