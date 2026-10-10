"""
South Korea natural gas (LNG) imports by origin region, monthly from Jan 1988: Korea Gas Corporation (KOGAS) file dataset
'Korea's natural gas imports by continent' (한국가스공사_한국의 대륙별 천연가스 수입 현황) on the Public Data Portal
(data.go.kr dataset 15088508, CSV, free download without login). Weight in tonnes, value in million US$ and unit price in
US$ per tonne for the total and for Oceania, North America, Latin America, Middle East, Asia, Africa, Europe, Russia, Other.

The portal page gives the current file's name (which carries the as-of date) and its download id; the file is downloaded only
when the name differs from the one recorded on the 'Release' sheet. Probe: discovery_archive/south_korea/KOREA_SOURCES_PROBE14.py.

    python3 south_korea/SOUTH_KOREA_LNG_IMPORTS.py [--out "output/Data and Chart Outputs/south_korea_lng_imports_monthly.xlsx"]
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

PAGE = "https://www.data.go.kr/data/15088508/fileData.do"
DL = "https://www.data.go.kr/cmm/cmm/fileDownload.do?atchFileId={fid}&fileDetailSn=1"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "ko-KR,ko;q=0.9"}
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "Data and Chart Outputs",
                           "south_korea_lng_imports_monthly.xlsx")
REGIONS = ["Total", "Oceania", "North_America", "Latin_America", "Middle_East", "Asia", "Africa", "Europe", "Russia", "Other"]
REVISION_MONTHS = 3


def get(url):
    last = None
    for i in range(4):
        try:
            r = requests.get(url, headers=UA, timeout=(15, 120))
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"  attempt {i + 1}/4 {url[:80]}: {type(e).__name__}", file=sys.stderr)
            time.sleep(5 * (i + 1))
    raise last


def release_info():
    html = get(PAGE).text
    name = re.search(r"파일데이터명[^<]*</[^>]+>\s*<[^>]+>\s*([^<]{3,120})", html)
    fid = re.search(r"atchFileId=(\w+)", html)
    if not fid:
        raise RuntimeError("no download id found on the data.go.kr dataset page")
    return (name.group(1).strip() if name else ""), fid.group(1)


def parse(raw):
    d = pd.read_csv(io.BytesIO(raw), encoding="cp949")
    if d.shape[1] != 1 + 3 * len(REGIONS):
        raise RuntimeError(f"unexpected column count {d.shape[1]}: {list(d.columns)[:6]}")
    out = {}
    for i, reg in enumerate(REGIONS):
        out[f"{reg}_Price_USD_per_t"] = pd.to_numeric(d.iloc[:, 1 + 3 * i], errors="coerce")
        out[f"{reg}_Value_USDm"] = pd.to_numeric(d.iloc[:, 2 + 3 * i], errors="coerce")
        out[f"{reg}_Weight_t"] = pd.to_numeric(d.iloc[:, 3 + 3 * i], errors="coerce")
    df = pd.DataFrame(out)
    df.index = pd.to_datetime(d.iloc[:, 0].astype(str).str.strip() + "-01")
    df.index.name = "month"
    return df.sort_index()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    old = rel_old = None
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Data", index_col=0)
            old.index = pd.to_datetime(old.index)
            rel_old = pd.read_excel(args.out, sheet_name="Release")
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
            old = rel_old = None
    name, fid = release_info()
    print(f"portal file: {name} ({fid})", flush=True)
    if old is not None and rel_old is not None and len(rel_old) and str(rel_old["file_name"].iloc[0]) == name:
        print("same file as the stored release; nothing to download.", flush=True)
        return
    new = parse(get(DL.format(fid=fid)).content)
    if old is not None:
        cut = new.index.max() - pd.DateOffset(months=REVISION_MONTHS)
        keep = old[old.index < cut]
        new = pd.concat([keep, new[(new.index >= cut) | ~new.index.isin(keep.index)]]).sort_index()
    tot = new["Total_Weight_t"]
    parts = new[[f"{r}_Weight_t" for r in REGIONS[1:]]].sum(axis=1)
    gap = ((parts - tot).abs() / tot).max()
    last = new.index.max()
    rel = pd.DataFrame({"file_name": [name], "download_id": [fid], "last_month": [last.strftime("%Y-%m")]})
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": new, "Release": rel}, [
        "UNITS",
        "Data: monthly natural gas (LNG) imports of South Korea by origin region: <Region>_Weight_t in tonnes, <Region>_Value_USDm in "
        "million US$, <Region>_Price_USD_per_t in US$ per tonne (value / weight as published). Regions: Oceania, North America, "
        "Latin America, Middle East, Asia, Africa, Europe, Russia, Other; Total is the file's own total (the regions add up to it "
        f"within {gap * 100:.2f}% in every month). Months are dated the 1st. No conversion to energy or volume units is made.",
        "",
        "SOURCE",
        "Korea Gas Corporation (KOGAS), 'Korea's natural gas imports by continent', Public Data Portal (data.go.kr) file dataset "
        "15088508, CSV, free download without login: https://www.data.go.kr/data/15088508/fileData.do",
        "",
        "UPDATES",
        "The portal page shows the current file name (it carries the as-of date, see the Release sheet); the file is downloaded only "
        f"when the name differs from the recorded one. Stored months older than {REVISION_MONTHS} months are kept, the rest "
        "re-read. The file lags by several months (latest month " + last.strftime("%b %Y") + ").",
    ], {"UNITS", "SOURCE", "UPDATES"})
    print(f"Saved {args.out}: {len(new)} months {new.index.min():%Y-%m}..{last:%Y-%m}; 2025 total "
          f"{new[new.index.year == 2025]['Total_Weight_t'].sum() / 1e6:.2f} Mt; region sum vs total max gap {gap * 100:.3f}%")


if __name__ == "__main__":
    main()
