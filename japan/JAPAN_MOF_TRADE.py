"""
Pull Japan's monthly energy imports from the Ministry of Finance's trade statistics press releases and keep them as a
growing archive (output/Data and Chart Outputs/japan_mof_energy_imports_monthly.xlsx).

Source: Ministry of Finance / Japan Customs, "Trade Statistics of Japan", the monthly press release PDF (報道発表; page 3,
major import commodities with quantity, value and change; https://www.customs.go.jp/toukei/shinbun/happyou.htm). For each
month the newest version of the release listed on that page is read (the file name ends in 5 = preliminary, 6, 7 =
final as the month ages). Quantities as published: crude oil thousand kilolitres; LNG, LPG, coal and steam coal (一般炭)
thousand tonnes (LNG is a weight, not converted to a gas volume); values million yen.
The release gives no origin split for these commodities; imports by origin country need the customs database query
(JavaScript form, not reachable) - see the Sources tab of the master.

Sheet 'Monthly': month (1st), LNG_kt, LNG_value_million_yen, Crude_oil_thousand_kl, Crude_oil_value_million_yen, LPG_kt,
Coal_kt, Steam_coal_kt. Sheet 'Releases': month, file, suffix, release title line.

Incremental: the committed workbook is the history store; a month is fetched again only when a newer version of its release
is listed or it is one of the last three months. History from Jan 2022.

    python3 japan/JAPAN_MOF_TRADE.py --out "output/Data and Chart Outputs/japan_mof_energy_imports_monthly.xlsx"
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

BASE = "https://www.customs.go.jp/toukei/shinbun/"
LIST_URL = BASE + "happyou.htm"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
START = pd.Timestamp("2022-01-01")
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "japan_mof_energy_imports_monthly.xlsx")
NUM = r"(-?[\d,]+(?:\.\d+)?)"
# label regex -> (quantity column, value column)
ITEMS = {
    r"粗油": ("Crude_oil_thousand_kl", "Crude_oil_value_million_yen"),
    r"液化天然ガス": ("LNG_kt", "LNG_value_million_yen"),
    r"液化石油ガス": ("LPG_kt", None),
    r"　石炭": ("Coal_kt", None),
    r"（一般炭）": ("Steam_coal_kt", None),
}
BOUNDS = {"LNG_kt": (1500, 10000), "Crude_oil_thousand_kl": (6000, 20000), "LPG_kt": (200, 2500), "Coal_kt": (7000, 22000)}


def get(url, tries=3):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers=UA, timeout=(10, 90))
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            time.sleep(4 * (i + 1))
    raise last


def listing():
    """{month: (suffix digit, relative file)} - the newest version of each month's release on the list page."""
    t = get(LIST_URL).content.decode("utf-8", "replace")
    best = {}
    for y, m, s, in re.findall(r"trade-st/\d{4}/(\d{4})(\d{2})(\d)\.pdf", t):
        month = pd.Timestamp(int(y), int(m), 1)
        if 1 <= int(m) <= 12 and (month not in best or int(s) > best[month][0]):
            best[month] = (int(s), f"trade-st/{y}/{y}{m}{s}.pdf")
    return best


def num(s):
    return float(s.replace(",", ""))


def parse_pdf(content, month):
    doc = pymupdf.open(stream=content, filetype="pdf")
    head = doc[0].get_text()
    mm = re.search(r"令和\s*(\d+)\s*年\s*(\d+)\s*月分", head)
    if mm and (2018 + int(mm.group(1)), int(mm.group(2))) != (month.year, month.month):
        raise ValueError(f"file is for {2018 + int(mm.group(1))}-{mm.group(2)}, expected {month:%Y-%m}")
    text = next((pg.get_text() for pg in doc if "液化天然ガス" in pg.get_text()), None)
    if text is None:
        raise ValueError("no page with 液化天然ガス")
    out = {}
    for label, (qcol, vcol) in ITEMS.items():
        m = re.search(label + r"\n" + NUM + r"\n" + NUM + r"\n" + NUM, text)
        if not m:
            continue
        out[qcol] = num(m.group(1))
        if vcol:
            out[vcol] = num(m.group(3))
    for col, (lo, hi) in BOUNDS.items():
        if col in out and not lo <= out[col] <= hi:
            raise ValueError(f"{col}={out[col]} outside {lo}-{hi}")
    if "LNG_kt" not in out:
        raise ValueError("LNG row not found")
    return out, head


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    old = rel = None
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Monthly", index_col=0, parse_dates=True)
            rel = pd.read_excel(args.out, sheet_name="Releases", index_col=0, parse_dates=True)
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
            old = rel = None
    lst = listing()
    print(f"  {len(lst)} months listed, {min(lst):%Y-%m}..{max(lst):%Y-%m}")
    last3 = sorted(lst)[-3:]
    rows, relrows = {}, {}
    if old is not None:
        rows = {k: v.dropna().to_dict() for k, v in old.iterrows()}
        relrows = {k: v.to_dict() for k, v in rel.iterrows()}
    for month in sorted(lst):
        if month < START:
            continue
        suffix, rel_file = lst[month]
        have = relrows.get(month)
        if have is not None and int(have.get("suffix", 0)) >= suffix and month not in last3:
            continue
        try:
            r = get(BASE + rel_file)
            vals, head = parse_pdf(r.content, month)
        except Exception as e:  # noqa: BLE001
            print(f"  {month:%Y-%m} {rel_file}: {type(e).__name__}: {e}", file=sys.stderr)
            continue
        rows[month] = vals
        relrows[month] = {"file": rel_file, "suffix": suffix,
                          "title": re.sub(r"\s+", " ", " ".join(re.findall(r"令和[^\n]*分貿易統計[^\n]*", head)))[:120]}
        print(f"  {month:%Y-%m}: LNG {vals['LNG_kt']:,.0f} kt, crude {vals.get('Crude_oil_thousand_kl', float('nan')):,.0f} thousand kl")
        time.sleep(0.3)
    if not rows:
        raise SystemExit("no data")
    data = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    cols = ["LNG_kt", "LNG_value_million_yen", "Crude_oil_thousand_kl", "Crude_oil_value_million_yen", "LPG_kt", "Coal_kt",
            "Steam_coal_kt"]
    data = data[[c for c in cols if c in data.columns]]
    data.index.name = "month"
    releases = pd.DataFrame.from_dict(relrows, orient="index").sort_index()
    releases.index.name = "month"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Monthly": data, "Releases": releases}, NOTES, TITLES)
    ann = data["LNG_kt"].tail(12).sum() / 1000
    print(f"Saved {len(data)} months ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}); last 12 months LNG {ann:.1f} Mt -> {args.out}")


NOTES = [
    "UNITS",
    "LNG_kt, LPG_kt, Coal_kt, Steam_coal_kt: thousand tonnes imported in the month (weights; LNG is not converted to a gas "
    "volume). Crude_oil_thousand_kl: thousand kilolitres. *_value_million_yen: import value, million yen.",
    "",
    "SOURCE",
    "Ministry of Finance (Japan Customs), Trade Statistics of Japan, monthly press release (page 3, major import commodities): "
    "https://www.customs.go.jp/toukei/shinbun/happyou.htm . Preliminary figures are replaced by the later versions of the same "
    "release as the month ages.",
    "",
    "UPDATES",
    "Incremental: a month is read again only when a newer version of its release is listed, plus the last three months.",
]
TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
