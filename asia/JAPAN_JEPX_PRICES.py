"""
Japan wholesale power prices: JEPX day-ahead spot market (no key, no login).

JEPX publishes one CSV per fiscal year (April-March), 30-minute rows, Shift-JIS:
  https://www.jepx.jp/market/excel/spot_<FY>.csv
Columns: date, time code (1-48), sell / buy bid volume, contract volume (kWh), system price (JPY/kWh), then the
area prices (Hokkaido, Tohoku, Tokyo, Chubu, Hokuriku, Kansai, Chugoku, Shikoku, Kyushu). Columns are found by
name, not position.

Output (standard layout; sheet "Daily", one row per day): System_JPY_kWh and <Area>_JPY_kWh = daily average of the
48 prices, Volume_MWh = contracted volume. Incremental: the stored sheet is read back and only fiscal years
containing days after (last stored day - 7) are downloaded.

    python3 asia/JAPAN_JEPX_PRICES.py --out "output/Data and Chart Outputs/japan_power_prices_daily.xlsx"
"""
import argparse
import io
import os
import sys
import unicodedata
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Referer": "https://www.jepx.jp/electricpower/market-data/spot/"}
URL = "https://www.jepx.jp/market/excel/spot_{fy}.csv"
DATA_START = date(2021, 4, 1)
AREAS = {"北海道": "Hokkaido", "東北": "Tohoku", "東京": "Tokyo", "中部": "Chubu", "北陸": "Hokuriku", "関西": "Kansai",
         "中国": "Chugoku", "四国": "Shikoku", "九州": "Kyushu"}
SHEET = "Daily"


def norm(s):
    return unicodedata.normalize("NFKC", str(s)).strip()


def parse(raw):
    for enc in ("cp932", "utf-8-sig"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return None
    df = pd.read_csv(io.StringIO(text), dtype=str)
    df.columns = [norm(c) for c in df.columns]
    dcol = next((c for c in df.columns if "年月日" in c or c.lower() == "date"), df.columns[0])
    num = lambda c: pd.to_numeric(df[c].map(lambda v: norm(v).replace(",", "")), errors="coerce")  # noqa: E731
    out = pd.DataFrame({"date": pd.to_datetime(df[dcol].map(norm), errors="coerce")})
    sys_col = next((c for c in df.columns if "システム" in c), None)
    if sys_col is None:
        return None
    out["System_JPY_kWh"] = num(sys_col)
    for jp, en in AREAS.items():
        c = next((c for c in df.columns if "エリアプライス" in c and jp in c), None)
        if c:
            out[f"{en}_JPY_kWh"] = num(c)
    vol = next((c for c in df.columns if "約定総量" in c), None)
    out["Volume_MWh"] = num(vol) / 1000.0 if vol else float("nan")
    out = out.dropna(subset=["date"])
    g = out.groupby("date")
    d = g.mean(numeric_only=True)
    d["Volume_MWh"] = g["Volume_MWh"].sum(min_count=1)
    d = d[g.size() >= 46]   # a full day has 48 half-hours (allow a couple of blanks)
    return d.round(3)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/japan_power_prices_daily.xlsx")
    args = ap.parse_args()
    old = pd.DataFrame()
    start = DATA_START
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name=SHEET)
            old["date"] = pd.to_datetime(old["date"])
            old = old.set_index("date")
            start = max(start, old.index.max().date() - timedelta(days=7))
        except Exception as e:  # noqa: BLE001
            print(f"could not read stored sheet ({type(e).__name__}); rebuilding")
    frames, notes = [], []
    for fy in range(start.year if start.month >= 4 else start.year - 1,
                    (date.today().year if date.today().month >= 4 else date.today().year - 1) + 1):
        r = requests.get(URL.format(fy=fy), headers=H, timeout=(15, 90))
        notes.append(f"FY{fy}: http {r.status_code}")
        if r.status_code != 200:
            continue
        d = parse(r.content)
        if d is not None and len(d):
            frames.append(d[d.index >= pd.Timestamp(start)])
    print("; ".join(notes))
    if not frames and old.empty:
        sys.exit("no JEPX data fetched")
    keep = [old[old.index < pd.Timestamp(start)]] if len(old) else []
    d = pd.concat(keep + frames).sort_index()
    d = d[~d.index.duplicated(keep="last")]
    d.index = d.index.strftime("%Y-%m-%d")
    d.index.name = "date"
    lines = ["UNITS", "JPY per kWh, daily average of the 48 half-hourly day-ahead prices (system price and the nine area",
             "prices). Volume_MWh = contracted volume of the day.", "",
             "SOURCE", "JEPX (Japan Electric Power Exchange) spot market data, https://www.jepx.jp/electricpower/market-data/spot/",
             "", "LAST RUN"] + notes
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {SHEET: d}, lines, {"UNITS", "SOURCE", "LAST RUN"})
    print(f"Saved {args.out}: {len(d)} days to {d.index.max()}")


if __name__ == "__main__":
    main()
