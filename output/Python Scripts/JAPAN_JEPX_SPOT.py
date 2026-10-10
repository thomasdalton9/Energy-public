"""
Pull the JEPX day-ahead (spot) market prices - the Japan Electric Power Exchange's own published results - and keep
them as a growing archive (output/Data and Chart Outputs/japan_jepx_spot_prices.xlsx).

Source: JEPX spot market "Trading Information" page (https://www.jepx.jp/electricpower/market-data/spot/); the data
download behind it is one CSV per fiscal year (April-March), https://www.jepx.jp/js/csv_read.php?dir=spot_summary&file=
spot_summary_<FY>.csv, 48 half-hourly slots a day: system price, nine area prices (Hokkaido ... Kyushu; Okinawa is not
in the market), bid and contracted volume. Shift-JIS text. Prices JPY per kWh, volumes kWh. Reachable from GitHub Actions
(probe 10 Oct 2026, discovery_archive/japan/).

Sheet 'Daily': date, simple average of the day's 48 slot prices (JPY/kWh) for the system price and each area, and the
day's contracted volume (GWh). A day with fewer than 48 slots (the day in progress) is left out. The yen figures are
converted to US$ in the master with the Federal Reserve H.10 rate (japan_fx_usd_daily.xlsx), not here.

Incremental: the committed workbook is the history store; each run fetches only the fiscal-year file(s) that contain
days after the stored history minus a 14-day revision window. History from fiscal 2016 (Apr 2016).

    python3 japan/JAPAN_JEPX_SPOT.py --out "output/Data and Chart Outputs/japan_jepx_spot_prices.xlsx"
"""
import argparse
import io
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://www.jepx.jp/js/csv_read.php"
PAGE = "https://www.jepx.jp/electricpower/market-data/spot/"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Referer": PAGE}
FIRST_FY = 2016
REVISION_DAYS = 14
AREAS = ["Hokkaido", "Tohoku", "Tokyo", "Chubu", "Hokuriku", "Kansai", "Chugoku", "Shikoku", "Kyushu"]
JP_AREAS = ["北海道", "東北", "東京", "中部", "北陸", "関西", "中国", "四国", "九州"]
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "japan_jepx_spot_prices.xlsx")


def fetch_year(fy):
    last = None
    for i in range(3):
        try:
            r = requests.get(URL, params={"dir": "spot_summary", "file": f"spot_summary_{fy}.csv"}, headers=UA,
                             timeout=(10, 120))
            r.raise_for_status()
            if len(r.content) < 1000:
                raise RuntimeError(f"short response ({len(r.content)} bytes)")
            for enc in ("utf-8-sig", "cp932"):
                try:
                    return pd.read_csv(io.StringIO(r.content.decode(enc)))
                except UnicodeDecodeError:
                    continue
            raise ValueError("undecodable (neither UTF-8 nor Shift-JIS)")
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  FY{fy} attempt {i + 1}/3: {type(e).__name__}: {e}", file=sys.stderr)
            time.sleep(5 * (i + 1))
    raise last


def daily_from(df):
    d = df.copy()
    d["date"] = pd.to_datetime(d["受渡日"].astype(str), format="%Y/%m/%d", errors="coerce")
    d = d.dropna(subset=["date"])
    out = pd.DataFrame(index=sorted(d["date"].unique()))
    g = d.groupby("date")
    n = g.size()
    out["System_JPY_per_kWh"] = g["システムプライス(円/kWh)"].mean()
    for en, jp in zip(AREAS, JP_AREAS):
        out[f"{en}_JPY_per_kWh"] = g[f"エリアプライス{jp}(円/kWh)"].mean()
    out["Contracted_GWh"] = g["約定総量(kWh)"].sum() / 1e6
    out = out.loc[n[n == 48].index.intersection(out.index)]   # whole days only
    return out.round(4)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    old = None
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
            old.index = pd.to_datetime(old.index)
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
    today = pd.Timestamp.today().normalize()
    cur_fy = today.year if today.month >= 4 else today.year - 1
    if old is not None and len(old):
        cut = old.index.max() - pd.Timedelta(days=REVISION_DAYS)
        first = cut.year if cut.month >= 4 else cut.year - 1
        keep = old[old.index < cut]
    else:
        first, keep = FIRST_FY, None
    frames = []
    for fy in range(max(first, FIRST_FY), cur_fy + 1):
        try:
            d = daily_from(fetch_year(fy))
        except Exception as e:  # noqa: BLE001
            print(f"  FY{fy} failed ({type(e).__name__}: {e}); skipped", file=sys.stderr)
            continue
        print(f"  FY{fy}: {len(d)} days {d.index.min():%Y-%m-%d}..{d.index.max():%Y-%m-%d}")
        frames.append(d)
    if not frames and old is None:
        raise SystemExit("no fiscal-year file could be read")
    new = pd.concat(frames) if frames else pd.DataFrame()
    if keep is not None:
        new = new[new.index >= keep.index.max() + pd.Timedelta(days=1)] if len(keep) else new
        data = pd.concat([keep, new])
    else:
        data = new
    data = data[~data.index.duplicated(keep="last")].sort_index()
    data.index.name = "date"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": data}, NOTES, TITLES)
    print(f"Saved {len(data)} days ({data.index.min():%Y-%m-%d}..{data.index.max():%Y-%m-%d}); latest system price "
          f"{data['System_JPY_per_kWh'].iloc[-1]:.2f} JPY/kWh -> {args.out}")


NOTES = [
    "UNITS",
    "*_JPY_per_kWh: simple average of the day's 48 half-hourly day-ahead (spot) prices, Japanese yen per kWh "
    "(System = the single-price market; Hokkaido ... Kyushu = area prices after market splitting). Contracted_GWh: the day's "
    "total contracted volume in the spot market. Okinawa is not in the JEPX market.",
    "",
    "SOURCE",
    "Japan Electric Power Exchange (JEPX), spot market trading information, spot_summary CSV by fiscal year: "
    "https://www.jepx.jp/electricpower/market-data/spot/ .",
    "",
    "DATES",
    "Delivery date. A day is kept only when all 48 slots are present.",
    "",
    "UPDATES",
    "Incremental: stored days before the last 14 are kept; the fiscal-year file(s) holding later days are re-read.",
]
TITLES = {"UNITS", "SOURCE", "DATES", "UPDATES"}

if __name__ == "__main__":
    main()
