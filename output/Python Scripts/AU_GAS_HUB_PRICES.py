"""
Australia east coast gas hub prices beyond the STTM: the Victorian Declared Wholesale Gas Market (DWGM) and the
Wallumbilla gas supply hub (Queensland, the LNG-linked hub). Public NEMWEB reports, no key.

  DWGM        https://nemweb.com.au/Reports/CURRENT/VicGas/int041_v4_market_and_reference_prices_1.csv
              (rolling ~14 gas days; first run backfills from the PublicRpts01..14.zip archives in the same folder)
              price_bod_gst_ex = beginning-of-day market price, imb_wtd_ave_price_gst_ex = daily imbalance-weighted
              average price, A$/GJ
  Wallumbilla https://nemweb.com.au/Reports/CURRENT/GSH/Benchmark_Price/  (AEMO's daily benchmark price, A$/GJ)

Writes au_gas_hub_prices.xlsx, sheet "Daily": date, DWGM_BOD, DWGM_daily_weighted, Wallumbilla_benchmark (A$/GJ).
Incremental: merges each run's rolling window into the saved rows (newest report wins); archives are read only
when nothing is saved yet. Runs daily (the windows are short).

Usage: python3 AU_GAS_HUB_PRICES.py [--out "output/Data and Chart Outputs/au_gas_hub_prices.xlsx"]
"""
import argparse
import io
import os
import re
import sys
import zipfile

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

NEM = "https://nemweb.com.au"
VIC = NEM + "/Reports/CURRENT/VicGas/"
GSH = NEM + "/Reports/CURRENT/GSH/Benchmark_Price/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_gas_hub_prices.xlsx")


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=(10, 180))
    r.raise_for_status()
    return r


def dwgm_frame(d):
    d = d.assign(date=pd.to_datetime(d["gas_date"], format="%d %b %Y", errors="coerce"),
                 stamp=pd.to_datetime(d["current_date"], format="%d %b %Y %H:%M:%S", errors="coerce"))
    d = d.dropna(subset=["date"]).sort_values("stamp").drop_duplicates("date", keep="last").set_index("date")
    return pd.DataFrame({"DWGM_BOD": pd.to_numeric(d["price_bod_gst_ex"], errors="coerce"),
                         "DWGM_daily_weighted": pd.to_numeric(d["imb_wtd_ave_price_gst_ex"], errors="coerce")})


def dwgm(backfill):
    parts = [pd.read_csv(io.BytesIO(get(VIC + "int041_v4_market_and_reference_prices_1.csv").content))]
    if backfill:
        for i in range(1, 15):
            try:
                z = zipfile.ZipFile(io.BytesIO(get(f"{VIC}PublicRpts{i:02d}.zip").content))
            except Exception as e:  # noqa: BLE001
                print(f"  PublicRpts{i:02d}: {type(e).__name__}", flush=True)
                continue
            members = [n for n in z.namelist() if n.lower().startswith("int041")]
            for n in members:
                parts.append(pd.read_csv(z.open(n)))
            print(f"  PublicRpts{i:02d}: {len(members)} INT041 reports", flush=True)
    out = dwgm_frame(pd.concat(parts, ignore_index=True))
    print(f"DWGM: {len(out)} gas days {out.index.min():%Y-%m-%d}..{out.index.max():%Y-%m-%d}", flush=True)
    return out


def aemo_csv(content):
    """AEMO MMS-style CSV (C/I/D rows) or a plain CSV -> DataFrame."""
    text = content.decode("utf-8", "replace")
    lines = text.splitlines()
    if lines and lines[0].startswith("C,"):
        head = next(ln for ln in lines if ln.startswith("I,"))
        cols = ["ROW_TYPE", "REPORT", "TABLE", "VERSION"] + head.split(",")[4:]   # the first 4 fields describe the report
        rows = [ln.split(",") for ln in lines if ln.startswith("D,")]
        return pd.DataFrame(rows, columns=cols[:len(rows[0])] if rows else cols)
    return pd.read_csv(io.StringIO(text))


def wallumbilla(backfill):
    html = get(GSH).text
    files = sorted(set(re.findall(r'href="([^"]+\.(?:zip|csv))"', html, re.I)))
    print(f"GSH benchmark folder: {len(files)} files, last {files[-3:]}", flush=True)
    if not files:
        return pd.DataFrame()
    frames = []
    for f in (files if backfill else files[-10:]):
        content = get(NEM + f if f.startswith("/") else f).content
        if f.lower().endswith(".zip"):
            z = zipfile.ZipFile(io.BytesIO(content))
            for n in z.namelist():
                frames.append(aemo_csv(z.read(n)))
        else:
            frames.append(aemo_csv(content))
    d = pd.concat(frames, ignore_index=True)
    d.columns = [str(c).strip().upper() for c in d.columns]
    dcol = next((c for c in d.columns if "GAS_DATE" in c or c.endswith("DATE")), None)
    pcol = next((c for c in d.columns if "PRICE" in c and "BENCHMARK" in c), None) or \
        next((c for c in d.columns if "PRICE" in c), None)
    print(f"GSH columns: {list(d.columns)[:20]} -> date {dcol}, price {pcol}", flush=True)
    if dcol is None or pcol is None:
        return pd.DataFrame()
    d = d.apply(lambda c: c.astype(str).str.strip('"'))
    if "PRODUCT_LOCATION" in d:
        print(f"  locations: {d['PRODUCT_LOCATION'].value_counts().head(8).to_dict()}; types: "
              f"{d.get('PRODUCT_TYPE', pd.Series(dtype=str)).value_counts().head(6).to_dict()}", flush=True)
        d = d[d["PRODUCT_LOCATION"].str.upper().isin(["WAL", "WALLUMBILLA"])]
    d = d.assign(date=pd.to_datetime(d[dcol].str[:10], format="%Y/%m/%d", errors="coerce"),
                 price=pd.to_numeric(d[pcol], errors="coerce")).dropna(subset=["date", "price"])
    out = d.groupby("date")["price"].mean().to_frame("Wallumbilla_benchmark")   # mean over Wallumbilla products
    print(f"Wallumbilla: {len(out)} gas days {out.index.min():%Y-%m-%d}..{out.index.max():%Y-%m-%d}", flush=True)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    try:
        old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
        old.index = pd.to_datetime(old.index, errors="coerce")
        old = old[old.index.notna()]
    except (FileNotFoundError, ValueError, KeyError, OSError):
        old = pd.DataFrame()
    parts = []
    for name, fn, col in (("DWGM", dwgm, "DWGM_BOD"), ("Wallumbilla", wallumbilla, "Wallumbilla_benchmark")):
        try:
            parts.append(fn(old.empty or col not in old or old[col].notna().sum() < 30))
        except Exception as e:  # noqa: BLE001 - keep the other hub
            print(f"{name} failed: {type(e).__name__}: {str(e)[:200]}", flush=True)
    new = pd.concat([p for p in parts if not p.empty], axis=1) if any(not p.empty for p in parts) else pd.DataFrame()
    if new.empty and old.empty:
        raise SystemExit("No hub prices")
    out = new.combine_first(old) if not old.empty else new
    out = out.sort_index()
    out.index.name = "date"
    print(out.tail(3).to_string(), flush=True)
    notes = [
        "UNITS",
        "A$ per GJ, excluding GST, per gas day.",
        "DWGM_BOD: Victorian Declared Wholesale Gas Market beginning-of-day market price (the 6am schedule price, the "
        "usual Victorian headline price). DWGM_daily_weighted: imbalance-weighted average of the day's five pricing "
        "schedules.",
        "Wallumbilla_benchmark: AEMO's Wallumbilla gas supply hub benchmark price (Queensland, linked to the "
        "Curtis Island LNG plants).",
        "",
        "COVERAGE",
        f"Daily, {out.index.min():%d %b %Y} to {out.index.max():%d %b %Y}. AEMO's current reports are rolling windows, "
        "so history builds from the archive read on the first run plus each daily run.",
        "",
        "SOURCE",
        f"AEMO NEMWEB: {VIC}int041_v4_market_and_reference_prices_1.csv (and PublicRpts archives); {GSH}",
        "https://aemo.com.au/energy-systems/gas/declared-wholesale-gas-market-dwgm",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": out.round(4)}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
