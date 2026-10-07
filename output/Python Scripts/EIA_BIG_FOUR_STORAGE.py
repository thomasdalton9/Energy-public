"""
US petroleum stocks, weekly, with the 'Big Four' (crude oil incl. SPR + motor gasoline + distillate + jet fuel) and each
piece, drawn year-on-year like the usual week-1-to-52 chart: 5-year average, the last four years and the current year.

Output: output/Data and Chart Outputs/us_petroleum_stocks_weekly.xlsx
  Data               weekly stocks, thousand barrels (EIA week ending Friday), one column per series
  <Name> by week     week 1-52 table (5-year average, last 4 years, current year) + native line chart (calendar weeks)
  WY <Name>          the repo's Oct-Sep water-year chart (5-year band, average, previous and current water year)
PNG charts (6 x 3.6 in): output/PNG Charts/us_<name>_storage.png

Source: EIA weekly petroleum status report, EIA API v2 route petroleum/stoc/wstk (series WCESTUS1 crude excl. SPR,
WCSSTUS1 SPR crude, WGTSTUS1 total motor gasoline, WDISTUS1 distillate fuel oil, WKJSTUS1 kerosene-type jet fuel; needs
EIA_API_KEY), falling back to EIA's keyless history workbooks https://www.eia.gov/dnav/pet/hist_xls/<series>w.xls.
Incremental: once the archive exists only the last 6 weeks are re-read (EIA revises the latest week).
Week number = (day of year - 1) // 7 + 1 of the report date, 53 folded into 52; '5-Year' = mean of the five complete years
before the current one (2021-2025 for 2026).
Usage: python3 americas/EIA_BIG_FOUR_STORAGE.py [--out PATH] [--png-dir DIR] [--synthetic]   (--synthetic = fake data, charts testing only)
"""
import argparse
import io
import os
import sys
import time
from datetime import date, timedelta

import numpy as np
import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root: xlsx_notes, xlsx_charts, water_year_chart

import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402
import water_year_chart  # noqa: E402

EIA_API = "https://api.eia.gov/v2/petroleum/stoc/wstk/data/"
EIA_XLS = "https://www.eia.gov/dnav/pet/hist_xls/{sid}w.xls"
SERIES = {  # column -> EIA series id
    "Crude excl SPR": "WCESTUS1",
    "SPR crude": "WCSSTUS1",
    "Gasoline": "WGTSTUS1",
    "Distillate": "WDISTUS1",
    "Jet fuel": "WKJSTUS1",
}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 60)
OVERLAP_WEEKS = 6
STALE_AFTER_DAYS = 21
START = "2010-01-01"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "us_petroleum_stocks_weekly.xlsx")
DEFAULT_PNG = os.path.join(ROOT, "output", "PNG Charts")
# charted series: key -> (sheet label, PNG file, title)
CHARTS = {
    "Big Four": ("Big Four", "us_big_four_storage.png", "U.S. Big Four Storage (Crude with SPR + Gasoline + Distillate + Jet Fuel)"),
    "Crude incl SPR": ("Crude incl SPR", "us_crude_incl_spr_storage.png", "U.S. Crude Oil Storage incl. SPR"),
    "Crude excl SPR": ("Crude excl SPR", "us_crude_excl_spr_storage.png", "U.S. Commercial Crude Oil Storage (excl. SPR)"),
    "SPR crude": ("SPR crude", "us_spr_storage.png", "U.S. Strategic Petroleum Reserve"),
    "Gasoline": ("Gasoline", "us_gasoline_storage.png", "U.S. Motor Gasoline Storage"),
    "Distillate": ("Distillate", "us_distillate_storage.png", "U.S. Distillate Fuel Oil Storage"),
    "Jet fuel": ("Jet fuel", "us_jet_fuel_storage.png", "U.S. Jet Fuel Storage"),
}
YEAR_COLOURS = ["A8A8A8", "8DB763", "BDBDBD", "1F3A93", "D62828", "0B4A3C"]   # 5-Year, then oldest -> current year
FETCH_ATTEMPTS = 3


def get(url, **kw):
    last = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"  attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            time.sleep(5 * attempt)
    raise last


def fetch_api(start):
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise RuntimeError("EIA_API_KEY not set")
    ids = list(SERIES.values())
    rows, offset = [], 0
    while True:
        params = {"api_key": key, "frequency": "weekly", "data[0]": "value", "start": start,
                  "sort[0][column]": "period", "sort[0][direction]": "asc", "offset": offset, "length": 5000}
        for i, sid in enumerate(ids):
            params[f"facets[series][{i}]"] = sid
        resp = get(EIA_API, params=params).json()["response"]
        data = resp.get("data", [])
        rows += data
        offset += len(data)
        if not data or offset >= int(resp.get("total", 0)):
            break
    df = pd.DataFrame(rows)
    if df.empty:
        raise RuntimeError("EIA API returned no rows")
    print(f"  EIA API: {len(df)} rows, units {set(df['units'])}", file=sys.stderr)
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    out = df.pivot_table(index="period", columns="series", values="value", aggfunc="last")
    out.index = pd.to_datetime(out.index)
    return out.rename(columns={v: k for k, v in SERIES.items()})


def fetch_xls(start):
    cols = {}
    for name, sid in SERIES.items():
        raw = pd.read_excel(io.BytesIO(get(EIA_XLS.format(sid=sid)).content), sheet_name="Data 1", header=2)
        s = pd.Series(pd.to_numeric(raw.iloc[:, 1], errors="coerce").values, index=pd.to_datetime(raw.iloc[:, 0], errors="coerce"))
        cols[name] = s[s.index.notna()].dropna()
    out = pd.DataFrame(cols)
    return out[out.index >= pd.Timestamp(start)]


def synthetic():
    """Fake data for chart testing only."""
    idx = pd.date_range("2016-01-01", date.today() - timedelta(days=30), freq="W-FRI")
    rng = np.random.default_rng(1)
    t = np.arange(len(idx))
    base = {"Crude excl SPR": 440000, "SPR crude": 600000, "Gasoline": 230000, "Distillate": 120000, "Jet fuel": 40000}
    out = pd.DataFrame(index=idx)
    for k, b in base.items():
        out[k] = b + 0.05 * b * np.sin(2 * np.pi * t / 52) + rng.normal(0, 0.01 * b, len(idx)).cumsum() * 0.1
    return out


def fetch(start):
    for name, fn in (("EIA API v2 (petroleum/stoc/wstk)", fetch_api), ("EIA history XLS", fetch_xls)):
        print(f"Fetching from {name} ...", file=sys.stderr)
        try:
            df = fn(start)
            if len(df):
                return df, name
        except Exception as e:  # noqa: BLE001
            print(f"  {name} failed: {type(e).__name__}: {e}", file=sys.stderr)
    raise SystemExit("all EIA sources failed")


def add_totals(df):
    df = df.copy()
    df["Crude incl SPR"] = df["Crude excl SPR"] + df["SPR crude"]
    df["Big Four"] = df["Crude incl SPR"] + df["Gasoline"] + df["Distillate"] + df["Jet fuel"]
    return df


def load_archive(path):
    try:
        d = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index)
    return d[[c for c in SERIES if c in d.columns]]


def week_table(s):
    """Week 1-52 table: '5-Year' (mean of the five years before the current one) and the current year plus the 4 before it."""
    s = s.dropna()
    s.index = pd.DatetimeIndex(s.index)
    cur = int(s.index.max().year)
    wk = np.minimum(52, (s.index.dayofyear - 1) // 7 + 1)
    piv = pd.DataFrame({"v": s.values, "year": s.index.year, "wk": wk}).groupby(["wk", "year"]).v.last().unstack("year")
    piv = piv.reindex(range(1, 53))
    hist = [y for y in range(cur - 5, cur) if y in piv.columns]
    t = pd.DataFrame({"Week": range(1, 53)})
    t["5-Year"] = piv[hist].mean(axis=1).round(0).values if hist else np.nan
    for y in range(cur - 4, cur + 1):
        t[str(y)] = piv[y].values if y in piv.columns else np.nan
    return t, cur, hist


def write_week_chart(wb, name, table, title, last):
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    sheet = f"{name} by week"[:31]
    if sheet in wb.sheetnames:
        del wb[sheet]
    ws = wb.create_sheet(sheet)
    ws.append(list(table.columns) + ["Axis label"])
    for i, row in enumerate(table.itertuples(index=False), start=2):
        vals = [None if (isinstance(v, float) and pd.isna(v)) else (float(v) if not isinstance(v, (int,)) else int(v)) for v in row]
        ws.append(vals + [f"Week {int(row[0])}" if (int(row[0]) - 1) % 4 == 0 else None])
    ncols = table.shape[1]
    n = len(table) + 1
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=2, max_col=ncols, min_row=1, max_row=n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=ncols + 1, min_row=2, max_row=n))
    for i, s in enumerate(ch.series):
        cur = i == len(ch.series) - 1
        s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=YEAR_COLOURS[min(i, 5)], w=44450 if cur else 19050,
                                                                       prstDash=None if cur else "dash"))
        s.smooth = False
        s.marker.symbol = "none"
    ch.title = f"{title} - thousand barrels (EIA weekly, to {pd.Timestamp(last):%d %b %Y})"
    ch.y_axis.title = "Thousand barrels"
    ch.y_axis.number_format = "#,##0"
    ch.y_axis.numFmt.sourceLinked = False
    vals = table.iloc[:, 1:].stack()
    lo, hi = float(vals.min()), float(vals.max())
    step = water_year_chart._nice_step((hi - lo) or 1.0)
    import math
    ch.y_axis.scaling.min = math.floor(lo / step) * step
    ch.y_axis.scaling.max = math.ceil(hi / step) * step
    ch.y_axis.majorUnit = step
    ch.x_axis.tickLblSkip = 1
    ch.x_axis.tickMarkSkip = 1
    ch.x_axis.tickLblPos = "low"
    xlsx_charts.rotated_labels(ch.x_axis)
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.height, ch.width = 12, 26
    xlsx_charts.tidy_layout(ch, True, None)
    ws.add_chart(ch, f"{chr(ord('A') + ncols + 2)}2")
    ws.column_dimensions["A"].width = 7


def draw_png(table, cur, title, path, last):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    f = 6 / 13 * 1.55
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=220)
    fig.subplots_adjust(left=0.15, right=0.975, top=0.80, bottom=0.12)
    ax.plot(table["Week"], table["5-Year"], color="#" + YEAR_COLOURS[0], lw=1.3 * f * 1.6, ls=(0, (4, 3)), label="5-Year")
    cols = table.columns[2:]
    for i, c in enumerate(cols):
        is_cur = i == len(cols) - 1
        ax.plot(table["Week"], table[c], color="#" + YEAR_COLOURS[min(i + 1, 5)], lw=(3.2 if is_cur else 1.3) * f * 1.6,
                ls="-" if is_cur else (0, (4, 3)), label=c, zorder=5 if is_cur else 3)
    lastrow = table[str(cur)].last_valid_index()
    if lastrow is not None:
        x, y = table.loc[lastrow, "Week"], table.loc[lastrow, str(cur)]
        ax.annotate("We are here", xy=(x, y), xytext=(x + 5.5, y), fontsize=7.5, color="#D62828", fontweight="bold", va="center",
                    arrowprops=dict(arrowstyle="<-", color="#D62828", lw=1.3, shrinkA=0, shrinkB=2))
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.grid(axis="y", color="#e3e3e3", lw=0.7)
    ax.set_axisbelow(True)
    ax.tick_params(length=0, labelsize=6.5)
    ax.set_xticks(range(1, 53, 4))
    ax.set_xticklabels([f"Week {w}" for w in range(1, 53, 4)], rotation=0, fontsize=5.5)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, p: f"{v:,.0f}"))
    ax.set_xlim(1, 58)
    ax.set_ylabel("Thousand barrels", fontsize=7, loc="top")
    leg = ax.legend(frameon=False, fontsize=6, ncol=6, loc="lower center", bbox_to_anchor=(0.5, 1.0), columnspacing=1.0, handlelength=2.2)
    fig.patches.append(matplotlib.patches.Rectangle((0, 0.88), 1, 0.12, transform=fig.transFigure, color="#0B4A3C", zorder=0))
    fig.text(0.02, 0.94, title, color="white", fontsize=8.5, fontweight="bold", va="center")
    fig.text(0.02, 0.02, f"Source: EIA weekly petroleum status report, to {pd.Timestamp(last):%d %b %Y}. 5-Year = 2021-2025 average." if False else
             f"Source: EIA (weekly, to {pd.Timestamp(last):%d %b %Y}). 5-Year = average of the five years before {cur}.", fontsize=5.5, color="#555")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fig.savefig(path, facecolor="white")
    plt.close(fig)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--png-dir", default=DEFAULT_PNG)
    ap.add_argument("--synthetic", action="store_true", help="fake data, for testing the charts only")
    a = ap.parse_args()

    if a.synthetic:
        new, source = synthetic(), "synthetic (TEST ONLY)"
    else:
        existing = load_archive(a.out)
        start = START if existing.empty else (existing.index.max() - timedelta(weeks=OVERLAP_WEEKS)).date().isoformat()
        new, source = fetch(start)
    if not a.synthetic:
        missing = [c for c in SERIES if c not in new.columns]
        if missing:
            raise SystemExit(f"series missing from the EIA response: {missing}")
        data = new if existing.empty else new.combine_first(existing).sort_index()
        data.loc[new.index, list(SERIES)] = new[list(SERIES)]
    else:
        data = new
    data = data[list(SERIES)].dropna(how="all").sort_index()
    data.index.name = "date"
    full = add_totals(data.dropna(subset=list(SERIES)))
    last = full.index.max()
    age = (date.today() - last.date()).days
    print(f"Source {source}; {len(data)} weeks, {data.index.min().date()} to {data.index.max().date()}; complete rows to {last.date()}", file=sys.stderr)

    notes = [
        "UNITS", "Thousand barrels (MBBL) of stocks at week end, US total. Big Four = crude oil incl. SPR + total motor gasoline + distillate fuel oil + kerosene-type jet fuel.", "",
        "COVERAGE", f"Weekly, EIA week ending Friday, from {data.index.min().date()}. Charts use week 1-52 = (day of year - 1) // 7 + 1 of the report date (53 folded into 52).",
        "5-Year = mean of the five complete years before the current one; the chart shows it with the previous four years and the current year.",
        "Calendar-year weeks as in EIA / trade charts; the 'WY' sheets give the repo's Oct-Sep water-year view of the same data.", "",
        "SOURCE", f"EIA weekly petroleum status report via EIA API v2 petroleum/stoc/wstk (series {', '.join(SERIES.values())}); fallback EIA's keyless history workbooks. Last run's source: {source}.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    xlsx_notes.write_workbook(a.out, {"Data": full[list(SERIES) + ["Crude incl SPR", "Big Four"]]}, notes, {"UNITS", "COVERAGE", "SOURCE"})

    from openpyxl import load_workbook
    wb = load_workbook(a.out)
    for key, (label, png, title) in CHARTS.items():
        t, cur, hist = week_table(full[key])
        write_week_chart(wb, label, t, title, last)
        water_year_chart.add_water_year_chart(a.out, full[key], title, "thousand barrels", sheet_name=f"WY {label}"[:31], y_decimals=0, wb=wb)
        draw_png(t, cur, title, os.path.join(a.png_dir, png), last)
    root, ext = os.path.splitext(a.out)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    wb.save(tmp)
    os.replace(tmp, a.out)
    print(f"Saved {a.out} and {len(CHARTS)} PNGs in {a.png_dir}", file=sys.stderr)
    print(full[["Crude incl SPR", "Gasoline", "Distillate", "Jet fuel", "Big Four"]].tail(4).round(0).to_string())
    if not a.synthetic and age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest complete week is {last.date()} ({age} days old)", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
