"""
PNG versions of the main Europe master charts, drawn from the master workbook's own chart-data tabs so they show
exactly what the Excel charts show (CLAUDE.md: every new or updated PNG chart is committed to output/PNG Charts/):

  europe_power_generation_by_source.png   Europe power generation by source, GWh per month
  europe_power_balance.png                Europe supply (by source + net imports) against load
  europe_power_balance_major_markets.png  the same for Germany, France, Italy and Spain
  europe_net_imports.png                  net electricity imports (+) and exports (-) by country
  europe_power_prices.png                 day-ahead prices, selected countries, monthly average
  europe_gas_storage_water_year.png       EU gas storage, AGSI-style water year (Oct-Sep)
  europe_gas_balance.png                  EU gas balance: supply and storage flows vs consumption\n  europe_gas_storage_flows.png            EU storage withdrawals (+) and injections (-)
  europe_lng_sendout.png                  LNG terminal send-out by country

Colours are the repo's fixed categorical order (xlsx_charts.PALETTE) so a fuel keeps its colour on every chart; no
chart or plot-area borders; dates mmm/yy.

Usage: python3 EUROPE_PNG_CHARTS.py [--master "output/Data and Chart Outputs/europe_master.xlsx"]
                                    [--out-dir "output/PNG Charts"]
"""
import argparse
import os
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_charts  # noqa: E402

PAL = ["#" + c for c in xlsx_charts.PALETTE]
GREY, INK, MUTED = "#9A9A9A", "#222222", "#6B6B6B"
FUEL_COLOURS = dict(zip(["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other", "Net imports"], PAL))
ENTSOE = "ENTSO-E Transparency Platform"


def tab(xl, name):
    """A chart-data tab: Date + series columns (extra note columns dropped)."""
    d = pd.read_excel(xl, name)
    d = d[[c for c in d.columns if not str(c).startswith("Unnamed")]]
    first = d.columns[0]
    keep = [first] + [c for c in d.columns[1:] if d[c].notna().sum() > 0 and not str(c).startswith(("Countries", "Notes", "Feeds"))]
    d = d[keep].dropna(subset=[first])
    d[first] = pd.to_datetime(d[first])
    return d.set_index(first).apply(pd.to_numeric, errors="coerce")


def style(ax):
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#BBBBBB")
    ax.tick_params(colors=MUTED, length=0, labelsize=9)
    ax.grid(axis="y", color="#E6E6E6", linewidth=0.7)
    ax.set_axisbelow(True)


def date_axis(ax, every=6):
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7] if every == 6 else [1]))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
    plt.setp(ax.get_xticklabels(), rotation=-45, ha="left")
    lo, hi = ax.dataLim.x0, ax.dataLim.x1
    ax.set_xlim(lo - 20, hi + 20)


def finish(fig, ax, title, ylabel, source, path, legend=True, ncol=8):
    ax.set_title(title, loc="left", fontsize=13, color=INK, fontweight="bold", pad=14)
    ax.set_ylabel(ylabel, color=MUTED, fontsize=9)
    if legend:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=ncol, frameon=False, fontsize=9, labelcolor=INK,
                  handlelength=1.2, columnspacing=1.4)
    ax.text(0, -0.40, "Source: " + source, transform=ax.transAxes, fontsize=8, color=MUTED, ha="left", va="top")
    fig.savefig(path, dpi=150, facecolor="white", bbox_inches="tight", pad_inches=0.25)
    plt.close(fig)
    print("wrote", path)


def stacked(ax, df, colours, width=24, lines=()):
    """Stacked monthly bars handling negative values (positives stack up, negatives down); `lines` are drawn on top."""
    bars = [c for c in df.columns if c not in lines]
    pos, neg = np.zeros(len(df)), np.zeros(len(df))
    for c in bars:
        v = df[c].fillna(0).values
        base = np.where(v >= 0, pos, neg)
        ax.bar(df.index, v, width=width, bottom=base, color=colours[c], label=c, linewidth=0.5, edgecolor="white")
        pos, neg = pos + np.clip(v, 0, None), neg + np.clip(v, None, 0)
    return bars


def latest(df):
    return df.index.max().strftime("%b/%y")


def fuel_chart(xl, tab_name, title, ylabel, path, lines=(), source=ENTSOE, drop=(), by_order=False):
    d = tab(xl, tab_name).drop(columns=list(drop), errors="ignore")
    fig, ax = plt.subplots(figsize=(11, 5.6))
    style(ax)
    cols = {c: FUEL_COLOURS.get(c, GREY) for c in d.columns}
    if by_order:   # series are countries, not fuels: the repo's fixed categorical order, "Other" last in grey
        cols = {c: (GREY if c == "Other" else PAL[i % len(PAL)]) for i, c in enumerate(d.columns)}
    cols.update({"Pumped & battery (net)": "#4A3AA7", "Load": INK})
    stacked(ax, d, cols, lines=lines)
    if "Load" in lines:
        ax.plot(d.index, d["Load"], color=INK, linewidth=2, label="Load (demand)")
    if "Pumped & battery (net)" in lines:
        ax.plot(d.index, d["Pumped & battery (net)"], color="#4A3AA7", linewidth=1.4, linestyle=(0, (4, 2)), label="Pumped storage & batteries (net)")
    date_axis(ax)
    ax.axhline(0, color="#BBBBBB", linewidth=0.8)
    finish(fig, ax, f"{title} (to {latest(d)})", ylabel, source, path, ncol=5 if lines else 8)


def gas_balance(xl, path):
    d = tab(xl, "EU gas balance data")
    fig, ax = plt.subplots(figsize=(11, 5.6))
    style(ax)
    bars = [c for c in d.columns if c != "Consumption"]
    cols = {c: PAL[i % len(PAL)] for i, c in enumerate(bars)}
    stacked(ax, d[bars], cols)
    ax.plot(d.index, d["Consumption"], color=INK, linewidth=2, label="Consumption (final consumers)")
    date_axis(ax)
    ax.axhline(0, color="#BBBBBB", linewidth=0.8)
    finish(fig, ax, f"EU gas balance: supply and storage flows vs consumption (to {latest(d)})", "TWh per month",
           "ENTSOG physical flows (operational data); Gas Infrastructure Europe ALSI (LNG) and AGSI+ (storage)", path, ncol=4)


def water_year(xl, tab_name, title, unit, path, source):
    d = pd.read_excel(xl, tab_name)
    wy = [c for c in d.columns if str(c).startswith("WY")]
    prev, cur = wy[0], wy[-1]
    x = np.arange(len(d))
    fig, ax = plt.subplots(figsize=(11, 5.4))
    style(ax)
    ax.fill_between(x, d["5Y min"], d["5Y max"], color="#AFC6D9", alpha=0.7, linewidth=0, label="5-year min-max")
    ax.plot(x, d["5Y average"], color="#6E6E6E", linewidth=1.5, linestyle=(0, (4, 3)), label="5-year average")
    ax.plot(x, d[prev], color="#EB6834", linewidth=2, label=str(prev).replace("WY ", "Water year "))
    ax.plot(x, d[cur], color="#0B3A66", linewidth=2.4, label=str(cur).replace("WY ", "Water year "))
    ticks = [i for i, lab in enumerate(d["Axis label"]) if isinstance(lab, str) and lab.strip()]
    ax.set_xticks(ticks)
    ax.set_xticklabels([d["Axis label"][i] for i in ticks], rotation=-45, ha="left")
    ax.set_xlim(0, len(d) - 1)
    ax.set_ylim(bottom=0)
    finish(fig, ax, title, unit, source, path, ncol=4)


def prices(xl, path):
    d = tab(xl, "EU Prices data")
    fig, ax = plt.subplots(figsize=(11, 5.4))
    style(ax)
    for i, c in enumerate(d.columns):
        ax.plot(d.index, d[c], color=PAL[i % len(PAL)], linewidth=2, label=c)
    date_axis(ax)
    finish(fig, ax, f"Europe day-ahead electricity prices, monthly average (to {latest(d)})", "EUR/MWh",
           ENTSOE + ", day-ahead prices", path, ncol=4)


def major_markets(xl, path):
    names = ["Germany", "France", "Italy", "Spain"]
    fig, axes = plt.subplots(2, 2, figsize=(14, 8.4), sharex=False)
    cols = {**FUEL_COLOURS, "Pumped & battery (net)": "#4A3AA7", "Load": INK}
    for ax, name in zip(axes.flat, names):
        d = tab(xl, f"{name} balance data")
        style(ax)
        stacked(ax, d, cols, width=24, lines=("Pumped & battery (net)", "Load"))
        ax.plot(d.index, d["Load"], color=INK, linewidth=1.8, label="Load (demand)")
        ax.plot(d.index, d["Pumped & battery (net)"], color="#4A3AA7", linewidth=1.2, linestyle=(0, (4, 2)),
                label="Pumped storage & batteries (net)")
        ax.axhline(0, color="#BBBBBB", linewidth=0.8)
        ax.set_title(f"{name} (to {latest(d)})", loc="left", fontsize=11, color=INK, fontweight="bold")
        date_axis(ax, every=12)
    handles, labels = axes.flat[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="lower center", ncol=6, frameon=False, fontsize=9, bbox_to_anchor=(0.5, -0.02))
    fig.suptitle("Power balance, major markets: supply by source and net imports vs load (GWh per month)", x=0.01, ha="left",
                 fontsize=13, fontweight="bold", color=INK)
    fig.text(0.01, -0.035, "Source: " + ENTSOE + " (generation, load, cross-border physical flows)", fontsize=8, color=MUTED, va="top")
    fig.tight_layout(rect=(0, 0.03, 1, 0.96))
    fig.savefig(path, dpi=150, facecolor="white", bbox_inches="tight", pad_inches=0.25)
    plt.close(fig)
    print("wrote", path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", default=os.path.join(ROOT, "output", "Data and Chart Outputs", "europe_master.xlsx"))
    ap.add_argument("--out-dir", default=os.path.join(ROOT, "output", "PNG Charts"))
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    xl = pd.ExcelFile(args.master)
    out = lambda n: os.path.join(args.out_dir, n)   # noqa: E731
    names = set(xl.sheet_names)
    jobs = [
        ("Europe generation total data", lambda: fuel_chart(
            xl, "Europe generation total data", "Europe power generation by source", "GWh per month",
            out("europe_power_generation_by_source.png"), source=ENTSOE + ", actual generation per production type")),
        ("Europe balance data", lambda: fuel_chart(
            xl, "Europe balance data", "Europe power balance: supply by source and net imports vs load", "GWh per month",
            out("europe_power_balance.png"), lines=("Pumped & battery (net)", "Load"),
            source=ENTSOE + " (generation, load, cross-border physical flows)")),
        ("Germany balance data", lambda: major_markets(xl, out("europe_power_balance_major_markets.png"))),
        ("EU gas balance data", lambda: gas_balance(xl, out("europe_gas_balance.png"))),
        ("EU Prices data", lambda: prices(xl, out("europe_power_prices.png"))),
        ("EU Storage EU data", lambda: water_year(
            xl, "EU Storage EU data", "EU gas storage", "TWh", out("europe_gas_storage_water_year.png"),
            "Gas Infrastructure Europe, AGSI+")),
        ("EU Storage flows data", lambda: fuel_chart(
            xl, "EU Storage flows data", "EU gas storage: withdrawals (+) and injections (-)", "GWh per month",
            out("europe_gas_storage_flows.png"), source="Gas Infrastructure Europe, AGSI+", by_order=True)),
        ("EU Send-out data", lambda: fuel_chart(
            xl, "EU Send-out data", "EU LNG terminal send-out by country", "GWh per month",
            out("europe_lng_sendout.png"), source="Gas Infrastructure Europe, ALSI", by_order=True)),
    ]
    netname = "Net imports flows data" if "Net imports flows data" in names else None
    if netname:
        jobs.append((netname, lambda: fuel_chart(
            xl, netname, "Europe net electricity imports (+) and exports (-) by country", "GWh per month",
            out("europe_net_imports.png"), source=ENTSOE + ", cross-border physical flows", by_order=True)))
    for tabname, fn in jobs:
        if tabname not in names:
            print(f"skipped (tab '{tabname}' not in the master)")
            continue
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            print(f"FAILED {tabname}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
