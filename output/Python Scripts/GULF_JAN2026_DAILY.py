"""
Partial daily view of the January 2026 Gulf Coast gas squeeze (static PNG, read-only on the repo's workbooks):

  1. Henry Hub spot (henry_hub_daily.xlsx, EIA RNGWHHD, USD/MMBtu; spot trades the day before flow, so Friday's price
     covers the weekend gas days, and the series has no weekend rows)
  2. Daily mean temperature of five Texas cities (texas_demand_regression.xlsx 'Texas T2M daily', NASA POWER, deg C)
  3. ERCOT and MISO gas burn for power, estimated daily (ercot_gas_burn_daily.xlsx / miso_gas_burn_daily.xlsx, Bcf/d;
     EIA-930 gas MWh x a calibrated heat rate: an estimate, not measured burn)
  4. US gas-fired generation (eia930_fuel_mix_daily.xlsx sheet US_Total, EIA-930, TWh/day)

Only these series exist daily in the repo. Production, consumption by sector and storage are monthly (see
gulf_coast_gas_balance.xlsx); there are no pipeline, daily production or daily LNG feedgas data for the month.
Usage: python3 americas/GULF_JAN2026_DAILY.py [--start 2026-01-12 --end 2026-02-03]
"""
import argparse
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT = os.path.join(ROOT, "output", "PNG Charts", "gulf_jan2026_daily.png")


def series(path, sheet, col, scale=1.0):
    d = pd.read_excel(os.path.join(DATA, path), sheet_name=sheet)
    d[d.columns[0]] = pd.to_datetime(d[d.columns[0]])
    return d.set_index(d.columns[0])[col] * scale


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default="2026-01-12")
    ap.add_argument("--end", default="2026-02-03")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--height", type=float, default=6.0, help="figure height in inches")
    ap.add_argument("--width", type=float, default=6.0, help="figure width in inches (fonts scale with it)")
    ap.add_argument("--title", default="January 2026 cold snap: daily data")
    a = ap.parse_args()
    s, e = pd.Timestamp(a.start), pd.Timestamp(a.end)
    hh = series("henry_hub_daily.xlsx", "Data", "Henry_Hub_USD_per_MMBtu")
    t = pd.read_excel(os.path.join(DATA, "texas_demand_regression.xlsx"), sheet_name="Texas T2M daily")
    t["date"] = pd.to_datetime(t["date"])
    t = t.set_index("date")
    erc = series("ercot_gas_burn_daily.xlsx", "Daily", "Gas_burn_Bcf_per_day")
    mis = series("miso_gas_burn_daily.xlsx", "Daily", "Gas_burn_Bcf_per_day")
    us = series("eia930_fuel_mix_daily.xlsx", "US_Total", "Natural_Gas_MWh", 1e-6)   # MWh -> TWh

    f = a.width / 13 * 1.55   # font/line scale relative to the original 13-inch layout
    plt.rcParams.update({"font.size": 9 * f})
    fig, axs2 = plt.subplots(2, 2, figsize=(a.width, a.height), sharex=True, dpi=220)
    axs = list(axs2.flatten())
    ink, grid = "#222", "#dcdcdc"
    for ax in axs:
        for sp in ax.spines.values():
            sp.set_visible(False)
        ax.grid(axis="y", color=grid, lw=0.8 * f)
        ax.tick_params(length=0)
        for d in pd.date_range(s, e):
            if d.weekday() >= 5:
                ax.axvspan(d - pd.Timedelta(hours=12), d + pd.Timedelta(hours=12), color="#f0eee8", lw=0, zorder=0)
    h = hh.loc[s:e]
    axs[0].plot(h.index, h.values, color="#C0392B", lw=2.4 * f, label="Spot")
    fm = series("henry_hub_daily.xlsx", "Data", "Henry_Hub_front_month_USD_per_MMBtu").loc[s:e]
    if len(fm):
        axs[0].plot(fm.index, fm.values, color="#1F5FBF", lw=1.8 * f, label="Front-month futures")
        axs[0].legend(frameon=False, fontsize=8 * f, loc="center right")
    pk = h.idxmax()
    axs[0].annotate(f"${h.max():.2f} on {pk:%d %b}", (pk, h.max()), (pk + pd.Timedelta(days=1.2), h.max() * 0.93),
                    fontsize=8.5 * f, color=ink)
    axs[0].set_ylabel("USD/MMBtu")
    axs[0].set_title("Henry Hub: spot and front month", loc="left", fontsize=9.5 * f)
    cols = ["#1F5FBF", "#7B3FA0", "#2E9E6B", "#E07B00", "#888888"]
    for c, col in zip(t.columns, cols):
        x = t[c].loc[s:e]
        axs[1].plot(x.index, x.values, lw=1.6 * f, color=col, label=c)
    axs[1].axhline(0, color="#444", lw=0.8 * f)
    axs[1].set_ylim(bottom=axs[1].get_ylim()[0] - 5)
    axs[1].set_ylabel("deg C, daily mean")
    axs[1].set_title("Texas city temperatures", loc="left", fontsize=9.5 * f)
    axs[1].legend(frameon=False, fontsize=7.5 * f, ncol=3, loc="lower left")
    axs[2].plot(erc.loc[s:e].index, erc.loc[s:e].values, color="#E07B00", lw=2 * f, label="ERCOT")
    axs[2].plot(mis.loc[s:e].index, mis.loc[s:e].values, color="#1F5FBF", lw=2 * f, label="MISO")
    axs[2].set_ylabel("Bcf/d")
    axs[2].set_title("Gas burn for power (estimate)", loc="left", fontsize=9.5 * f)
    axs[2].legend(frameon=False, fontsize=8 * f, loc="upper left")
    axs[3].plot(us.loc[s:e].index, us.loc[s:e].values, color="#2E7D4F", lw=2 * f)
    axs[3].set_ylabel("TWh/day")
    axs[3].set_title("US gas-fired generation (EIA-930)", loc="left", fontsize=9.5 * f)
    for ax in axs:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=6 if (e - s).days <= 35 else 14))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
        ax.tick_params(labelbottom=True)
    fig.suptitle(a.title, x=0.01, ha="left",
                 fontsize=12 * f, fontweight="bold", color=ink)
    fig.text(0.01, 0.005, "Sources: EIA (Henry Hub spot, EIA-930), Yahoo Finance NG=F (NYMEX front month,\nunadjusted at rolls), NASA POWER, own gas-burn estimates (EIA-930 gas MWh x heat rate).\nProduction, storage, sector demand: monthly only. Shaded = weekends.",
             fontsize=6 * f, color="#555")
    fig.subplots_adjust(left=0.09, right=0.98, top=0.92, bottom=0.11, hspace=0.3, wspace=0.28)
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    fig.savefig(a.out, facecolor="white")
    print("saved", a.out)


if __name__ == "__main__":
    main()
