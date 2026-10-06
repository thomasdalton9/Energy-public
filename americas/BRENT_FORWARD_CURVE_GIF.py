"""
Animated GIF of the Brent forward curve moving through time, plus two PNGs, from
output/Data and Chart Outputs/brent_forward_curve.xlsx (see BRENT_FORWARD_CURVE.py for source and caveats).

One frame per week (the last trade date of each week), x = contract month, y = USD/bbl, fixed axes,
current curve bold, previous 4 frames faded as a trail, ~10 fps, loops. The curves of past dates contain only
contracts Yahoo still lists today (front months of past dates are missing); the title says so.
"""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image

XLSX = "output/Data and Chart Outputs/brent_forward_curve.xlsx"
OUTDIR = "output/PNG Charts"
BLUE, GREY, ORANGE = "#1F5FA8", "#8A8F98", "#D9822B"
SRC = "Source: NYMEX Brent Last Day Financial (BZ) daily closes via Yahoo Finance; contracts listed today only"


def style(ax):
    for s in ax.spines.values():
        s.set_visible(False)
    ax.grid(axis="y", color="#E3E5E8", lw=0.8)
    ax.tick_params(length=0, labelsize=9)


def main():
    cur = pd.read_excel(XLSX, sheet_name="Curve", index_col=0)
    cur.index = pd.to_datetime(cur.index)
    cur.columns = pd.to_datetime([f"{c}-01" for c in cur.columns])
    cur = cur.sort_index().dropna(how="all")
    sp = pd.read_excel(XLSX, sheet_name="Spreads", index_col=0)
    sp.index = pd.to_datetime(sp.index)
    os.makedirs(OUTDIR, exist_ok=True)
    wk = cur.groupby(cur.index.to_period("W")).tail(1).index
    START = pd.Timestamp("2021-01-01")   # owner: from 2021, not 2018-2020
    frames = [d for d in wk if d >= START]
    c21 = cur[cur.index >= START]
    ymin, ymax = np.floor(c21.min().min() / 10) * 10 - 5, np.ceil(c21.max().max() / 10) * 10 + 5
    xl = (cur.columns.min() - pd.Timedelta(days=20), cur.columns.max() + pd.Timedelta(days=20))
    spread_col = "Dec 2026 less Dec 2027"
    n_c = cur.notna().sum(axis=1)
    imgs = []
    fig, ax = plt.subplots(figsize=(9.6, 5.4), dpi=100)
    for k, d in enumerate(frames):
        ax.clear()
        style(ax)
        for j, b in enumerate(range(max(0, k - 4), k)):
            r = cur.loc[frames[b]].dropna()
            ax.plot(r.index, r.values, color=GREY, lw=1.4, alpha=0.18 + 0.14 * j)
        r = cur.loc[d].dropna()
        ax.plot(r.index, r.values, color=BLUE, lw=3)
        ax.set_xlim(*xl)
        ax.set_ylim(ymin, ymax)
        ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7]))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
        ax.set_ylabel("USD per barrel", fontsize=9)
        s = sp[spread_col].get(d, np.nan) if spread_col in sp else np.nan
        shape = "" if pd.isna(s) else f"   |   Dec26 less Dec27: {s:+.1f} ({'backwardation' if s > 0 else 'contango'})"
        ax.set_title(f"Brent forward curve, trade date {d:%d %b %Y}{shape}", fontsize=12, loc="left", color="#222")
        fig.text(0.01, 0.01, f"{SRC}. {int(n_c[d])} contracts shown.", fontsize=7, color="#555")
        fig.subplots_adjust(left=0.07, right=0.98, top=0.9, bottom=0.12)
        fig.canvas.draw()
        im = Image.fromarray(np.asarray(fig.canvas.buffer_rgba())[:, :, :3])
        imgs.append(im.quantize(colors=48, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE))
        if k == 0:
            for t in list(fig.texts):
                pass
        for t in list(fig.texts):
            t.remove()
    plt.close(fig)
    durs = [100] * len(imgs)
    durs[-1] = 2000
    path = os.path.join(OUTDIR, "brent_forward_curve.gif")
    imgs[0].save(path, save_all=True, append_images=imgs[1:], duration=durs, loop=0, optimize=True, disposal=1)
    print(f"GIF {path}: {len(imgs)} frames, {os.path.getsize(path) / 1e6:.1f} MB, {frames[0]:%Y-%m-%d} to {frames[-1]:%Y-%m-%d}")

    # PNG 1: latest vs earlier snapshots (contracts through Dec 2027)
    snap = pd.read_excel(XLSX, sheet_name="Snapshots", index_col=0)
    snap.index = pd.to_datetime(snap.index)
    snap = snap[snap.index <= "2027-12-31"].dropna(how="all")
    fig, ax = plt.subplots(figsize=(9.6, 5.4), dpi=110)
    style(ax)
    cols = [BLUE, ORANGE, "#2E9E6B", "#8E5CC4", GREY]
    for c, colr in zip([c for c in snap.columns if not c.startswith("first")], cols):
        s_ = snap[c].dropna()
        ax.plot(s_.index, s_.values, color=colr, lw=3 if c.startswith("latest") else 1.8, label=c)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
    ax.set_ylabel("USD per barrel", fontsize=9)
    ax.legend(frameon=False, fontsize=8, loc="upper right")
    ax.set_title("Brent forward curve by contract month: latest vs earlier dates", loc="left", fontsize=12)
    fig.text(0.01, 0.01, SRC, fontsize=7, color="#555")
    fig.subplots_adjust(left=0.07, right=0.98, top=0.9, bottom=0.12)
    fig.savefig(os.path.join(OUTDIR, "brent_forward_curve.png"))
    plt.close(fig)
    # PNG: history of curves, first trade date of each month (since Jan 2021), fading with age, latest bold
    last = cur.index.max()
    win = cur[cur.index >= START]
    firsts = list(win.groupby(win.index.to_period("M")).head(1).index)
    firsts = [d for d in firsts if d != last]
    fig, ax = plt.subplots(figsize=(11, 6), dpi=110)
    style(ax)
    cmap = plt.get_cmap("viridis")
    lab = {}
    for tgt, nm in ((1, "1 month ago"), (3, "3 months ago"), (6, "6 months ago"), (12, "12 months ago"), (24, "24 months ago"), (36, "36 months ago"), (48, "48 months ago")):
        c_ = [d for d in firsts if d <= last - pd.DateOffset(months=tgt)]
        if c_:
            lab[c_[-1]] = nm
    for i, d in enumerate(firsts):
        age = (len(firsts) - i) / len(firsts)          # 1 = oldest
        r = cur.loc[d].dropna()
        ax.plot(r.index, r.values, color=cmap(0.15 + 0.7 * (1 - age)), lw=1.3, alpha=0.18 + 0.5 * (1 - age),
                label=f"{d:%d %b %Y} ({lab[d]})" if d in lab else None)
    r = cur.loc[last].dropna()
    ax.plot(r.index, r.values, color="#C0392B", lw=3.5, label=f"{last:%d %b %Y} (latest)")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
    ax.set_ylabel("USD per barrel", fontsize=9)
    ax.legend(frameon=False, fontsize=8, loc="upper right")
    ax.set_title(f"Brent forward curve: latest and first trade date of each month, {firsts[0]:%b %Y} to {last:%b %Y} (older = fainter)",
                 loc="left", fontsize=11)
    fig.text(0.01, 0.01, SRC + ". Past curves show contracts listed today only (expired front months absent).", fontsize=7, color="#555")
    fig.subplots_adjust(left=0.06, right=0.98, top=0.92, bottom=0.1)
    fig.savefig(os.path.join(OUTDIR, "brent_forward_curve_history.png"))
    plt.close(fig)
    # PNG 2: calendar spreads
    fig, ax = plt.subplots(figsize=(9.6, 5.4), dpi=110)
    style(ax)
    for c, colr in zip(sp.columns, cols):
        s_ = sp[c].dropna()
        ax.plot(s_.index, s_.values, color=colr, lw=1.8, label=c)
    ax.axhline(0, color="#444", lw=0.8)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
    ax.set_ylabel("USD per barrel (positive = backwardation)", fontsize=9)
    ax.legend(frameon=False, fontsize=8)
    ax.set_title("Brent calendar spreads of fixed December contracts", loc="left", fontsize=12)
    fig.text(0.01, 0.01, SRC, fontsize=7, color="#555")
    fig.subplots_adjust(left=0.07, right=0.98, top=0.9, bottom=0.12)
    fig.savefig(os.path.join(OUTDIR, "brent_calendar_spreads.png"))
    plt.close(fig)


if __name__ == "__main__":
    main()
