"""Animated bubble chart: GDP per head vs electricity demand per head, key years every 5 years, tweened."""
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.animation import FFMpegWriter, PillowWriter
from matplotlib.lines import Line2D
from matplotlib.ticker import FuncFormatter
import fundamentals as F

M, O = sys.argv[1], sys.argv[2]
d = F.load(M + "/output/Data and Chart Outputs")
dem, pop, ppp = d["Demand_TWh"], d["Population_m"], d["GDP_PPP_USD_bn"]
KEYS = [2000, 2005, 2010, 2015, 2020, 2024]
codes = [c for c in dem.columns if c != "WLD" and c in pop and c in ppp]
ok = [c for c in codes if all(not pd.isna(dem[c].get(y)) and not pd.isna(pop[c].get(y)) and not pd.isna(ppp[c].get(y))
                               and dem[c].get(y) > 0 for y in KEYS) and pop[c].get(2024, 0) >= 0.3]
gdp = pd.DataFrame({c: [ppp[c][y] / pop[c][y] for y in KEYS] for c in ok}, index=KEYS)
kwh = pd.DataFrame({c: [dem[c][y] * 1e9 / (pop[c][y] * 1e6) for y in KEYS] for c in ok}, index=KEYS)
twh = pd.DataFrame({c: [dem[c][y] for y in KEYS] for c in ok}, index=KEYS)
popf = pd.DataFrame({c: [pop[c][y] for y in KEYS] for c in ok}, index=KEYS)
print(len(ok), "countries in every key year")


def fit(y):
    x, v, w = gdp.loc[y].values, np.log(kwh.loc[y].values), np.sqrt(popf.loc[y].values)
    best = None
    for K in np.linspace(5000, 20000, 31):
        for g0 in np.linspace(5, 80, 38):
            for b in np.linspace(0.8, 3.0, 23):
                r = v - np.log(K / (1 + (g0 / x) ** b))
                s = np.sum(w * r * r)
                if best is None or s < best[0]:
                    best = (s, K, g0, b)
    return best[1:]


fits = {y: fit(y) for y in KEYS}
print(fits)
south = ["IND", "PAK", "BGD", "LKA", "NPL", "BTN"]
sea = ["THA", "VNM", "PHL", "IDN", "MYS", "SGP", "MMR", "KHM", "LAO", "BRN", "TLS"]
hi = [c for c in south + sea if c in ok]
name = {v: k for k, v in F.ISO3.items()}
ref = {"CHN": "China", "JPN": "Japan", "KOR": "Korea", "USA": "US", "DEU": "Germany", "GBR": "UK", "BRA": "Brazil",
       "NGA": "Nigeria"}
colour = {c: ("#EB6834" if c in south else "#2A78D6") for c in hi}
TWEEN, HOLD = 14, 10
frames = []
for i, y in enumerate(KEYS):
    frames += [(i, 0.0)] * HOLD
    if i < len(KEYS) - 1:
        frames += [(i, t / TWEEN) for t in range(1, TWEEN)]


def at(frame_df, i, t, log=True):
    a, b = frame_df.iloc[i], frame_df.iloc[min(i + 1, len(KEYS) - 1)]
    return np.exp(np.log(a) * (1 - t) + np.log(b) * t) if log else a * (1 - t) + b * t


fig, ax = plt.subplots(figsize=(12, 7.5))
sz = lambda t: 30 + np.sqrt(t) * 22   # noqa: E731
gs = np.exp(np.linspace(np.log(0.6), np.log(200), 200))


def draw(fr):
    i, t = fr
    ax.clear()
    g, k, tw = at(gdp, i, t), at(kwh, i, t), at(twh, i, t)
    year = KEYS[i] + t * (KEYS[min(i + 1, len(KEYS) - 1)] - KEYS[i])
    ax.text(0.98, 0.04, f"{year:.0f}", transform=ax.transAxes, ha="right", va="bottom", fontsize=64, color="#E3E3E3",
            fontweight="bold", zorder=0)
    bg = [c for c in ok if c not in hi]
    ax.scatter(g[bg], k[bg], s=sz(tw[bg]), color="#C8C8C8", alpha=0.55, edgecolor="white", linewidth=0.8, zorder=1)
    # curve: interpolate the fitted parameters between key years
    p0, p1 = np.array(fits[KEYS[i]]), np.array(fits[KEYS[min(i + 1, len(KEYS) - 1)]])
    K, g0, b = p0 * (1 - t) + p1 * t
    ax.plot(gs, K / (1 + (g0 / gs) ** b), color="#555555", lw=1.6, zorder=2)
    K24, g24, b24 = fits[2024]
    if year < 2024:
        ax.plot(gs, K24 / (1 + (g24 / gs) ** b24), color="#555555", lw=1, ls=":", zorder=2)
    # trails: each highlighted country's path through the key years so far
    for c in hi:
        xs = list(gdp[c].iloc[: i + 1]) + [g[c]]
        ys = list(kwh[c].iloc[: i + 1]) + [k[c]]
        ax.plot(xs, ys, color=colour[c], lw=1, alpha=0.45, zorder=2)
    ax.scatter(g[hi], k[hi], s=sz(tw[hi]), color=[colour[c] for c in hi], alpha=0.88, edgecolor="white",
               linewidth=2, zorder=3)
    for c in hi:
        ax.annotate(name.get(c, c), (g[c], k[c]), xytext=(7, 4), textcoords="offset points", fontsize=8.5,
                    color="#222222", zorder=5)
    for c, n in ref.items():
        if c in ok:
            ax.annotate(n, (g[c], k[c]), xytext=(5, -11), textcoords="offset points", fontsize=8, color="#7A7A7A")
    ax.set_xscale("log"); ax.set_yscale("log")
    ax.set_xlim(0.6, 200); ax.set_ylim(15, 70000)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}k" if v >= 1 else f"${v * 1000:,.0f}"))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.set_xlabel("GDP per head, PPP (constant 2021 international $, log scale)")
    ax.set_ylabel("Electricity demand per head (kWh per year, log scale)")
    ax.set_title("GDP per head vs electricity demand per head, 2000-2024 (bubble size = total demand)", loc="left",
                 fontsize=13)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(color="#e8e8e8", lw=0.6); ax.set_axisbelow(True)
    h = [Line2D([], [], marker="o", ls="", ms=9, mfc=c, mec="white", label=l) for c, l in
         [("#EB6834", "South Asia"), ("#2A78D6", "Southeast Asia"), ("#C8C8C8", "Rest of world")]]
    h += [Line2D([], [], color="#555555", lw=1.6, label="World fit this year (saturating curve)"),
          Line2D([], [], color="#555555", lw=1, ls=":", label="World fit 2024")]
    ax.legend(handles=h, loc="upper left", frameon=False, fontsize=9)


fig.text(0.01, 0.035, f"Source: Ember yearly electricity data (demand); World Bank (GDP PPP, population). {len(ok)} countries "
         "with data in every key year (2000, 2005, 2010, 2015, 2020, 2024).", fontsize=8, color="#6B6B6B")
fig.text(0.01, 0.012, "Lines trace each South / Southeast Asian country's path through the key years; motion between key "
         "years is interpolated.", fontsize=8, color="#6B6B6B")
fig.subplots_adjust(left=0.08, right=0.975, top=0.93, bottom=0.13)
from matplotlib.animation import FuncAnimation
anim = FuncAnimation(fig, draw, frames=frames, interval=110)
anim.save(O + "/gdp_vs_kwh_bubble_2000_2024.mp4", writer=FFMpegWriter(fps=9, bitrate=2400), dpi=110)
anim.save(O + "/gdp_vs_kwh_bubble_2000_2024.gif", writer=PillowWriter(fps=9), dpi=80)
draw((len(KEYS) - 1, 0.0)); fig.savefig(O + "/anim_last.png", dpi=90)
draw((0, 0.0)); fig.savefig(O + "/anim_first.png", dpi=90)
print("done")
