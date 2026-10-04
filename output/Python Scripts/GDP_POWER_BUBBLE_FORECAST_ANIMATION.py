"""Animated bubble charts, one per continent: GDP per head vs electricity demand per head, key years every 5 years
(2000-2024), tweened; the continent's countries coloured by sub-region, the rest of the world grey."""
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.animation import FFMpegWriter, FuncAnimation, PillowWriter
from matplotlib.lines import Line2D
from matplotlib.ticker import FuncFormatter
import fundamentals as F

M, O = sys.argv[1], sys.argv[2]
only = sys.argv[3:]   # optional: continent names to render
d = F.load(M + "/output/Data and Chart Outputs")
dem, pop, ppp = d["Demand_TWh"], d["Population_m"], d["GDP_PPP_USD_bn"]
wb = pd.read_excel(M + "/output/Data and Chart Outputs/macro_drivers.xlsx", sheet_name="Countries")
wbname = dict(zip(wb.iso3, wb.name))
KEYS = [2000, 2005, 2010, 2015, 2020, 2024]
import forecast_model as FM
FG, FP, FK, FCURVE, _ = FM.forecast_v2(d)
FYEARS = [2030, 2035]
PAL = ["#2A78D6", "#EB6834", "#1BAF7A", "#EDA100"]

S = lambda s: s.split()   # noqa: E731
CONTINENTS = {
    "South & Southeast Asia": [("Southeast Asia", S("THA VNM PHL IDN MYS SGP MMR KHM LAO BRN TLS")),
                               ("South Asia", S("IND PAK BGD LKA NPL BTN"))],
    "Asia": [("East Asia", S("CHN JPN KOR PRK MNG HKG MAC")),
             ("Southeast Asia", S("THA VNM PHL IDN MYS SGP MMR KHM LAO BRN TLS")),
             ("South Asia", S("IND PAK BGD LKA NPL BTN AFG MDV")),
             ("Middle East, Central Asia & Caucasus",
              S("ARE BHR IRN IRQ ISR JOR KWT LBN OMN PSE QAT SAU SYR YEM KAZ KGZ TJK TKM UZB ARM AZE GEO"))],
    "Europe": [("Western & Northern", S("GBR IRL FRA BEL NLD LUX DEU AUT CHE DNK NOR SWE FIN ISL")),
               ("Southern & Turkey", S("ESP PRT ITA GRC MLT CYP TUR")),
               ("Central & Eastern", S("POL CZE SVK HUN ROU BGR SVN HRV SRB BIH MNE MKD ALB XKX EST LVA LTU UKR BLR "
                                       "MDA RUS"))],
    "Africa": [("North Africa", S("DZA EGY LBY MAR TUN")),
               ("Sub-Saharan Africa", list(wb[wb.region == "Sub-Saharan Africa"].iso3) + ["DJI"])],
    "North America": [("United States & Canada", S("USA CAN")),
                      ("Mexico & Central America", S("MEX GTM BLZ SLV HND NIC CRI PAN")),
                      ("Caribbean", S("CUB DOM HTI JAM PRI TTO BHS BRB ATG DMA GRD KNA LCA VCT ABW CUW"))],
    "South America": [("Brazil & the Guianas", S("BRA GUY SUR")),
                      ("Andean", S("BOL COL ECU PER VEN")),
                      ("Southern Cone", S("ARG CHL PRY URY"))],
    "Oceania": [("Australia & New Zealand", S("AUS NZL")),
                ("Pacific islands", S("PNG FJI SLB VUT WSM TON KIR FSM NCL PYF"))],
}

codes = [c for c in dem.columns if c != "WLD" and c in pop and c in ppp]
ok = [c for c in codes if all(not pd.isna(dem[c].get(y)) and not pd.isna(pop[c].get(y)) and not pd.isna(ppp[c].get(y))
                               and dem[c].get(y) > 0 for y in KEYS) and pop[c].get(2024, 0) >= 0.3]
ok = [c for c in ok if c in FK.columns]
gdp = pd.DataFrame({c: [ppp[c][y] / pop[c][y] for y in KEYS] + [FG[c][y] for y in FYEARS] for c in ok}, index=KEYS + FYEARS)
kwh = pd.DataFrame({c: [dem[c][y] * 1e9 / (pop[c][y] * 1e6) for y in KEYS] + [FK[c][y] for y in FYEARS] for c in ok},
                   index=KEYS + FYEARS)
twh = pd.DataFrame({c: [dem[c][y] for y in KEYS] + [FK[c][y] * FP[c][y] / 1000 for y in FYEARS] for c in ok},
                   index=KEYS + FYEARS)
popf = pd.DataFrame({c: [pop[c][y] for y in KEYS] + [FP[c][y] for y in FYEARS] for c in ok}, index=KEYS + FYEARS)
ACT = KEYS
KEYS = KEYS + FYEARS
SHORT = {"USA": "US", "GBR": "UK", "KOR": "Korea", "PRK": "North Korea", "RUS": "Russia", "IRN": "Iran",
         "COD": "DR Congo", "COG": "Congo", "EGY": "Egypt", "VEN": "Venezuela", "SYR": "Syria", "YEM": "Yemen",
         "LAO": "Laos", "VNM": "Vietnam", "HKG": "Hong Kong", "CZE": "Czechia", "SVK": "Slovakia", "TUR": "Turkey",
         "BHS": "Bahamas", "KGZ": "Kyrgyzstan", "MKD": "North Macedonia", "CIV": "Côte d'Ivoire", "GMB": "Gambia",
         "TZA": "Tanzania", "PNG": "Papua New Guinea", "BIH": "Bosnia", "DOM": "Dominican Rep.", "TTO": "Trinidad",
         "ARE": "UAE", "MDA": "Moldova", "SAU": "Saudi Arabia", "PRI": "Puerto Rico", "BRN": "Brunei"}
label_of = lambda c: SHORT.get(c, str(wbname.get(c, c)))   # noqa: E731


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


fits = {y: fit(y) for y in ACT}
for y in FYEARS:
    fits[y] = fits[2024]   # the forecast moves countries along the 2024 curve
TWEEN, HOLD = 14, 10
frames = []
for i, y in enumerate(KEYS):
    frames += [(i, 0.0)] * HOLD
    if i < len(KEYS) - 1:
        frames += [(i, t / TWEEN) for t in range(1, TWEEN)]


def at(f, i, t):
    a, b = f.iloc[i], f.iloc[min(i + 1, len(KEYS) - 1)]
    return np.exp(np.log(a) * (1 - t) + np.log(b) * t)


sz = lambda t: 30 + np.sqrt(t) * 22   # noqa: E731
gs = np.exp(np.linspace(np.log(0.6), np.log(200), 200))
MAX_LABELS = 16
VIEW = {"Europe": ((6, 150), (800, 70000), 11)}   # zoomed axes and fewer labels where countries cluster


def render(continent, groups):
    colour, hi = {}, []
    for k, (_, members) in enumerate(groups):
        for c in members:
            if c in ok and c not in colour:
                colour[c] = PAL[k]
                hi.append(c)
    # label (and trail) the largest by 2024 demand; small continents get every country labelled
    (xl, yl, nlab) = VIEW.get(continent, ((0.6, 200), (15, 70000), MAX_LABELS))
    lab = sorted(hi, key=lambda c: -twh[c].iloc[-1])[:nlab]
    fig, ax = plt.subplots(figsize=(12, 7.5))

    def draw(fr):
        i, t = fr
        ax.clear()
        g, k, tw = at(gdp, i, t), at(kwh, i, t), at(twh, i, t)
        year = KEYS[i] + t * (KEYS[min(i + 1, len(KEYS) - 1)] - KEYS[i])
        fc = year > 2024.01
        if fc:
            ax.text(0.98, (0.04 if continent != "Europe" else 0.80) + 0.13, "FORECAST", transform=ax.transAxes,
                    ha="right", va="bottom", fontsize=20, color="#C9C9C9", fontweight="bold", zorder=0)
        ax.text(0.98, 0.04 if continent != "Europe" else 0.80, f"{year:.0f}", transform=ax.transAxes, ha="right", va="bottom", fontsize=64,
                color="#E3E3E3", fontweight="bold", zorder=0)
        bg = [c for c in ok if c not in colour]
        ax.scatter(g[bg], k[bg], s=sz(tw[bg]), color="#D2D2D2", alpha=0.5, edgecolor="white", linewidth=0.8, zorder=1)
        p0, p1 = np.array(fits[KEYS[i]]), np.array(fits[KEYS[min(i + 1, len(KEYS) - 1)]])
        K, g0, b = p0 * (1 - t) + p1 * t
        ax.plot(gs, K / (1 + (g0 / gs) ** b), color="#555555", lw=1.6, zorder=2)
        if year < 2023.99:
            K24, g24, b24 = fits[2024]
            ax.plot(gs, K24 / (1 + (g24 / gs) ** b24), color="#555555", lw=1, ls=":", zorder=2)
        na = len(ACT)
        for c in lab:
            xs = list(gdp[c].iloc[: i + 1]) + [g[c]]
            ys = list(kwh[c].iloc[: i + 1]) + [k[c]]
            ax.plot(xs[:na], ys[:na], color=colour[c], lw=1, alpha=0.45, zorder=2)
            if len(xs) > na:   # forecast part of the path dashed
                ax.plot(xs[na - 1:], ys[na - 1:], color=colour[c], lw=1.2, alpha=0.7, ls="--", zorder=2)
        order = sorted(hi, key=lambda c: -tw[c])   # big bubbles first so small ones stay visible on top
        ax.scatter(g[order], k[order], s=sz(tw[order]), color=[colour[c] for c in order], alpha=0.5 if fc else 0.85,
                   edgecolor=[colour[c] for c in order] if fc else "white", linewidth=1.6, zorder=3,
                   hatch="////" if fc else None)
        for c in lab:
            right = g[c] > 0.55 * ax.get_xlim()[1] if False else g[c] > 0.55 * xl[1]
            ax.annotate(label_of(c), (g[c], k[c]), xytext=(-7, 4) if right else (7, 4), textcoords="offset points",
                        fontsize=8.5, color="#222222", zorder=5, ha="right" if right else "left")
        ax.set_xscale("log"); ax.set_yscale("log")
        ax.set_xlim(*xl); ax.set_ylim(*yl)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}k" if v >= 1 else f"${v * 1000:,.0f}"))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        ax.set_xlabel("GDP per head, PPP (constant 2021 international $, log scale)")
        ax.set_ylabel("Electricity demand per head (kWh per year, log scale)")
        ax.set_title(f"{continent}: GDP per head vs electricity demand per head, 2000-2035 (bubble size = total demand)",
                     loc="left", fontsize=13)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        ax.grid(color="#e8e8e8", lw=0.6); ax.set_axisbelow(True)
        h = [Line2D([], [], marker="o", ls="", ms=9, mfc=PAL[j], mec="white", label=n)
             for j, (n, m) in enumerate(groups) if any(c in ok for c in m)]
        h += [Line2D([], [], marker="o", ls="", ms=9, mfc="#D2D2D2", mec="white", label="Rest of world"),
              Line2D([], [], color="#555555", lw=1.6, label="World fit this year (saturating curve)"),
              Line2D([], [], color="#555555", lw=1, ls=":", label="World fit 2024 (forecasts move along it)"),
              Line2D([], [], color="#777777", lw=1.2, ls="--", label="Forecast path (hatched bubbles)")]
        ax.legend(handles=h, loc="upper left", frameon=False, fontsize=9)

    fig.text(0.01, 0.045, f"Source: Ember yearly electricity data (demand); World Bank (GDP PPP, population). {len(ok)} "
             "countries with data in every key year; over 0.3 m people.\nForecast: IMF WEO GDP and population (2029-31 "
             "rates held after 2031); demand per head grows at 50% own 2014-24 elasticity + 50% the 2024 curve's slope "
             "(backtest 2015-24: median error 13%).",
             fontsize=8, color="#6B6B6B")
    fig.text(0.01, 0.012, f"{continent}: {len(hi)} countries; the {len(lab)} largest by demand are labelled with "
             "their path through the key years. Motion between key years is interpolated.", fontsize=8,
             color="#6B6B6B")
    fig.subplots_adjust(left=0.08, right=0.975, top=0.93, bottom=0.155)
    slug = continent.lower().replace(" & ", "_").replace(" ", "_")
    anim = FuncAnimation(fig, draw, frames=frames, interval=110)
    anim.save(f"{O}/gdp_vs_kwh_bubble_2000_2035_forecast_{slug}.gif", writer=PillowWriter(fps=9), dpi=80)
    anim.save(f"{O}/gdp_vs_kwh_bubble_2000_2035_forecast_{slug}.mp4", writer=FFMpegWriter(fps=9, bitrate=2400), dpi=110)
    draw((len(KEYS) - 1, 0.0)); fig.savefig(f"{O}/fanim_last_{slug}.png", dpi=90)
    plt.close(fig)
    print(continent, len(hi), "countries;", "labelled:", ", ".join(label_of(c) for c in lab), flush=True)


for cont, groups in CONTINENTS.items():
    if not only or cont in only:
        render(cont, groups)
