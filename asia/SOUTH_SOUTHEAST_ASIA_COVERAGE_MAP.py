"""
Map of South and Southeast Asia (the countries in SOUTH_SOUTHEAST_ASIA_MASTER.py) coloured by how much of each
country's gas and power data this repo pulls, in the style of south_america/COVERAGE_MAP.py:

  green  - good coverage: gas demand by sector, gas production/supply and power generation by type from a raw source
           (production counts as covered where the country produces no gas)
  blue   - partial: power generation by type from the grid operator, or a gas demand split, but not the full set
  amber  - Ember only: power generation by type from Ember's compiled monthly or yearly data (fallback, not raw)
  grey   - no data pulled (no country in scope at present; kept for future additions)
Dots mark gas-producing countries (dark blue: we pull domestic production; red: produces gas, no production data)
and countries with hydro reservoir data (light blue). Malaysia is split: the GSO feed covers Peninsular Malaysia
only, so Sabah and Sarawak are coloured as Ember-only.

The classification below is maintained by hand: update COVERAGE when a pull is added. Country shapes are Natural
Earth 1:110m (bundled with geopandas 0.14); Singapore is too small for that scale and is drawn as a marker.

Usage: python3 SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP.py [--out "output/PNG Charts/south_southeast_asia_coverage_map.png"]
"""
import argparse
import os
import warnings

import geopandas as gpd
import matplotlib
from shapely.geometry import MultiPolygon

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

GREEN, BLUE, AMBER, GREY, OUT, EDGE = "#1BAF7A", "#2A78D6", "#EDA100", "#C9C9C9", "#F0F0F0", "#FFFFFF"
FILL = {"green": GREEN, "blue": BLUE, "amber": AMBER}

# Natural Earth name -> (class, what we have)
COVERAGE = {
    "India": ("green", "Gas demand by sector, production, LNG imports (PPAC); power: CEA conventional plants daily "
                       "+ Ember monthly; hydro reservoirs (CEA); IEX prices, coal stocks"),
    "Bangladesh": ("green", "Gas production by company, LNG, demand split (Petrobangla daily); power PGCB"),
    "Thailand": ("green", "Gas supply by field, Myanmar pipeline, LNG, demand by sector (EPPO/PTT); power EPPO; "
                          "reservoirs RID"),
    "Singapore": ("green", "Gas demand by sector, pipeline vs LNG imports (EMA); power EMC/NEMS; no production"),
    "Sri Lanka": ("blue", "Power PUCSL (no gas market)"),
    "Bhutan": ("blue", "Power BPSO (no gas)"),
    "Malaysia": ("blue", "Power GSO, Peninsular Malaysia only; Single Buyer prices; no gas data"),
    "Pakistan": ("blue", "Power: CPPA-G / NEPRA monthly filings (national grid); gas production by province (PBS, to Mar 2024); "
                         "reservoirs IRSA/WAPDA"),
    "Vietnam": ("blue", "Power: EVN / NSMO daily (from May 2023); gas: NSO monthly tables (from 2015)"),
    "Philippines": ("blue", "Power: IEMOP plant dispatch by fuel (Ember shown in the master until a year of history); DOE capacity; "
                             "prices; PAGASA dam levels; no gas data"),
    "Indonesia": ("blue", "Gas production by contractor + use by sector (Ditjen Migas); power: Ember yearly"),
    "Myanmar": ("amber", "Power: Ember yearly"),
    "Cambodia": ("blue", "Power: EAC annual reports (raw, annual rows); EAC capacity; no gas"),
    "Laos": ("amber", "Power: Ember yearly"),
    "Brunei": ("amber", "Power: Ember yearly"),
    "Nepal": ("blue", "Power: NEA load dispatch daily reports (Apr 2023 - Jan 2025, gap Feb 2025 - Sep 2026, filled from Ember); no gas"),
    "Timor-Leste": ("amber", "Power: Ember yearly"),
}
SABAH_SARAWAK = "amber"     # Malaysia east of 105E: no GSO feed, Ember (national) only
GAS_PRODUCERS = {"India": True, "Bangladesh": True, "Thailand": True, "Pakistan": True, "Malaysia": False,
                 "Indonesia": True, "Myanmar": False, "Brunei": False, "Vietnam": True, "Philippines": False}
HYDRO = {"India", "Thailand", "Philippines", "Pakistan", "Sri Lanka"}   # CEA, RID, PAGASA, IRSA/WAPDA, PUCSL
DOT_HAVE, DOT_MISSING, DOT_HYDRO = "#0B3A66", "#E34948", "#8FD3FF"
SINGAPORE = (103.82, 1.35)
OUT_OF_SCOPE = {"China", "Taiwan", "Afghanistan", "Iran", "Tajikistan", "Turkmenistan", "Uzbekistan", "Kyrgyzstan",
                "Oman", "United Arab Emirates", "Papua New Guinea", "Australia", "Japan", "South Korea", "North Korea"}
# label position overrides (lon, lat) for in-country labels
LABEL_AT = {"India": (78.5, 21.5), "Indonesia": (114.0, -1.0), "Myanmar": (95.8, 21.5),
            "Thailand": (100.9, 15.6), "Pakistan": (69.5, 28.5), "Laos": (103.0, 19.9)}
# small countries: label beside them with a leader line -> (label lon, label lat)
CALLOUTS = {"Sri Lanka": (85.5, 5.0), "Bhutan": (92.5, 30.0), "Nepal": (82.0, 31.5), "Bangladesh": (86.5, 19.5),
            "Cambodia": (100.0, 9.0), "Vietnam": (110.5, 17.5), "Brunei": (110.0, 9.0), "Timor-Leste": (128.5, -12.0),
            "Malaysia": (109.0, 6.3), "Philippines": (132.0, 17.0)}
DOT_AT = {"India": (78.5, 19.4), "Indonesia": (114.0, -3.0), "Malaysia": (101.9, 3.2), "Myanmar": (95.8, 19.6),
          "Thailand": (100.3, 14.0), "Pakistan": (69.5, 26.6), "Philippines": (122.8, 10.6), "Bangladesh": (90.2, 23.8),
          "Vietnam": (106.6, 11.2), "Brunei": (114.7, 4.4)}
CALLOUT_TEXT = {"Malaysia": "Malaysia (Peninsular)"}
CALLOUT_XY = {"Malaysia": (102.3, 4.5), "Philippines": (121.5, 16.0)}
HYDRO_AT = {"India": (80.5, 19.4), "Thailand": (102.2, 14.0), "Philippines": (124.6, 10.6)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "PNG Charts", "south_southeast_asia_coverage_map.png"))
    args = ap.parse_args()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    xmin, xmax, ymin, ymax = 60, 142, -12.5, 37.5
    world = world.cx[xmin:xmax, ymin:ymax].copy()

    def colour(name):
        if name in COVERAGE:
            return FILL[COVERAGE[name][0]]
        return OUT if name in OUT_OF_SCOPE else GREY

    world["fill"] = world["name"].map(colour)
    fig, ax = plt.subplots(figsize=(14, 9.5), dpi=150)
    world.plot(ax=ax, color=world["fill"], edgecolor=EDGE, linewidth=0.6)
    # Malaysia: recolour Sabah + Sarawak (polygons east of 105E)
    my = world[world["name"] == "Malaysia"].geometry.iloc[0]
    parts = list(my.geoms) if isinstance(my, MultiPolygon) else [my]
    east = [p for p in parts if p.centroid.x > 105]
    if east:
        gpd.GeoSeries(east).plot(ax=ax, color=FILL[SABAH_SARAWAK], edgecolor=EDGE, linewidth=0.6)
        ax.text(114.6, 2.6, "Sabah &\nSarawak", fontsize=7.5, color="#333333", ha="center", va="center", fontweight="bold")
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_axis_off()

    for _, row in world.iterrows():
        name = row["name"]
        if name not in COVERAGE:
            continue
        p = row.geometry.representative_point()
        px, py = LABEL_AT.get(name, (p.x, p.y))
        if name in CALLOUTS:
            lx, ly = CALLOUTS[name]
            ax.annotate(CALLOUT_TEXT.get(name, name), xy=CALLOUT_XY.get(name, (p.x, p.y)), xytext=(lx, ly), fontsize=8.5,
                        color="#333333", ha="center",
                        arrowprops=dict(arrowstyle="-", color="#888888", lw=0.6))
        else:
            dark = COVERAGE[name][0] == "amber"
            ax.text(px, py, name, fontsize=9.5, color="#222222" if dark else "white", ha="center", va="center",
                    fontweight="bold")
        if name in GAS_PRODUCERS:
            dx, dy = DOT_AT.get(name, (px, py - 1.8))
            ax.plot(dx, dy, "o", markersize=9, color=DOT_HAVE if GAS_PRODUCERS[name] else DOT_MISSING,
                    markeredgecolor="white", markeredgewidth=1.2, zorder=5)
        if name in HYDRO:
            hx, hy = HYDRO_AT.get(name, (px, py - 1.8))
            ax.plot(hx, hy, "o", markersize=9, color=DOT_HYDRO, markeredgecolor="#0B3A66", markeredgewidth=0.8, zorder=5)
    # Singapore: below Natural Earth 1:110m resolution - a marker in its coverage colour
    ax.plot(*SINGAPORE, "o", markersize=9, color=FILL[COVERAGE["Singapore"][0]], markeredgecolor="#333333",
            markeredgewidth=0.8, zorder=6)
    ax.annotate("Singapore", xy=SINGAPORE, xytext=(97.0, -3.5), fontsize=8.5, color="#333333", ha="center",
                arrowprops=dict(arrowstyle="-", color="#888888", lw=0.6))
    for txt, (x, y) in {"China (not in scope)": (100.0, 32.5)}.items():
        ax.text(x, y, txt, fontsize=8, color="#999999", ha="center")

    handles = [Patch(facecolor=GREEN, label="Good coverage: gas demand by sector,\ngas production/supply and\n"
                                            "power generation by type (raw)"),
               Patch(facecolor=BLUE, label="Partial: power generation by type from\nthe grid operator, or a gas demand split"),
               Patch(facecolor=AMBER, label="Ember only: power generation by type\n(compiled monthly or yearly, fallback)"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=DOT_HAVE, markeredgecolor="white",
                      label="Gas producer: domestic production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=DOT_MISSING, markeredgecolor="white",
                      label="Gas producer: no production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=DOT_HYDRO, markeredgecolor="#0B3A66",
                      label="Hydro reservoir / dam-level data")]
    ax.legend(handles=handles, loc="upper left", frameon=False, fontsize=9.5, bbox_to_anchor=(0.0, -0.01), ncol=2,
              labelspacing=1.0, columnspacing=3.0)
    ax.set_title("South and Southeast Asia: Gas and Power Data Coverage", fontsize=15, fontweight="bold", loc="left")
    import datetime
    import textwrap
    order = {"green": 0, "blue": 1, "amber": 2}
    notes = []
    for name in sorted(COVERAGE, key=lambda n: (order[COVERAGE[n][0]], n)):
        what = COVERAGE[name][1]
        if name == "Malaysia":
            what += "; Sabah + Sarawak: Ember national total minus GSO (estimate)"
        wrapped = textwrap.wrap(f"{name}: {what}", 190, subsequent_indent="    ")
        notes.append("\u2022 " + wrapped[0])
        notes += ["   " + w for w in wrapped[1:]]
    notes += ["", f"Coverage as of {datetime.date.today():%d %B %Y}. Map: Natural Earth 1:110m."]
    ax.text(0.0, -0.30, "\n".join(notes), fontsize=8.3, color="#444444", transform=ax.transAxes, va="top",
            linespacing=1.45)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fig.savefig(args.out, bbox_inches="tight", facecolor="white")
    print(f"Saved {args.out}")
    for name, (cls, what) in sorted(COVERAGE.items(), key=lambda kv: (["green", "blue", "amber"].index(kv[1][0]), kv[0])):
        print(f"  {cls:5s} {name}: {what}")


if __name__ == "__main__":
    main()
