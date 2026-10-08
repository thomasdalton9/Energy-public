"""
Southeast Asia zoom of the data coverage map. Classification, colours, dots and legend text are imported from
SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP.py (COVERAGE, GAS_PRODUCERS, HYDRO ...), so the two maps cannot disagree; only the
crop (lon 92-142, lat -12 to 29), label positions and callouts for the small countries are specific to this map.
Natural Earth 1:110m via geopandas; Singapore is drawn as a marker (below that resolution), Brunei and Timor-Leste
get callouts. Neighbouring non-SE-Asia land is drawn in light grey and left unlabelled except for the names needed
for orientation.

Usage: python3 SOUTHEAST_ASIA_COVERAGE_MAP.py [--out "output/PNG Charts/southeast_asia_coverage_map.png"]
"""
import argparse
import datetime
import os
import sys
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP as base  # noqa: E402

import geopandas as gpd  # noqa: E402
import matplotlib  # noqa: E402
from shapely.geometry import MultiPolygon  # noqa: E402

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = base.ROOT
SE_ASIA = ["Thailand", "Vietnam", "Philippines", "Indonesia", "Malaysia", "Singapore", "Myanmar", "Cambodia", "Laos",
           "Brunei", "Timor-Leste"]
XMIN, XMAX, YMIN, YMAX = 92, 142, -12, 29
# in-country labels (lon, lat); dots sit below/beside them
LABEL_AT = {"Thailand": (101.0, 16.6), "Myanmar": (96.3, 21.8), "Laos": (102.9, 19.9), "Indonesia": (114.5, -1.2)}
# callouts: name -> (label lon, label lat, target lon, target lat)
CALLOUTS = {"Vietnam": (112.5, 17.0, 108.3, 15.5), "Cambodia": (103.0, 8.3, 104.5, 12.6),
            "Malaysia": (97.0, 5.0, 101.9, 4.2), "Brunei": (118.5, 7.6, 114.7, 4.6),
            "Timor-Leste": (130.5, -10.6, 125.9, -8.8),
            "Philippines": (132.0, 15.5, 123.0, 15.0), "Singapore": (99.5, -1.5, 103.82, 1.35)}
DOT_AT = {"Thailand": (100.3, 14.3), "Myanmar": (95.8, 19.8), "Indonesia": (114.5, -3.2), "Philippines": (122.9, 11.3),
          "Vietnam": (106.4, 11.0), "Malaysia": (102.2, 3.0), "Brunei": (114.7, 4.4)}
HYDRO_AT = {"Thailand": (102.4, 14.3), "Philippines": (124.8, 11.3)}
CONTEXT = {"China": (104.0, 26.5), "India": (94.5, 26.8), "Bangladesh": (94.2, 23.6), "Australia": (133.0, -11.2)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "PNG Charts", "southeast_asia_coverage_map.png"))
    args = ap.parse_args()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    world = world.cx[XMIN:XMAX, YMIN:YMAX].copy()
    cov = base.COVERAGE
    world["fill"] = world["name"].map(lambda n: base.FILL[cov[n][0]] if n in SE_ASIA and n in cov else base.OUT)

    fig, ax = plt.subplots(figsize=(13, 10.5), dpi=150)
    ax.set_facecolor("#F7FAFD")
    world.plot(ax=ax, color=world["fill"], edgecolor=base.EDGE, linewidth=0.6)
    my = world[world["name"] == "Malaysia"].geometry.iloc[0]
    parts = list(my.geoms) if isinstance(my, MultiPolygon) else [my]
    east = [p for p in parts if p.centroid.x > 105]
    if east:
        gpd.GeoSeries(east).plot(ax=ax, color=base.FILL[base.SABAH_SARAWAK], edgecolor=base.EDGE, linewidth=0.6)
        ax.text(113.6, 3.0, "Sabah &\nSarawak\n(Ember)", fontsize=7.5, color="#333333", ha="center", va="center",
                fontweight="bold")
    ax.set_xlim(XMIN, XMAX)
    ax.set_ylim(YMIN, YMAX)
    ax.set_aspect("equal")
    ax.set_xticks([])
    ax.set_yticks([])
    for s in ax.spines.values():
        s.set_visible(False)

    for name, (x, y) in CONTEXT.items():
        ax.text(x, y, name, fontsize=8, color="#999999", ha="center", style="italic")

    for _, row in world.iterrows():
        name = row["name"]
        if name not in SE_ASIA or name not in cov or name == "Singapore":
            continue
        p = row.geometry.representative_point()
        if name in CALLOUTS:
            lx, ly, tx, ty = CALLOUTS[name]
            label = "Malaysia (Peninsular)" if name == "Malaysia" else name
            ax.annotate(label, xy=(tx, ty), xytext=(lx, ly), fontsize=9, color="#333333", ha="center",
                        fontweight="bold", arrowprops=dict(arrowstyle="-", color="#888888", lw=0.7))
        else:
            px, py = LABEL_AT.get(name, (p.x, p.y))
            dark = cov[name][0] == "amber"
            ax.text(px, py, name, fontsize=10.5, color="#222222" if dark else "white", ha="center", va="center",
                    fontweight="bold")
        dx, dy = DOT_AT.get(name, (p.x, p.y - 1.5))
        if name in base.GAS_PRODUCERS:
            ax.plot(dx, dy, "o", markersize=10, color=base.DOT_HAVE if base.GAS_PRODUCERS[name] else base.DOT_MISSING,
                    markeredgecolor="white", markeredgewidth=1.2, zorder=5)
        if name in base.HYDRO:
            hx, hy = HYDRO_AT.get(name, (dx, dy))
            ax.plot(hx, hy, "o", markersize=10, color=base.DOT_HYDRO, markeredgecolor=base.DOT_HAVE,
                    markeredgewidth=0.8, zorder=5)
    # Timor-Leste is tiny at 1:110m: add a marker so it is visible
    ax.plot(125.9, -8.8, "s", markersize=7, color=base.FILL[cov["Timor-Leste"][0]], markeredgecolor="#333333",
            markeredgewidth=0.8, zorder=6)
    sx, sy = base.SINGAPORE
    ax.plot(sx, sy, "o", markersize=10, color=base.FILL[cov["Singapore"][0]], markeredgecolor="#333333",
            markeredgewidth=0.9, zorder=6)
    lx, ly, tx, ty = CALLOUTS["Singapore"]
    ax.annotate("Singapore", xy=(tx, ty), xytext=(lx, ly), fontsize=9, color="#333333", ha="center", fontweight="bold",
                arrowprops=dict(arrowstyle="-", color="#888888", lw=0.7))

    handles = [Patch(facecolor=base.GREEN, label="Good coverage: gas demand by sector,\ngas production/supply and\n"
                                                 "power generation by type (raw)"),
               Patch(facecolor=base.BLUE, label="Partial: power generation by type from\nthe grid operator, or a gas demand split"),
               Patch(facecolor=base.AMBER, label="Ember only: power generation by type\n(compiled monthly or yearly, fallback)"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=base.DOT_HAVE, markeredgecolor="white",
                      label="Gas producer: domestic production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=base.DOT_MISSING, markeredgecolor="white",
                      label="Gas producer: no production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=9, color=base.DOT_HYDRO,
                      markeredgecolor=base.DOT_HAVE, label="Hydro reservoir / dam-level data"),
               Patch(facecolor=base.OUT, label="Outside Southeast Asia scope")]
    ax.legend(handles=handles, loc="upper left", frameon=False, fontsize=9.5, bbox_to_anchor=(0.0, -0.01), ncol=2,
              labelspacing=1.0, columnspacing=3.0)
    ax.set_title("Southeast Asia data coverage", fontsize=16, fontweight="bold", loc="left")
    import textwrap
    order = {"green": 0, "blue": 1, "amber": 2}
    lines = []
    for name in sorted((n for n in SE_ASIA if n in cov), key=lambda n: (order[cov[n][0]], SE_ASIA.index(n))):
        what = cov[name][1]
        if name == "Malaysia":
            what += "; Sabah + Sarawak: Ember national total minus GSO (estimate)"
        wrapped = textwrap.wrap(f"{name}: {what}", 165, subsequent_indent="    ")
        lines.append("\u2022 " + wrapped[0])
        lines += ["   " + w for w in wrapped[1:]]
    lines += ["", f"Coverage as of {datetime.date.today():%d %B %Y}. Classification as in SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP.py. "
                  "Map: Natural Earth 1:110m."]
    ax.text(0.0, -0.285, "\n".join(lines), fontsize=8.3, color="#444444", transform=ax.transAxes, va="top",
            linespacing=1.45)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fig.savefig(args.out, bbox_inches="tight", facecolor="white")
    print(f"Saved {args.out}")


if __name__ == "__main__":
    main()
