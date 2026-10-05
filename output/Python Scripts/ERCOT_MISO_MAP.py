"""
Map of the ERCOT and MISO footprints, US LNG export plants and the Henry Hub / Waha pricing hubs.

Boundaries: HIFLD "Control Areas" (balancing-authority) polygons, cached in maps/us_balancing_authorities.geojson
(fetched once from HIFLD's public ArcGIS service by discovery_archive/FETCH_US_BALANCING_AUTHORITIES.py; simplified
to 0.05 deg). State outlines: Natural Earth 1:50m admin-1, cached in maps/north_america_admin1.geojson.
No network access is needed to redraw the map.

LNG plants: nameplate (peak/liquefaction) capacity in Bcf/d and status from EIA "U.S. liquefaction capacity" /
"Natural gas export terminals" tables and FERC "North American LNG Export Terminals" lists; site coordinates are the
terminal locations (rounded to 0.01 deg). Capacities are approximate, rounded to 0.1.

Output: output/PNG Charts/ercot_miso_lng_map.png
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.lines import Line2D
from matplotlib.path import Path
from matplotlib.patches import PathPatch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BA = os.path.join(ROOT, "maps", "us_balancing_authorities.geojson")
ADMIN = os.path.join(ROOT, "maps", "north_america_admin1.geojson")
OUT = os.path.join(ROOT, "output", "PNG Charts", "ercot_miso_lng_map.png")

XLIM, YLIM = (-108.5, -75.0), (24.0, 49.5)
MISO_C, ERCOT_C = "#2a6fb0", "#d9822b"
INK, MUTED = "#222222", "#666666"

# name, lon, lat, Bcf/d, status, label x, label y, ha (label positions in map degrees; labels sit offshore with leader lines)
LNG = [
    ("Cameron", -93.33, 29.79, 2.0, "op", -92.3, 28.9, "left"),
    ("Calcasieu Pass", -93.34, 29.76, 1.6, "op", -92.3, 28.2, "left"),
    ("Port Arthur LNG", -93.95, 29.84, 1.9, "uc", -92.3, 27.5, "left"),
    ("Golden Pass", -93.92, 29.74, 2.4, "uc", -92.3, 26.8, "left"),
    ("Sabine Pass", -93.87, 29.72, 4.5, "op", -92.3, 26.1, "left"),
    ("Plaquemines", -89.96, 29.60, 3.4, "op", -87.6, 28.7, "left"),
    ("Freeport", -95.31, 28.94, 2.1, "op", -95.2, 27.2, "center"),
    ("Corpus Christi\n(Stage 3 under constr.)", -97.25, 27.87, 2.4, "op", -97.9, 26.9, "right"),
    ("Rio Grande LNG", -97.20, 25.97, 3.6, "uc", -95.8, 25.3, "left"),
    ("Cove Point", -76.39, 38.40, 0.8, "op", -77.0, 36.9, "right"),
    ("Elba Island", -80.99, 32.08, 0.35, "op", -80.2, 30.6, "center"),
]
HUBS = [
    ("Henry Hub\n(Erath, LA)", -92.04, 30.13, -90.4, 32.2, "left"),
    ("Waha\n(Reeves / Pecos Co., TX)", -103.15, 31.0, -104.3, 28.9, "center"),
]


def rings(geom):
    polys = [geom["coordinates"]] if geom["type"] == "Polygon" else geom["coordinates"]
    for poly in polys:
        yield poly


def draw(ax, geom, **kw):
    for poly in rings(geom):
        verts, codes = [], []
        for ring in poly:
            verts += ring
            codes += [Path.MOVETO] + [Path.LINETO] * (len(ring) - 2) + [Path.CLOSEPOLY]
        ax.add_patch(PathPatch(Path(verts, codes), **kw))


def main():
    ba = json.load(open(BA))["features"]
    adm = json.load(open(ADMIN))["features"]
    pick = {"MIDCONTINENT INDEPENDENT": "miso", "ELECTRIC RELIABILITY COUNCIL OF TEXAS": "ercot"}
    geo = {}
    for f in ba:
        for k, v in pick.items():
            if f["properties"]["NAME"].startswith(k):
                geo[v] = f["geometry"]
    assert set(geo) == {"miso", "ercot"}, geo.keys()

    fig, ax = plt.subplots(figsize=(11.5, 11.3), dpi=150)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    for f in adm:
        draw(ax, f["geometry"], facecolor="#f4f1ea", edgecolor="#b9b4a8", lw=0.5, zorder=1)
    draw(ax, geo["miso"], facecolor=MISO_C, edgecolor=MISO_C, alpha=0.35, lw=1.4, zorder=2)
    draw(ax, geo["ercot"], facecolor=ERCOT_C, edgecolor=ERCOT_C, alpha=0.40, lw=1.4, zorder=3)
    ax.text(-91.5, 43.2, "MISO", color=MISO_C, fontsize=26, fontweight="bold", ha="center", alpha=0.9, zorder=4)
    ax.text(-99.3, 32.0, "ERCOT", color="#a85d12", fontsize=22, fontweight="bold", ha="center", zorder=4)

    halo = [pe.withStroke(linewidth=2.5, foreground="white")]
    for name, x, y, cap, st, lx, ly, ha in LNG:
        op = st == "op"
        ax.scatter([x], [y], s=30 + cap * 45, marker="o", facecolor="#1b7f5c" if op else "white",
                   edgecolor="#1b7f5c", lw=1.8, zorder=6 if op else 7)
        ax.annotate((f"{name}\n{cap:g} Bcf/d" if "\n" in name or ha != "left" else f"{name}  {cap:g} Bcf/d"), (x, y), (lx, ly), fontsize=8.5, color=INK, ha=ha, va="center",
                    zorder=7, path_effects=halo,
                    arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.7, shrinkA=0, shrinkB=4))
    for name, x, y, lx, ly, ha in HUBS:
        ax.scatter([x], [y], s=150, marker="*", facecolor="#c0262d", edgecolor="white", lw=0.8, zorder=8)
        ax.annotate(name, (x, y), (lx, ly), fontsize=9.5, fontweight="bold", color="#c0262d", path_effects=halo,
                    ha=ha, va="center", zorder=8,
                    arrowprops=dict(arrowstyle="-", color="#c0262d", lw=0.8, shrinkA=0, shrinkB=5))

    ax.set_xlim(*XLIM)
    ax.set_ylim(*YLIM)
    ax.set_aspect(1 / 0.82)
    ax.set_xticks([])
    ax.set_yticks([])
    for s in ax.spines.values():
        s.set_visible(False)

    handles = [
        Line2D([], [], marker="s", ls="", ms=12, mfc=MISO_C, mec=MISO_C, alpha=0.5, label="MISO footprint (balancing area)"),
        Line2D([], [], marker="s", ls="", ms=12, mfc=ERCOT_C, mec=ERCOT_C, alpha=0.55, label="ERCOT footprint (balancing area)"),
        Line2D([], [], marker="o", ls="", ms=9, mfc="#1b7f5c", mec="#1b7f5c", label="LNG export plant, operating"),
        Line2D([], [], marker="o", ls="", ms=9, mfc="white", mec="#1b7f5c", mew=1.8, label="LNG export plant, under construction"),
        Line2D([], [], marker="*", ls="", ms=14, mfc="#c0262d", mec="white", label="Gas pricing hub"),
    ]
    ax.legend(handles=handles, loc="center left", bbox_to_anchor=(0.0, 0.56), frameon=False, fontsize=9,
              title="Marker size scales with\nnameplate capacity (Bcf/d)", title_fontsize=8.5, alignment="left")

    fig.suptitle("ERCOT and MISO footprints, US LNG export plants and the Henry Hub and Waha pricing hubs",
                 x=0.02, y=0.975, ha="left", fontsize=15, fontweight="bold", color=INK)
    fig.text(0.02, 0.012,
             "Sources: footprints = HIFLD Control Areas (balancing-authority polygons, simplified); states = Natural Earth 1:50m.\n"
             "LNG plants = EIA US liquefaction capacity / export terminals and FERC North American LNG terminal lists; capacities are approximate\n"
             "nameplate (Bcf/d), status as listed there, coordinates rounded to 0.01 deg. Hubs: Henry Hub, Erath LA; Waha, Reeves/Pecos County TX.",
             ha="left", va="bottom", fontsize=7.5, color=MUTED)
    fig.subplots_adjust(left=0.01, right=0.99, top=0.945, bottom=0.075)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    fig.savefig(OUT, facecolor="white")
    print("saved", OUT)


if __name__ == "__main__":
    main()
