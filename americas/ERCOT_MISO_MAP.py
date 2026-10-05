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
import sys
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
CLEAN = "--clean" in sys.argv  # no title, no inset: every plant labelled on the main map
OUT = os.path.join(ROOT, "output", "PNG Charts", "ercot_miso_lng_map_clean.png" if CLEAN else "ercot_miso_lng_map.png")

XLIM, YLIM = (-108.5, -68.0), (21.0, 49.5)
MISO_C, ERCOT_C = "#2a6fb0", "#d9822b"
INK, MUTED = "#222222", "#666666"
GREEN = "#1b7f5c"
INSET = (-98.0, -88.5, 26.7, 31.0)  # lon0, lon1, lat0, lat1 of the Gulf Coast inset

# name, lon, lat, Bcf/d, status, main-map label x, y, ha (None = part of the Sabine-Cameron cluster, labelled in the inset),
# inset label x, y, ha
LNG = [
    ("Cameron", -93.33, 29.79, 2.0, "op", None, None, None, -91.7, 28.6, "left"),
    ("Calcasieu Pass", -93.34, 29.76, 1.6, "op", None, None, None, -92.9, 27.9, "center"),
    ("Port Arthur LNG", -93.95, 29.84, 1.9, "uc", None, None, None, -95.6, 30.65, "right"),
    ("Golden Pass", -93.92, 29.74, 2.4, "uc", None, None, None, -95.6, 30.05, "right"),
    ("Sabine Pass", -93.87, 29.72, 4.5, "op", None, None, None, -94.2, 27.6, "center"),
    ("Plaquemines", -89.96, 29.60, 3.4, "op", -90.0, 27.3, "right", -89.7, 28.3, "center"),
    ("Freeport", -95.31, 28.94, 2.1, "op", -95.0, 27.8, "left", -95.9, 27.9, "center"),
    ("Corpus Christi\n(Stage 3 under constr.)", -97.25, 27.87, 2.4, "op", -96.1, 26.2, "left", -97.9, 27.05, "left"),
    ("Rio Grande LNG", -97.20, 25.97, 3.6, "uc", -94.8, 23.4, "left", None, None, None),
    ("Cove Point", -76.39, 38.40, 0.8, "op", -77.0, 36.9, "right", None, None, None),
    ("Elba Island", -80.99, 32.08, 0.35, "op", -80.2, 30.6, "center", None, None, None),
]
# main map shows the five Sabine-Cameron plants as one grouped marker
CLUSTER = (-93.7, 29.77, "Sabine-Cameron\ncluster (5 plants,\nsee inset)", -91.0, 28.95, "right")
# name, lon, lat, main label x, y, ha, inset label x, y, ha
# Waha gas hub sits near Waha, Reeves/Pecos Co. (about 31.3N, 103.2W), just outside the ERCOT polygon
# --clean layout: main-map label x, y, ha for every plant (the five Sabine-Cameron plants labelled in the Gulf)
CLEAN_LABELS = {  # east-most plant gets the top label, so the leader lines fan out without crossing
    "Cameron": (-91.0, 27.0, "left"), "Calcasieu Pass": (-91.0, 26.0, "left"),
    "Sabine Pass": (-91.0, 25.0, "left"), "Golden Pass": (-91.0, 24.0, "left"), "Port Arthur LNG": (-91.0, 23.0, "left"),
    "Plaquemines": (-88.0, 28.4, "left"), "Rio Grande LNG": (-98.5, 22.6, "left"),
}
HUBS = [
    ("Henry Hub\n(Erath, LA)", -92.04, 30.13, -90.4, 32.2, "left", -91.2, 30.55, "left"),
    ("Waha\n(Reeves / Pecos Co., TX)", -103.2, 31.33, -104.3, 28.9, "center", None, None, None),
]


def cap_txt(cap):
    return f"{cap + 1e-9:.1f} Bcf/d"


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

    fig, ax = plt.subplots(figsize=(12.6, 11.9), dpi=150)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    def base(a, miso_lw=1.4):
        for f in adm:
            draw(a, f["geometry"], facecolor="#f4f1ea", edgecolor="#b9b4a8", lw=0.5, zorder=1)
        draw(a, geo["miso"], facecolor=MISO_C, edgecolor=MISO_C, alpha=0.35, lw=miso_lw, zorder=2)
        draw(a, geo["ercot"], facecolor=ERCOT_C, edgecolor=ERCOT_C, alpha=0.40, lw=1.4, zorder=3)

    base(ax)
    ax.text(-91.5, 43.2, "MISO", color=MISO_C, fontsize=26, fontweight="bold", ha="center", alpha=0.9, zorder=4)
    ax.text(-99.3, 32.0, "ERCOT", color="#a85d12", fontsize=22, fontweight="bold", ha="center", zorder=4)

    halo = [pe.withStroke(linewidth=2.5, foreground="white")]
    lead = dict(arrowstyle="-", color=MUTED, lw=0.7, shrinkA=0, shrinkB=3)

    def marker(a, x, y, cap, op, scale, z=6):
        a.scatter([x], [y], s=scale[0] + cap * scale[1], marker="o", facecolor=GREEN if op else "white",
                  edgecolor=GREEN, lw=1.6, zorder=z if op else z + 1)

    # ---- main map
    for name, x, y, cap, st, lx, ly, ha, *_ in LNG:
        if CLEAN and name in CLEAN_LABELS:
            lx, ly, ha = CLEAN_LABELS[name]
        if lx is None:
            continue
        marker(ax, x, y, cap, st == "op", (30, 45))
        mname = name.replace("(Stage 3 under constr.)", "(Stage 3 under\nconstr.)")
        sep = "  " if CLEAN and name in CLEAN_LABELS else "\n"
        ax.annotate(f"{mname}{sep}{cap_txt(cap)}", (x, y), (lx, ly), fontsize=8.5, color=INK, ha=ha, va="center",
                    zorder=7, path_effects=halo, arrowprops=lead)
    cx, cy, ctxt, clx, cly, cha = CLUSTER
    if not CLEAN:
        ax.scatter([cx], [cy], s=170, marker="o", facecolor=GREEN, edgecolor="white", lw=1.2, zorder=6)
        ax.text(cx, cy, "5", color="white", fontsize=8, fontweight="bold", ha="center", va="center", zorder=7)
        ax.annotate(ctxt, (cx, cy), (clx, cly), fontsize=8.5, color=INK, ha=cha, va="center", zorder=7,
                    path_effects=halo, arrowprops=lead)
    for name, x, y, lx, ly, ha, *_ in HUBS:
        ax.scatter([x], [y], s=150, marker="*", facecolor="#c0262d", edgecolor="white", lw=0.8, zorder=8)
        ax.annotate(name, (x, y), (lx, ly), fontsize=9.5, fontweight="bold", color="#c0262d", path_effects=halo,
                    ha=ha, va="center", zorder=8,
                    arrowprops=dict(arrowstyle="-", color="#c0262d", lw=0.8, shrinkA=0, shrinkB=5))
    ax.text(-108.3, 32.5, "ERCOT footprint simplified\n(HIFLD polygon); Waha lies\nat/just outside its western edge",
            fontsize=7.2, color=MUTED, ha="left", va="bottom", style="italic", path_effects=halo, zorder=8)

    # ---- Gulf Coast inset (lower right, over the Atlantic / Southeast)
    x0, x1, y0, y1 = INSET
    iw = 21.5
    ih = iw * (y1 - y0) / (x1 - x0) * 0.82  # same aspect as the main map
    ins = ax.inset_axes([-89.5, 21.2, iw, ih], transform=ax.transData, zorder=20)
    ins.set_visible(not CLEAN)
    ins.set_facecolor("#e8f1f8")
    base(ins, miso_lw=1.0)
    ins.set_xlim(x0, x1)
    ins.set_ylim(y0, y1)
    ins.set_aspect(1 / 0.82)
    ins.set_xticks([])
    ins.set_yticks([])
    for sp in ins.spines.values():
        sp.set_visible(False)
    ins.text(x1 - 0.1, y0 + 0.1, "Gulf Coast LNG inset (zoom)", fontsize=9, fontweight="bold", color=INK, va="bottom", ha="right",
             path_effects=halo, zorder=9)
    for name, x, y, cap, st, _lx, _ly, _ha, ilx, ily, iha in LNG:
        if ilx is None:
            continue
        marker(ins, x, y, cap, st == "op", (22, 12), z=6)
        nm = name.replace("\n", " ")
        ins.annotate(f"{nm}\n{cap_txt(cap)}", (x, y), (ilx, ily), fontsize=8, color=INK, ha=iha, va="center",
                     zorder=8, path_effects=halo, arrowprops=lead, annotation_clip=False)
    for name, x, y, _lx, _ly, _ha, ilx, ily, iha in HUBS:
        if ilx is None:
            continue
        ins.scatter([x], [y], s=110, marker="*", facecolor="#c0262d", edgecolor="white", lw=0.8, zorder=9)
        ins.annotate(name, (x, y), (ilx, ily), fontsize=8.5, fontweight="bold", color="#c0262d",
                     path_effects=halo, ha=iha, va="center", zorder=9,
                     arrowprops=dict(arrowstyle="-", color="#c0262d", lw=0.8, shrinkA=0, shrinkB=4))

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
    ax.legend(handles=handles, loc="lower left", bbox_to_anchor=(0.0, 0.0), frameon=True, facecolor="white",
              edgecolor="none", framealpha=1.0, fontsize=9, borderpad=0.8,
              title="Marker size scales with\nnameplate capacity (Bcf/d)", title_fontsize=8.5, alignment="left")

    if not CLEAN:
        fig.suptitle("ERCOT and MISO footprints, US LNG export plants and the Henry Hub and Waha pricing hubs",
                     x=0.02, y=0.975, ha="left", fontsize=15, fontweight="bold", color=INK)
    fig.text(0.02, 0.012,
             "Sources: footprints = HIFLD Control Areas (balancing-authority polygons, simplified); states = Natural Earth 1:50m.\n"
             "LNG plants = EIA US liquefaction capacity / export terminals and FERC North American LNG terminal lists; capacities are approximate\n"
             "nameplate (Bcf/d), status as listed there, coordinates rounded to 0.01 deg. Hubs: Henry Hub, Erath LA; Waha, Reeves/Pecos County TX.",
             ha="left", va="bottom", fontsize=7.5, color=MUTED)
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99 if CLEAN else 0.945, bottom=0.085)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    fig.savefig(OUT, facecolor="white")
    print("saved", OUT)


if __name__ == "__main__":
    main()
