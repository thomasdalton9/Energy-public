"""
Map of South & Central America (and the Caribbean countries in the master
workbook) coloured by how much of each country's gas and power data this
repo pulls:

  green - good coverage: gas demand by sector, gas production/supply and
          power generation by type (production counts as covered where the
          country produces no gas)
  blue  - power generation by type, or a gas demand split, but not the full set
  grey  - no data pulled

The classification below is maintained by hand: update COVERAGE when a pull
is added. Country shapes are Natural Earth 1:110m (bundled with geopandas
0.14), so no download is needed.

Usage: python3 COVERAGE_MAP.py [--out "output/PNG Charts/south_central_america_coverage_map.png"]
"""
import argparse
import os
import warnings

import geopandas as gpd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

GREEN, BLUE, GREY, OUT, EDGE = "#1BAF7A", "#2A78D6", "#C9C9C9", "#F0F0F0", "#FFFFFF"

# Natural Earth name -> (class, what we have)
COVERAGE = {
    "Argentina": ("green", "Gas demand by sector, production, exports; power CAMMESA"),
    "Brazil": ("green", "Gas demand (MME to Jun-25) + ANP grid flows, supply; power ONS"),
    "Colombia": ("green", "Gas demand by sector, supply by field + LNG; power XM"),
    "Peru": ("green", "Gas demand by sector, production by lot, LNG exports; power COES"),
    "Ecuador": ("green", "Gas by use + LNG imports, Amistad production; power (Ember, CENACE daily from Sep-26)"),
    "Uruguay": ("green", "Gas demand by sector (no domestic production); power ADME"),
    "Bolivia": ("blue", "Gas demand by sector; power CNDC; no production series"),
    "Chile": ("blue", "Power CNE; gas imports only (no demand split)"),
    "Trinidad and Tobago": ("blue", "Gas use by sector + production; no power"),
    "Guatemala": ("blue", "Power AMM"),
    "Honduras": ("blue", "Power ODS (daily from Jun-26, monthly history)"),
    "El Salvador": ("blue", "Power SIGET (monthly); gas for power estimated"),
    "Nicaragua": ("blue", "Power CNDC"),
    "Costa Rica": ("blue", "Power CENCE"),
    "Panama": ("blue", "Power CND; gas for power estimated"),
    "Puerto Rico": ("blue", "Power + gas burn by plant (EIA, monthly)"),
    "Dominican Rep.": ("blue", "Power OC-SENI; gas for power (SIE)"),
    "Jamaica": ("blue", "Power + gas, annual only"),
}
OUT_OF_SCOPE = {"Mexico", "United States of America", "Canada", "Greenland"}
# Small countries get their label beside them, with a leader line: name -> (label lon, label lat)
CALLOUTS = {
    "Trinidad and Tobago": (-56.0, 12.5), "Puerto Rico": (-59.5, 20.5), "Dominican Rep.": (-63.5, 24.5),
    "Jamaica": (-80.0, 15.5), "Chile": (-82.0, -28.0), "El Salvador": (-95.5, 11.0), "Belize": (-94.0, 18.5), "Costa Rica": (-90.5, 7.0),
    "Panama": (-80.0, 4.5), "Honduras": (-81.5, 18.0), "Nicaragua": (-79.0, 13.5), "Guatemala": (-97.5, 15.5),
    "Uruguay": (-49.5, -35.0), "Haiti": (-72.5, 26.5), "Cuba": (-84.0, 26.0), "Guyana": (-55.0, 9.5),
    "Suriname": (-51.0, 7.5), "Ecuador": (-87.0, -2.0), "Paraguay": (-52.0, -24.0),
}
ANCHOR = {"Chile": (-70.6, -28.0)}
LABEL_NAME = {"Dominican Rep.": "Dominican Rep.", "Trinidad and Tobago": "Trinidad & Tobago"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "PNG Charts", "south_central_america_coverage_map.png"))
    args = ap.parse_args()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    xmin, xmax, ymin, ymax = -118, -32, -56, 28
    world = world.cx[xmin:xmax, ymin:ymax].copy()

    def colour(name):
        if name in COVERAGE:
            return {"green": GREEN, "blue": BLUE}[COVERAGE[name][0]]
        return OUT if name in OUT_OF_SCOPE else GREY

    world["fill"] = world["name"].map(colour)
    fig, ax = plt.subplots(figsize=(10, 12), dpi=150)
    world.plot(ax=ax, color=world["fill"], edgecolor=EDGE, linewidth=0.6)
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_axis_off()

    for _, row in world.iterrows():
        name = row["name"]
        if name in OUT_OF_SCOPE or name in ("Falkland Is.", "France", "Bahamas"):
            continue
        p = row.geometry.representative_point()
        if name in ANCHOR:   # long, thin countries: point the label at a visible part
            p = type(p)(*ANCHOR[name])
        text = LABEL_NAME.get(name, name)
        if name in CALLOUTS:
            lx, ly = CALLOUTS[name]
            ax.annotate(text, xy=(p.x, p.y), xytext=(lx, ly), fontsize=7.5, color="#333333", ha="center",
                        arrowprops=dict(arrowstyle="-", color="#888888", lw=0.6))
        else:
            ax.text(p.x, p.y, text, fontsize=8.5, color="white" if name in COVERAGE else "#333333",
                    ha="center", va="center", fontweight="bold")
    ax.text(-104, 24, "Mexico\n(not in scope)", fontsize=7.5, color="#999999", ha="center")
    ax.text(-53.5, 3.4, "Fr. Guiana", fontsize=6.5, color="#333333", ha="center")

    handles = [Patch(facecolor=GREEN, label="Good coverage: gas demand by sector,\ngas production/supply and\n"
                                            "power generation by type"),
               Patch(facecolor=BLUE, label="Partial: power generation by type,\nor a gas demand split"),
               Patch(facecolor=GREY, label="No data")]
    ax.legend(handles=handles, loc="lower left", frameon=False, fontsize=9, bbox_to_anchor=(0.0, 0.06),
              labelspacing=1.0)
    ax.set_title("South and Central America: Gas and Power Data Coverage", fontsize=14,
                 fontweight="bold", loc="left")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fig.savefig(args.out, bbox_inches="tight", facecolor="white")
    print(f"Saved {args.out}")
    for name, (cls, what) in sorted(COVERAGE.items(), key=lambda kv: (kv[1][0] != "green", kv[0])):
        print(f"  {cls:5s} {name}: {what}")


if __name__ == "__main__":
    main()
