"""
Map of North America (USA, Canada, Mexico) coloured by how much gas and
power data this repo pulls. The USA is split by power market (ISO/RTO and
the large utility balancing areas EIA-930 reports), not shown as one block.

  green - good coverage: gas demand by sector, gas production/supply and
          power generation by type
  blue  - power generation by type, or a gas demand split, but not the full set
  grey  - no data pulled
Dots: dark blue = gas producer with production data pulled, red = gas
producer without. Triangles: US LNG export plants whose feedgas is tracked.

COVERAGE / BA_COVERAGE are maintained by hand: update them when a pull is
added. Shapes:
  maps/us_balancing_authorities.geojson  HIFLD Control Areas (simplified),
                                          from discovery_archive/FETCH_US_BALANCING_AUTHORITIES.py
  maps/north_america_admin1.geojson      Natural Earth 1:50m states/provinces (simplified)

Usage: python3 NORTH_AMERICA_COVERAGE_MAP.py [--out "output/PNG Charts/north_america_coverage_map.png"]
"""
import argparse
import os

import geopandas as gpd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAPS = os.path.join(ROOT, "maps")
CRS = "ESRI:102009"   # North America Lambert conformal conic

GREEN, BLUE, GREY, EDGE = "#1BAF7A", "#2A78D6", "#C9C9C9", "#FFFFFF"
DOT_HAVE, DOT_MISSING = "#0B3A66", "#E34948"

# US power markets / balancing areas with power generation by type (EIA-930 daily, plus MISO's own 5-min feed).
# HIFLD control-area NAME (upper case, prefix match) -> (label, class)
BA_COVERAGE = {
    "PJM INTERCONNECTION": ("PJM", "blue"),
    "MIDCONTINENT INDEPENDENT": ("MISO", "blue"),
    "CALIFORNIA INDEPENDENT": ("CAISO", "blue"),
    "ELECTRIC RELIABILITY COUNCIL OF TEXAS": ("ERCOT", "blue"),
    "SOUTHWEST POWER POOL": ("SPP", "blue"),
    "NEW YORK INDEPENDENT": ("NYISO", "blue"),
    "ISO NEW ENGLAND": ("ISO-NE", "blue"),
    "SOUTHERN COMPANY SERVICES": ("Southern Co.", "blue"),
    "TENNESSEE VALLEY AUTHORITY": ("TVA", "blue"),
}
# Canadian provinces / Mexico (Natural Earth names) -> (class, note)
COVERAGE = {
    "Ontario": ("blue", "IESO generation by fuel"),
}
NOTES = {"Mexico": "Mexico\n(electricity demand only)"}
GAS_PRODUCERS = {"United States of America": False, "Canada": False, "Mexico": False}
DOT_AT = {"United States of America": (-107.0, 40.0), "Canada": (-114.0, 54.0), "Mexico": (-102.0, 19.5)}
# LNG export plants with feedgas tracked in lng_feedgas_daily.xlsx (lon, lat)
LNG_PLANTS = {"Sabine Pass": (-93.87, 29.73), "Cameron": (-93.33, 30.03), "Calcasieu Pass": (-93.34, 29.78),
              "Corpus Christi": (-97.20, 27.88), "Freeport": (-95.30, 28.93), "Cove Point": (-76.38, 38.39),
              "Elba Island": (-81.00, 32.08), "Plaquemines": (-89.68, 29.48), "Golden Pass": (-93.92, 29.77)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "PNG Charts", "north_america_coverage_map.png"))
    args = ap.parse_args()

    admin = gpd.read_file(os.path.join(MAPS, "north_america_admin1.geojson"))
    bas = gpd.read_file(os.path.join(MAPS, "us_balancing_authorities.geojson"))
    admin = admin[~admin["name"].isin(["Hawaii"])]
    if "Mexico" not in set(admin["admin"]):   # Natural Earth 1:50m admin-1 has no Mexico: add the country outline
        import warnings
        import pandas as pd
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
        mx = world[world["name"] == "Mexico"][["geometry"]].assign(admin="Mexico", name="Mexico", postal="MX")
        admin = gpd.GeoDataFrame(pd.concat([admin, mx.to_crs(admin.crs)], ignore_index=True), crs=admin.crs)

    def admin_fill(row):
        if row["admin"] == "Canada" and row["name"] in COVERAGE:
            return {"green": GREEN, "blue": BLUE}[COVERAGE[row["name"]][0]]
        return GREY

    admin["fill"] = admin.apply(admin_fill, axis=1)
    admin = admin.to_crs(CRS)

    def ba_class(name):
        n = str(name).upper()
        return next((v for k, v in BA_COVERAGE.items() if n.startswith(k)), None)

    bas["cov"] = bas["NAME"].map(ba_class)
    covered = bas[bas["cov"].notna()].copy()
    covered["label"] = covered["cov"].map(lambda v: v[0])
    covered["fill"] = covered["cov"].map(lambda v: {"green": GREEN, "blue": BLUE}[v[1]])
    us_land = admin[admin["admin"] == "United States of America"].dissolve()
    covered = covered.to_crs(CRS)
    land = us_land.geometry.iloc[0].buffer(0)
    covered["geometry"] = covered.geometry.buffer(0).intersection(land)   # clip to land (buffer(0) fixes invalid rings)
    covered["area"] = covered.area
    covered = covered.sort_values("area", ascending=False)   # big footprints first, small ones on top

    fig, ax = plt.subplots(figsize=(12, 11), dpi=150)
    admin.plot(ax=ax, color=admin["fill"], edgecolor=EDGE, linewidth=0.35)
    covered.plot(ax=ax, color=covered["fill"], edgecolor=EDGE, linewidth=1.0)
    # country outlines a little heavier
    admin.dissolve(by="admin").boundary.plot(ax=ax, color="white", linewidth=1.2)

    def xy(lon, lat):
        p = gpd.GeoSeries(gpd.points_from_xy([lon], [lat]), crs=4326).to_crs(CRS).iloc[0]
        return p.x, p.y

    label_at = {"PJM": (-79.0, 39.8), "MISO": (-91.5, 42.5), "CAISO": (-120.0, 37.0), "ERCOT": (-99.0, 31.0),
                "SPP": (-100.0, 39.5), "NYISO": (-75.3, 42.9), "ISO-NE": (-71.6, 44.2), "Southern Co.": (-85.8, 32.4),
                "TVA": (-86.6, 35.6)}
    for label, (lon, lat) in label_at.items():
        x, y = xy(lon, lat)
        ax.text(x, y, label, fontsize=8.5, color="white", fontweight="bold", ha="center", va="center")
    for name, text, (lon, lat) in [("Ontario", "Ontario\n(IESO)", (-85.5, 50.5)), ("Canada", "Canada", (-106.0, 54.5)),
                                   ("Mexico", NOTES["Mexico"], (-102.0, 22.0)),
                                   ("US", "Rest of USA:\nno regional data", (-113.5, 41.0))]:
        x, y = xy(lon, lat)
        ax.text(x, y, text, fontsize=8, color="white" if name == "Ontario" else "#333333", ha="center", va="center",
                fontweight="bold" if name == "Ontario" else "normal")

    for country, have in GAS_PRODUCERS.items():
        x, y = xy(*DOT_AT[country])
        ax.plot(x, y, "o", markersize=9, color=DOT_HAVE if have else DOT_MISSING, markeredgecolor="white",
                markeredgewidth=1.2, zorder=6)
    for name, (lon, lat) in LNG_PLANTS.items():
        x, y = xy(lon, lat)
        ax.plot(x, y, "^", markersize=7, color="#252525", markeredgecolor="white", markeredgewidth=0.8, zorder=7)

    corners = [xy(lon, lat) for lon, lat in [(-127.0, 49.0), (-125.0, 31.0), (-117.0, 14.0), (-86.0, 13.5),
                                              (-62.0, 45.0), (-62.0, 56.0), (-128.0, 56.0)]]
    xs, ys = [c[0] for c in corners], [c[1] for c in corners]
    ax.set_xlim(min(xs), max(xs))
    ax.set_ylim(min(ys), max(ys))
    ax.set_axis_off()

    handles = [Patch(facecolor=GREEN, label="Good coverage: gas demand by sector,\ngas production/supply and\n"
                                            "power generation by type"),
               Patch(facecolor=BLUE, label="Partial: power generation by type,\nor a gas demand split"),
               Patch(facecolor=GREY, label="No data"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_HAVE, markeredgecolor="white",
                      label="Gas producer: domestic production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_MISSING, markeredgecolor="white",
                      label="Gas producer: no production data"),
               Line2D([], [], marker="^", linestyle="none", markersize=7, color="#252525", markeredgecolor="white",
                      label="LNG export plant: feedgas tracked")]
    ax.legend(handles=handles, loc="lower left", frameon=False, fontsize=9, labelspacing=1.0)
    ax.set_title("North America: Gas and Power Data Coverage (USA by power market)", fontsize=14,
                 fontweight="bold", loc="left")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fig.savefig(args.out, bbox_inches="tight", facecolor="white")
    print(f"Saved {args.out}; power markets coloured: {sorted(covered['label'].unique())}")


if __name__ == "__main__":
    main()
