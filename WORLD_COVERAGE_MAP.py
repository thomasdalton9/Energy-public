"""
World map coloured by how much gas and power data this repo pulls, country by country (the regional maps in
south_america/COVERAGE_MAP.py, asia/SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP.py and americas/NORTH_AMERICA_COVERAGE_MAP.py
show the detail):

  green  - good: gas demand and supply/production plus power generation by type, from raw sources (grid operators,
           gas TSOs, ministries, statistics offices). A country with no gas market needs only the power side.
  blue   - partial: raw power generation by type, or raw gas data, but not both
  amber  - Ember only: power generation by type compiled by Ember (monthly or yearly fallback)
  grey   - annual history only (Ember yearly, Energy Institute, World Bank / IMF: the long-term pages)
Dots: dark blue = gas producer whose production we pull, red = significant gas producer without production data,
light blue = hydro reservoir / dam-level data.

COVERAGE is maintained by hand: update it when a pull is added (South & Central America is read from
south_america/COVERAGE_MAP.py so the two stay in step). Country shapes: Natural Earth 1:110m (bundled with
geopandas 0.14), Equal Earth projection centred on the Pacific (150°E), so the Atlantic is the map's
edge; countries crossing 30°W (Greenland) are cut there.

Usage: python3 WORLD_COVERAGE_MAP.py [--out "output/PNG Charts/world_coverage_map.png"]
"""
import argparse
import importlib.util
import os
import warnings

import geopandas as gpd
import matplotlib
from shapely.geometry import box

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = os.path.dirname(os.path.abspath(__file__))
CRS = "+proj=robin +lon_0=150 +datum=WGS84 +units=m +no_defs"   # Robinson, Pacific-centred
SEAM = -30.0   # the map's edge (150°E - 180°): shapes crossing it are cut there so they do not smear across the map
GREEN, BLUE, AMBER, GREY, EDGE = "#1BAF7A", "#2A78D6", "#EDA100", "#D9D9D9", "#FFFFFF"
FILL = {"green": GREEN, "blue": BLUE, "amber": AMBER}
DOT_HAVE, DOT_MISSING, DOT_HYDRO = "#0B3A66", "#E34948", "#8FD3FF"

# South & Central America and the Caribbean: the regional map's own classification
_spec = importlib.util.spec_from_file_location("sa_cov", os.path.join(ROOT, "south_america", "COVERAGE_MAP.py"))
sa = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sa)

COVERAGE = dict(sa.COVERAGE)
COVERAGE.update({
    # North America
    "United States of America": ("green", "Gas demand by sector, supply, storage (EIA); power by type per balancing "
                                          "authority (EIA-930); LNG feedgas"),
    "Canada": ("green", "Gas supply and disposition (StatCan); power by type (StatCan, IESO)"),
    "Mexico": ("blue", "Gas imports from the US, pipeline capacity; CENACE demand; power by type Ember only"),
    # South & Southeast Asia
    "India": ("green", "Gas (PPAC); power CEA daily + NITI Aayog ICED; reservoirs; IEX prices"),
    "Bangladesh": ("green", "Gas production and distribution (Petrobangla daily); power PGCB"),
    "Thailand": ("green", "Gas supply and demand by sector (EPPO); power EPPO/EGAT; reservoirs RID"),
    "Pakistan": ("blue", "Power CPPA-G/NEPRA (+ K-Electric where filed); gas production by province (PBS, to Mar-24), "
                         "LNG import value; reservoirs IRSA"),
    "Vietnam": ("blue", "Power EVN/NSMO daily; gas production (NSO); no demand split"),
    "Indonesia": ("blue", "Gas production by contractor + use by sector (Ditjen Migas); power Ember yearly"),
    "Philippines": ("blue", "Power IEMOP plant dispatch by fuel; prices; dam levels; no gas data"),
    "Malaysia": ("blue", "Power GSO (Peninsular); SMP prices; no gas data"),
    "Sri Lanka": ("blue", "Power PUCSL; reservoirs (no gas market)"),
    "Bhutan": ("blue", "Power BPSO (no gas market)"),
    "Nepal": ("blue", "Power NEA LDC (no gas market)"),
    "Cambodia": ("blue", "Power EAC annual reports (no gas market)"),
    "Myanmar": ("amber", "Power Ember yearly"),
    "Laos": ("amber", "Power Ember yearly"),
    "Brunei": ("amber", "Power Ember yearly"),
    "Timor-Leste": ("amber", "Power Ember yearly"),
    # Australia & New Zealand
    "Australia": ("green", "Gas demand by sector, production, storage, LNG (AEMO, WA GBB); power NEM + WEM; hydro "
                           "storage (Tasmania)"),
    "New Zealand": ("green", "Gas (MBIE); power by fuel (EMI)"),
    # Europe: ENTSO-E power by type + ENTSOG / national TSO gas balances (consumption, production, imports)
    "United Kingdom": ("green", "Gas NTS demand by sector + production (National Gas, NSTA); power Elexon/NESO"),
    "Ireland": ("green", "Gas GNI; power EirGrid / ENTSO-E SEM"),
    "Norway": ("green", "Gas flows by destination (Gassco); power ENTSO-E"),
    "Switzerland": ("blue", "Power ENTSO-E / Swissgrid"),
    "Turkey": ("blue", "Power EPIAS"),
    "Cyprus": ("blue", "Power TSOC (no gas market)"),
    "Serbia": ("blue", "Power ENTSO-E"),
    "Bosnia and Herz.": ("blue", "Power ENTSO-E"),
    "Montenegro": ("blue", "Power ENTSO-E"),
    "North Macedonia": ("blue", "Power ENTSO-E"),
    "Albania": ("blue", "Power ENTSO-E (from May-26)"),
    "Kosovo": ("blue", "Power ENTSO-E"),
    "Luxembourg": ("blue", "Power inside the DE-LU zone; gas via ENTSOG"),
    # Africa / East Asia
    "South Africa": ("blue", "Power by type (Eskom daily)"),
    "China": ("blue", "Power by type and gas production (NBS monthly); no demand split"),
})
for c in ["Austria", "Belgium", "Bulgaria", "Croatia", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany",
          "Greece", "Hungary", "Italy", "Latvia", "Lithuania", "Netherlands", "Poland", "Portugal", "Romania",
          "Slovakia", "Slovenia", "Spain", "Sweden"]:
    COVERAGE[c] = ("green", "Power ENTSO-E; gas balance ENTSOG / national TSO, storage AGSI+, LNG ALSI")

GAS_PRODUCERS = dict(sa.GAS_PRODUCERS)
GAS_PRODUCERS.update({
    "United States of America": True, "Canada": True, "Mexico": False, "India": True, "Bangladesh": True,
    "Thailand": True, "Pakistan": True, "Vietnam": True, "Indonesia": True, "Malaysia": False, "Myanmar": False,
    "Brunei": False, "Philippines": False, "Australia": True, "New Zealand": True, "China": True,
    "United Kingdom": True, "Norway": True, "Netherlands": True, "Germany": True, "Romania": True, "Italy": True,
    "Poland": True, "Denmark": True, "Hungary": True, "Croatia": True, "Ireland": True, "Ukraine": False,
    "Russia": False, "Qatar": False, "Iran": False, "Saudi Arabia": False, "United Arab Emirates": False,
    "Algeria": False, "Egypt": False, "Nigeria": False, "Turkmenistan": False, "Uzbekistan": False,
    "Kazakhstan": False, "Oman": False, "Azerbaijan": False, "Iraq": False, "Kuwait": False, "Libya": False,
    "Israel": False, "Mozambique": False, "Japan": None, "Angola": False,
})
GAS_PRODUCERS = {k: v for k, v in GAS_PRODUCERS.items() if v is not None}
HYDRO = set(sa.HYDRO) | {"India", "Thailand", "Philippines", "Pakistan", "Sri Lanka", "Australia"}
# dot positions (lon, lat) where the representative point is awkward
DOT_AT = {"Chile": (-71.0, -36.0), "Norway": (9.0, 61.5), "Croatia": (16.0, 45.3), "Denmark": (9.3, 56.0),
          "United Kingdom": (-1.5, 53.0), "Indonesia": (114.0, -1.5), "Malaysia": (102.0, 4.0),
          "Russia": (60.0, 61.0), "Canada": (-112.0, 55.0), "United States of America": (-100.0, 39.0),
          "Philippines": (122.5, 12.5), "Vietnam": (106.5, 11.5), "Ireland": (-8.0, 53.2)}
HYDRO_OFFSET = (3.5, 0.0)   # hydro dot sits to the right of the gas dot (degrees)
SMALL = {"Singapore": (103.82, 1.35, "green"), "Trinidad and Tobago": (-61.3, 10.45, "green"),
         "Malta": (14.4, 35.9, None)}
EUROPE_BOX = (-12.0, 34.0, 33.0, 71.5)   # lon_min, lat_min, lon_max, lat_max for the inset
INSET_AT = [0.585, 0.12, 0.20, 0.30]   # figure fraction: left, bottom, width, height


def cut_at_seam(g, eps=1e-6):
    """Split a shape that crosses the map's edge (SEAM) into its two sides."""
    if not g.intersects(box(SEAM - eps, -90, SEAM + eps, 90)):
        return g
    return g.intersection(box(-180, -90, SEAM - eps, 90)).union(g.intersection(box(SEAM + eps, -90, 180, 90)))


def load_world():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    world = world[world["name"] != "Antarctica"].copy()
    world.loc[world["name"] == "N. Cyprus", "name"] = "Cyprus"
    world.loc[world["name"] == "Somaliland", "name"] = "Somalia"
    world["geometry"] = world.geometry.map(cut_at_seam)
    world["fill"] = world["name"].map(lambda n: FILL[COVERAGE[n][0]] if n in COVERAGE else GREY)
    return world


def pt(lon, lat):
    return gpd.GeoSeries(gpd.points_from_xy([lon], [lat]), crs=4326).to_crs(CRS).iloc[0]


def draw(ax, world, dots=True, small_dots=9, extent=None):
    world.plot(ax=ax, color=world["fill"], edgecolor=EDGE, linewidth=0.4)
    for name, (lon, lat, cls) in SMALL.items():
        if cls:
            p = pt(lon, lat)
            ax.plot(p.x, p.y, "o", markersize=small_dots * 0.6, color=FILL[cls], markeredgecolor="#333333",
                    markeredgewidth=0.6, zorder=6)
    if not dots:
        return
    geo = world.set_index("name").to_crs(4326).geometry
    for name, have in GAS_PRODUCERS.items():
        if name in DOT_AT:
            lon, lat = DOT_AT[name]
        elif name in geo.index:
            p = geo[name].representative_point()
            lon, lat = p.x, p.y
        else:
            continue
        q = pt(lon, lat)
        ax.plot(q.x, q.y, "o", markersize=small_dots * 0.62, color=DOT_HAVE if have else DOT_MISSING,
                markeredgecolor="white", markeredgewidth=0.8, zorder=7)
    for name in HYDRO:
        if name in DOT_AT:
            lon, lat = DOT_AT[name]
        elif name in geo.index:
            p = geo[name].representative_point()
            lon, lat = p.x, p.y
        else:
            continue
        if name in GAS_PRODUCERS:
            lon += HYDRO_OFFSET[0]
        q = pt(lon, lat)
        ax.plot(q.x, q.y, "o", markersize=small_dots * 0.62, color=DOT_HYDRO, markeredgecolor=DOT_HAVE,
                markeredgewidth=0.7, zorder=7)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "PNG Charts", "world_coverage_map.png"))
    args = ap.parse_args()
    world = load_world().to_crs(CRS)

    fig = plt.figure(figsize=(18, 10.5), dpi=150)
    ax = fig.add_axes([0.0, 0.08, 1.0, 0.86])
    ax.set_anchor("N")
    draw(ax, world)
    ax.set_axis_off()
    ax.set_title("World: Gas and Power Data Coverage", fontsize=18, fontweight="bold", loc="left", x=0.01)

    # Europe inset (bottom, over the South Pacific between New Zealand and South America)
    ins = fig.add_axes(INSET_AT)
    draw(ins, world, small_dots=11)
    lo, la, hi_lo, hi_la = EUROPE_BOX
    a, b = pt(lo, la), pt(hi_lo, hi_la)
    c, d = pt(lo, hi_la), pt(hi_lo, la)
    ins.set_xlim(min(a.x, c.x), max(b.x, d.x)); ins.set_ylim(a.y, b.y)
    ins.set_xticks([]); ins.set_yticks([])
    for s in ins.spines.values():
        s.set_edgecolor("#9A9A9A"); s.set_linewidth(0.8)
    ins.set_title("Europe", fontsize=10, loc="left", color="#333333")

    counts = {k: sum(1 for v in COVERAGE.values() if v[0] == k) for k in ("green", "blue", "amber")}
    handles = [Patch(facecolor=GREEN, label=f"Good: gas demand + supply and power generation by type, raw sources "
                                            f"({counts['green']} countries)"),
               Patch(facecolor=BLUE, label=f"Partial: raw power by type, or raw gas data, not both "
                                           f"({counts['blue']})"),
               Patch(facecolor=AMBER, label=f"Ember only: compiled power by type ({counts['amber']})"),
               Patch(facecolor=GREY, label="Annual history only (Ember yearly, Energy Institute, World Bank / IMF)"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_HAVE, markeredgecolor="white",
                      label="Gas producer: production data pulled"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_MISSING, markeredgecolor="white",
                      label="Gas producer: no production data"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_HYDRO, markeredgecolor=DOT_HAVE,
                      label="Hydro reservoir / dam-level data")]
    fig.legend(handles=handles, loc="lower center", ncol=3, frameon=False, fontsize=10, bbox_to_anchor=(0.5, 0.0),
               columnspacing=2.5, labelspacing=0.9)
    fig.text(0.995, 0.005, "Singapore, Trinidad & Tobago shown as markers (below map resolution). Regional detail: "
             "the South & Central America, North America and South & Southeast Asia coverage maps.", ha="right",
             fontsize=8, color="#6B6B6B")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fig.savefig(args.out, facecolor="white", bbox_inches="tight", pad_inches=0.15)
    print(f"Saved {args.out}: {counts}")


if __name__ == "__main__":
    main()
