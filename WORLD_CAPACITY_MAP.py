"""
World map of installed generation CAPACITY coverage, country by country: which countries this repo has MONTHLY
installed capacity for (and how good it is), which only annual, and which none. The companion of WORLD_COVERAGE_MAP.py
(gas and power data) and drawn with its projections, Europe inset and two views.

  dark green  - monthly capacity by fuel, multi-year history from the country's own source
  light green - monthly by fuel but short, stale or with holes (Japan Apr 2024 - Mar 2026; Taiwan from Jan 2025;
                Ecuador monthly only from Apr 2024)
  blue        - current fleet only: the source shows today's capacity, so monthly history builds from the first run
                (Oct 2026); Indonesia = renewables only
  amber       - annual official capacity only (ENTSO-E 1 January figures for Europe, StatCan, COES, CNDC, MIEM, DOE ...)
  beige       - Ember yearly capacity only (long-term pages; Mexico's capacity tab is the Ember fallback)
  grey        - no capacity data in the repo
Red dot: a large power system (roughly 45 GW or more installed) on one of our masters with no monthly capacity.

CAPACITY is maintained by hand from the capacity survey of 12 Oct 2026 (every *_power_capacity.xlsx and the China
NEA, Japan METI, Taiwan E-STAT and Indonesia workbooks opened and their dates checked). Update it when a capacity
pull is added or a short series grows.

Usage: python3 WORLD_CAPACITY_MAP.py [--view pacific|atlantic|both] [--out PATH (one view only)]
"""
import argparse
import importlib.util
import os
import warnings

import geopandas as gpd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

ROOT = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("wcm", os.path.join(ROOT, "WORLD_COVERAGE_MAP.py"))
wcm = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(wcm)   # projections, seam cutting, point helper, Europe box

COLOURS = {"monthly": "#1B8A5A", "monthly_short": "#7FD1A8", "current": "#2A78D6", "annual": "#EDA100",
           "ember": "#F3DDB0"}
GREY, EDGE, DOT_GAP = "#D9D9D9", "#FFFFFF", "#E34948"

ENTSOE_ANNUAL = ["Italy", "Netherlands", "Austria", "Slovakia", "Romania", "Bulgaria", "Greece", "Croatia", "Slovenia",
                 "Sweden", "Finland", "Estonia", "Latvia", "Lithuania", "Serbia", "Bosnia and Herz.", "Montenegro",
                 "North Macedonia", "Albania", "Kosovo"]
EMBER_ONLY = ["Venezuela", "Guyana", "Suriname", "Guatemala", "El Salvador", "Honduras", "Nicaragua", "Costa Rica",
              "Panama", "Belize", "Trinidad and Tobago", "Puerto Rico", "Jamaica", "Dominican Rep.", "Mexico",
              "Pakistan", "Thailand", "Vietnam", "Nepal", "Myanmar", "Laos", "Brunei", "Timor-Leste", "Bhutan",
              "Singapore", "Malta", "Cyprus", "Israel", "Egypt", "Lebanon",
              # Africa master: Ember yearly (South Africa: Eskom nuclear and pumped storage only)
              "Nigeria", "South Africa", "Cameroon", "Côte d'Ivoire", "Benin", "Togo", "Eq. Guinea", "Mauritania",
              "Mali", "Burkina Faso", "Guinea", "Sierra Leone", "Liberia", "Gambia", "Guinea-Bissau", "Senegal",
              "Niger"]

# name (Natural Earth): (class, note). Europe, the Americas, East Med and Africa are filled to a gap-free monthly series
# (capacity_fill.py; *_capacity_monthly_complete.xlsx) - the class is the BASIS behind it. Asia is not filled yet.
CAPACITY = {
    "United States of America": ("monthly", "EIA-860M, Jan 2021 -, by fuel and ISO region"),
    "India": ("monthly", "CEA/NPP + MNRE, Jan 2019 -"),
    "Argentina": ("monthly", "CAMMESA"), "Brazil": ("monthly", "ANEEL SIGA + MMGD"), "Chile": ("monthly", "CNE"),
    "Colombia": ("monthly", "XM"),
    "Australia": ("monthly", "AEMO registration list, Jan 2021 -"),
    "South Korea": ("monthly", "KPX EPSIS, Jan 2018 -"),
    "China": ("monthly", "NEA, Dec 2020 - (January never published; thermal not split)"),
    "Germany": ("monthly", "MaStR register rebuild, Jan 2015 -"), "France": ("monthly", "ODRE register rebuild"),
    "Spain": ("monthly", "REE REData"), "Hungary": ("monthly", "MAVIR"), "Czechia": ("monthly", "ERU"),
    "Switzerland": ("monthly", "BFE/Pronovo register rebuild"), "Turkey": ("monthly", "EPDK sector reports, Feb 2017 -"),
    "Japan": ("monthly_short", "METI table 1-1, Apr 2024 - Mar 2026 only"),
    "Taiwan": ("monthly_short", "E-STAT 3-03, monthly from Jan 2025"),
    "Ecuador": ("monthly_short", "ARCONEL, monthly from Apr 2024"),
    "United Kingdom": ("monthly_short", "solar/wind/storage monthly; thermal DUKES annual"),
    "Ireland": ("monthly_short", "EirGrid wind + register rebuild; thermal annual"),
    "Belgium": ("monthly_short", "Elia wind and solar monthly; rest ENTSO-E annual"),
    "Denmark": ("monthly_short", "Energinet wind and solar monthly; thermal ENTSO-E annual"),
    "Norway": ("monthly_short", "NVE hydro and wind register rebuild"),
    "Portugal": ("monthly_short", "DGEG renewables monthly; thermal annual"),
    "Bangladesh": ("current", "BPDB current fleet, history from Oct 2026"),
    "Malaysia": ("current", "GSO current fleet, Peninsular only"),
    "Sri Lanka": ("current", "PUCSL current fleet"),
    "Indonesia": ("current", "ESDM EBTKE, renewables only, one snapshot"),
    "Poland": ("annual", "ENTSO-E annual; URE wind quarterly"),
    "Canada": ("annual", "StatCan year-end, 2019 - 2024"),
    "Peru": ("annual", "COES annual statistics"), "Bolivia": ("annual", "CNDC"), "Uruguay": ("annual", "MIEM/DNE"),
    "Paraguay": ("annual", "Itaipu, EBY, Ember for ANDE plants"),
    "Philippines": ("annual", "DOE Power Statistics, 2003 - 2025"), "Cambodia": ("annual", "EAC annual reports"),
    "New Zealand": ("annual", "MBIE plant-type table"),
    "Luxembourg": ("annual", "ILR annual"), "Jordan": ("annual", "NEPCO annual reports"),
    "Ghana": ("annual", "Energy Commission"), "Morocco": ("annual", "ONEE via HCP yearbooks"),
}
CAPACITY.update({c: ("annual", "ENTSO-E installed capacity, 1 January") for c in ENTSOE_ANNUAL})
CAPACITY.update({c: ("ember", "Ember yearly") for c in EMBER_ONLY if c not in CAPACITY})

BIG_GAPS = ["Canada", "Mexico", "Italy", "Netherlands", "Poland", "Egypt", "South Africa", "Pakistan", "Vietnam",
            "Thailand", "Indonesia"]
SMALL = {"Singapore": (103.82, 1.35), "Trinidad and Tobago": (-61.3, 10.45), "Malta": (14.4, 35.9),
         "Luxembourg": (6.13, 49.8)}
DOT_AT = dict(wcm.DOT_AT)
DOT_AT.update({"Mexico": (-102.0, 23.5), "South Africa": (24.0, -30.0), "France": (2.5, 46.8),
               "Netherlands": (5.6, 52.3), "Indonesia": (114.0, -1.5)})


def load_world():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    world = world[world["name"] != "Antarctica"].copy()
    world.loc[world["name"] == "N. Cyprus", "name"] = "Cyprus"
    world.loc[world["name"] == "Somaliland", "name"] = "Somalia"
    world["geometry"] = world.geometry.map(wcm.cut_at_seam)
    world["fill"] = world["name"].map(lambda n: COLOURS[CAPACITY[n][0]] if n in CAPACITY else GREY)
    return world


def draw(ax, world, small_dots=9):
    world.plot(ax=ax, color=world["fill"], edgecolor=EDGE, linewidth=0.4)
    for name, (lon, lat) in SMALL.items():
        p = wcm.pt(lon, lat)
        ax.plot(p.x, p.y, "o", markersize=small_dots * 0.6, color=COLOURS[CAPACITY[name][0]],
                markeredgecolor="#333333", markeredgewidth=0.6, zorder=6)
    geo = world.set_index("name").to_crs(4326).geometry
    for name in BIG_GAPS:
        if name in DOT_AT:
            lon, lat = DOT_AT[name]
        else:
            p = geo[name].representative_point()
            lon, lat = p.x, p.y
        q = wcm.pt(lon, lat)
        ax.plot(q.x, q.y, "o", markersize=small_dots * 0.62, color=DOT_GAP, markeredgecolor="white",
                markeredgewidth=0.8, zorder=7)


def render(out):
    world = load_world().to_crs(wcm.CRS)
    fig = plt.figure(figsize=(18, 10.5), dpi=150)
    ax = fig.add_axes([0.0, 0.08, 1.0, 0.86])
    ax.set_anchor("N")
    draw(ax, world)
    ax.set_axis_off()
    ax.set_title("World: Installed Generation Capacity Coverage (monthly vs annual)", fontsize=18,
                 fontweight="bold", loc="left", x=0.01)

    ins = fig.add_axes(wcm.INSET_AT)
    draw(ins, world, small_dots=11 * wcm.INSET_AT[2] / 0.25)
    lo, la, hi_lo, hi_la = wcm.EUROPE_BOX
    a, b = wcm.pt(lo, la), wcm.pt(hi_lo, hi_la)
    c, d = wcm.pt(lo, hi_la), wcm.pt(hi_lo, la)
    ins.set_xlim(min(a.x, c.x), max(b.x, d.x)); ins.set_ylim(a.y, b.y)
    ins.set_xticks([]); ins.set_yticks([])
    for s in ins.spines.values():
        s.set_edgecolor("#9A9A9A"); s.set_linewidth(0.8)
    ins.set_title("Europe", fontsize=10, loc="left", color="#333333")

    n = {k: sum(1 for v in CAPACITY.values() if v[0] == k) for k in COLOURS}
    handles = [Patch(facecolor=COLOURS["monthly"], label=f"Monthly national source, all main fuels ({n['monthly']})"),
               Patch(facecolor=COLOURS["monthly_short"], label=f"Monthly for some fuels, or short / with holes "
                                                                f"({n['monthly_short']})"),
               Patch(facecolor=COLOURS["current"], label=f"Current fleet only, history building from Oct 2026 "
                                                         f"({n['current']})"),
               Patch(facecolor=COLOURS["annual"], label=f"Annual official (ENTSO-E Europe, StatCan, ...) ({n['annual']})"),
               Patch(facecolor=COLOURS["ember"], label=f"Ember yearly fallback only ({n['ember']})"),
               Patch(facecolor=GREY, label="No capacity data"),
               Line2D([], [], marker="o", linestyle="none", markersize=8, color=DOT_GAP, markeredgecolor="white",
                      label="Large system with no monthly national capacity source")]
    fig.legend(handles=handles, loc="lower center", ncol=3, frameon=False, fontsize=10, bbox_to_anchor=(0.5, 0.0),
               columnspacing=2.5, labelspacing=0.9)
    fig.text(0.995, 0.005, "Singapore, Trinidad & Tobago, Malta, Luxembourg shown as markers. Survey of the repo's "
             "capacity workbooks, 12 Oct 2026. Europe, the Americas, East Med and Africa are filled to gap-free monthly series "
             "(colour = basis behind the fill); Asia and Oceania are not filled yet.",
             ha="right", fontsize=8, color="#6B6B6B")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    fig.savefig(out, facecolor="white", bbox_inches="tight", pad_inches=0.15)
    print(f"Saved {out}: {n}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--view", choices=["pacific", "atlantic", "both"], default="both")
    ap.add_argument("--out", default=None, help="output path (one view only); default world_capacity_map[_atlantic].png")
    args = ap.parse_args()
    base = os.path.join(ROOT, "output", "Data and Chart Outputs", "Master Outputs", "world_capacity_map")
    for view in (["pacific", "atlantic"] if args.view == "both" else [args.view]):
        wcm.set_view(view)
        render(args.out if args.out and args.view != "both" else base + ("" if view == "pacific" else "_atlantic") + ".png")


if __name__ == "__main__":
    main()
