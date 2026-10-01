"""
Register of large industrial natural-gas users in South America, Central America and the Caribbean:
fertiliser (ammonia/urea), methanol, steel DRI, alumina, cement (where gas-fired), glass, ceramics,
petrochemicals and LNG liquefaction. A reference list, not a time series.

The data live in a hand-curated CSV, south_america/industrial_gas_users.csv (one row per plant, every
capacity with its source URL). Edit the CSV, then rebuild:

    python3 south_america/INDUSTRIAL_GAS_USERS.py [--png-dir "output/PNG Charts"] [--no-png]
    python3 add_charts.py "output/Data and Chart Outputs/latin_america_industrial_gas_users.xlsx"

Writes output/Data and Chart Outputs/latin_america_industrial_gas_users.xlsx:
  Units               sources, units, completeness by sector, as-of date
  Plants              one row per plant: nameplate capacity (as published and in Mt/yr of product; ceramics in
                      million m2/yr), status, location, reported gas use where a source states it, source URLs
  By country & sector plant counts by status and operating capacity
then add_charts.py adds the native chart (operating plants by country, stacked by sector).
With --png-dir it also renders PNGs: the same bar chart and a map of the plants sized by capacity
(geopandas 0.14 naturalearth_lowres, as south_america/COVERAGE_MAP.py).

No gas use is estimated. 'reported_gas_mcm_d' is filled only where a company or official source states the
plant's gas use or contract volume.
"""
import argparse
import datetime as dt
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

CSV = os.path.join(ROOT, "south_america", "industrial_gas_users.csv")
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "latin_america_industrial_gas_users.xlsx")
STEM = "latin_america_industrial_gas_users"

SECTOR_ORDER = ["LNG liquefaction", "Ammonia/urea", "Methanol", "Steel DRI", "Alumina", "Cement", "Glass",
                "Ceramics", "Petrochemical", "Other"]
COMPLETENESS = {
    "Ammonia/urea": "Near complete: every Trinidad plant (MEEI list), Venezuela's three Pequiven complexes, Argentina "
                    "(Profertil), Bolivia (Bulo Bulo), Brazil (Petrobras FAFEN-BA/SE, ANSA, UFN-III, Yara Cubatao). "
                    "Small or idle units (e.g. Colombia's Ferticol) are not listed.",
    "Methanol": "Complete for the region's methanol plants (Trinidad x8, Venezuela x2, Chile, Argentina).",
    "LNG liquefaction": "Complete for export LNG (Atlantic LNG, Peru LNG, Argentina's Hilli FLNG under construction) plus "
                        "Ecuador's small Bajo Alto plant; Bolivia's small-scale Rio Grande plant is not listed.",
    "Steel DRI": "Complete for gas-based DRI/HBI (Trinidad, Venezuela, Argentina). Brazil, Peru and Colombia steel is "
                 "blast-furnace, coal-based DRI or scrap-EAF, so not listed.",
    "Alumina": "Jamalco (LNG-fired) and Brazil's Alunorte (fuel-oil to gas switch) are included; Venezuela's Bauxilum "
               "is listed although its fuel is unconfirmed. Alumar, CBA, Alpart and Windalco are not listed "
               "(gas use unconfirmed).",
    "Cement": "Partial: gas firing is documented for Trinidad (TCL) and Bolivia; Venezuela and Argentina plants are "
              "listed but their kiln fuel mix is unconfirmed. Brazil, Colombia, Chile, Peru, Ecuador, Central America "
              "and the Dominican Republic burn mainly coal/petcoke: not listed.",
    "Glass": "Partial: Brazil's four float-glass producers only. Container glass (O-I, Verallia...) and float glass "
             "elsewhere (Colombia, Argentina) are not listed.",
    "Ceramics": "Partial: Brazil's four largest tile groups (2024 production, ANFACER / Ceramic World Review), "
                "Colombia (Ceramica Italia, Corona Sopo new plant), Ecuador (Graiman, Italpisos). Brazil alone "
                "has ~60 companies / 71 plants (ANFACER); Peru, Argentina and Mexico-owned plants are not listed.",
    "Petrochemical": "Ethane-fed crackers only (Dow Bahia Blanca, Braskem Duque de Caxias); their feed is ethane "
                     "(an NGL) separated from natural gas.",
}
SOURCES = [
    "USGS Minerals Yearbook, country chapters, Table 2 'Structure of the mineral industry' "
    "(https://www.usgs.gov/centers/national-minerals-information-center): Argentina, Bolivia, Brazil, Chile 2019; "
    "Islands of the Caribbean 2019; Venezuela 2022.",
    "Trinidad & Tobago Ministry of Energy and Energy Industries (MEEI): ammonia, methanol and urea plant tables "
    "(https://www.energy.gov.tt/our-business/lng-petrochemicals/petrochemicals/).",
    "Global Energy Monitor, Global Steel Plant Tracker wiki pages (gem.wiki, CC BY 4.0): plant coordinates and "
    "DRI capacities.",
    "Company releases and filings (Methanex, Nutrien, Nucor, Petrobras, Yara, Century Aluminum, Atlantic LNG), "
    "ANFACER / Ceramic World Review (ceramics), trade and national press: one or more URLs on each row of 'Plants'.",
]

COLS = ["country", "plant", "operator", "sector", "products", "capacity", "capacity_unit", "capacity_check",
        "capacity_mt_yr", "ceramics_mm2_yr", "status", "status_as_of", "counted_as", "city", "region", "lat", "lon",
        "coords", "reported_gas_mcm_d", "reported_gas_source", "notes", "sources"]


def load(path=CSV):
    d = pd.read_csv(path, dtype=str, keep_default_na=False)
    for c in ("capacity", "capacity_mt_yr", "ceramics_mm2_yr", "lat", "lon", "reported_gas_mcm_d"):
        d[c] = pd.to_numeric(d[c].str.replace(",", ""), errors="coerce")
    return d


def counted_as(status):
    s = status.lower()
    if s.startswith("operating"):
        return "Operating"
    if "construction" in s:
        return "Under construction"
    if s.startswith(("idle", "mothball")):
        return "Idle"
    if s.startswith("closed"):
        return "Closed"
    return "Status unverified"


def build_plants(d):
    p = d.assign(counted_as=d["status"].map(counted_as))
    p["sector"] = pd.Categorical(p["sector"], SECTOR_ORDER, ordered=True)
    p = p.sort_values(["country", "sector", "plant"]).reset_index(drop=True)
    p["sector"] = p["sector"].astype(str)
    p.index = pd.RangeIndex(1, len(p) + 1, name="id")
    return p[COLS]


def pivot(p):
    op = p[p["counted_as"] == "Operating"]
    keys = ["country", "sector"]
    out = pd.DataFrame({
        "plants": p.groupby(keys).size(),
        "operating": op.groupby(keys).size(),
        "idle_or_closed": p[p["counted_as"].isin(["Idle", "Closed"])].groupby(keys).size(),
        "under_construction": p[p["counted_as"] == "Under construction"].groupby(keys).size(),
        "status_unverified": p[p["counted_as"] == "Status unverified"].groupby(keys).size(),
        "operating_capacity_mt_yr": op.groupby(keys)["capacity_mt_yr"].sum(min_count=1),
        "operating_ceramics_mm2_yr": op.groupby(keys)["ceramics_mm2_yr"].sum(min_count=1),
    }).reset_index()
    tot = out.groupby("country").sum(numeric_only=True, min_count=1).reset_index()
    tot["sector"] = "All sectors"
    out = pd.concat([out, tot], ignore_index=True)
    out["_o"] = out["sector"].map({s: i for i, s in enumerate(SECTOR_ORDER + ["All sectors"])})
    out = out.sort_values(["country", "_o"]).drop(columns="_o")
    grand = out[out["sector"] == "All sectors"].sum(numeric_only=True, min_count=1)
    grand["country"], grand["sector"] = "ALL", "All sectors"
    out = pd.concat([out, grand.to_frame().T], ignore_index=True)
    counts = ["plants", "operating", "idle_or_closed", "under_construction", "status_unverified"]
    for c in counts:
        out[c] = pd.to_numeric(out[c]).fillna(0).astype(int)
    for c in ("operating_capacity_mt_yr", "operating_ceramics_mm2_yr"):
        out[c] = pd.to_numeric(out[c], errors="coerce").round(3)
    return out.set_index(["country", "sector"])


def notes(p):
    asof = dt.date.today().isoformat()
    op = p[p["counted_as"] == "Operating"]
    top = op.sort_values("capacity_mt_yr", ascending=False).head(10)
    rep = p[p["reported_gas_mcm_d"].notna()]
    lines = ["Latin America & Caribbean industrial gas users - plant register", "",
             "AS-OF", f"Register compiled {asof}; each row's status carries its own as-of date (status_as_of).",
             "A reference list of large gas-using industrial plants, not a time series. Mexico is out of scope.",
             "Edit south_america/industrial_gas_users.csv and rerun south_america/INDUSTRIAL_GAS_USERS.py "
             "(then add_charts.py) to rebuild.", "",
             "UNITS",
             "capacity / capacity_unit: nameplate as published by the source (kt/yr, t/d, Mt/yr, million m2).",
             "capacity_check: whether an automated fetch found the figure on a cited page (Oct-2026); 'UNVERIFIED' "
             "rows rest on figures whose pages block automated access - treat with care.",
             "capacity_mt_yr: the same capacity in million tonnes per year of the product named in capacity_unit "
             "(t/d x 350 days); LNG in Mt/yr (MTPA). Products differ by sector (ammonia, urea, methanol, DRI, "
             "LNG, alumina, cement, glass, ethylene), so add capacities up within a sector only.",
             "ceramics_mm2_yr: ceramic tile capacity or production in million m2 per year (not converted to tonnes).",
             "reported_gas_mcm_d: gas use in million m3/day ONLY where a company or official source states it "
             "(source in reported_gas_source); no gas use is estimated.",
             "counted_as: Operating / Idle / Closed / Under construction / Status unverified, from 'status'.",
             "lat / lon: approximate (see 'coords': GEM wiki plant coordinates where available, otherwise town or "
             "industrial-estate level).", "",
             "BY COUNTRY & SECTOR",
             "Plant counts by status; operating capacity summed in Mt/yr of product, ceramics separately in million m2/yr.",
             "", "LARGEST OPERATING PLANTS BY CAPACITY (Mt/yr of product)"]
    for _, r in top.iterrows():
        lines.append(f"  {r['capacity_mt_yr']:.2f} - {r['plant']} ({r['country']}, {r['sector']}, {r['capacity_unit']})")
    lines += ["", "REPORTED GAS USE"]
    for _, r in rep.iterrows():
        lines.append(f"  {r['reported_gas_mcm_d']:.2f} mcm/d - {r['plant']} ({r['country']}): {r['reported_gas_source']}")
    lines += ["", "COMPLETENESS"]
    for s in SECTOR_ORDER:
        if s in COMPLETENESS:
            lines.append(f"  {s}: {COMPLETENESS[s]}")
    lines += ["Countries with no plant listed: Uruguay, Paraguay, Central America (Guatemala-Panama, Belize), "
              "Dominican Republic, Puerto Rico - gas there goes to power; cement kilns burn coal/petcoke.", "",
              "SOURCES"] + SOURCES + ["Every row of 'Plants' carries its own source URL(s); statuses marked "
                                      "'unverified' were not confirmed by a dated source."]
    return lines


def count_table(p):
    """Operating plants per country (largest first) x sector; Petrochemical folds into 'Other' (8 palette slots)."""
    op = p[p["counted_as"] == "Operating"].copy()
    op["sector"] = op["sector"].where(op["sector"] != "Petrochemical", "Other")
    t = op.pivot_table(index="country", columns="sector", values="plant", aggfunc="count").fillna(0).astype(int)
    t = t[[s for s in SECTOR_ORDER if s in t.columns]]
    t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    t.index = [c.replace("Trinidad and Tobago", "Trinidad & Tobago") for c in t.index]
    return t


def render_pngs(p, png_dir):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    from xlsx_charts import OTHER_GREY, PALETTE
    os.makedirs(png_dir, exist_ok=True)
    colours = {s: "#" + c for s, c in zip(SECTOR_ORDER[:8], PALETTE)}
    colours["Other"] = "#" + OTHER_GREY
    src = ("Source: plant register compiled from USGS Minerals Yearbook, Trinidad MEEI, Global Energy Monitor (gem.wiki) "
           "and company reports; see the workbook's Units tab.")

    # 1) operating plants by country, stacked by sector (same table as the native chart)
    t = count_table(p)
    fig, ax = plt.subplots(figsize=(11, 6), dpi=150)
    bottom = pd.Series(0.0, index=t.index)
    for s in t.columns:
        label = "Other (petrochemical)" if s == "Other" else s
        ax.bar(t.index, t[s], bottom=bottom, color=colours[s], label=label, width=0.65, edgecolor="white", linewidth=1)
        bottom += t[s]
    for x, v in zip(t.index, bottom):
        ax.text(x, v + 0.2, f"{int(v)}", ha="center", va="bottom", fontsize=8, color="#333333")
    ax.set_ylim(0, bottom.max() * 1.1)
    ax.set_ylabel("number of operating plants")
    ax.set_title("Large industrial gas users: operating plants by country and sector", loc="left", fontsize=12,
                 fontweight="bold")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.grid(axis="y", color="#E5E5E5", linewidth=0.6)
    ax.set_axisbelow(True)
    plt.setp(ax.get_xticklabels(), fontsize=9)
    ax.legend(frameon=False, fontsize=8, ncol=3, loc="upper right")
    fig.text(0.01, 0.005, src, fontsize=7, color="#666666")
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    bar_png = os.path.join(png_dir, f"{STEM}__Chart.png")
    fig.savefig(bar_png, facecolor="white")
    plt.close(fig)
    out = [bar_png]

    # 2) map sized by nameplate capacity (Mt/yr of product)
    try:
        import warnings
        import geopandas as gpd
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            world = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    except Exception as e:  # noqa: BLE001
        print(f"map skipped: {e}")
        return out
    xmin, xmax, ymin, ymax = -95, -32, -56, 20
    world = world.cx[xmin - 30:xmax + 5, ymin - 5:ymax + 15]
    fig, ax = plt.subplots(figsize=(10, 12), dpi=150)
    world.plot(ax=ax, color="#EDEDED", edgecolor="white", linewidth=0.6)
    m = p[p["counted_as"] != "Closed"].dropna(subset=["lat", "lon"])
    m = m.assign(live=m["counted_as"] == "Operating", sec=m["sector"].where(m["sector"] != "Petrochemical", "Other"),
                 tonnage=m["capacity_mt_yr"].notna())
    # one marker per site, sector and live/not-live (Point Lisas holds ~20 plants at the same approximate point)
    g = m.groupby(["lat", "lon", "sec", "live", "tonnage", "country"], as_index=False)["capacity_mt_yr"].sum(min_count=1)
    k, fixed = 110, 14                      # marker area (pt^2) = k x Mt/yr; plants without tonnage get `fixed`

    def area(mt):
        return fixed if pd.isna(mt) else max(mt * k, fixed)

    def draw(axis, q, dx=None):
        q = q.assign(_a=q["capacity_mt_yr"].map(area)).sort_values("_a", ascending=False)   # small on top
        for _, r in q.iterrows():
            x, y = r["lon"], r["lat"]
            if dx is not None and abs(y - 10.40) < 0.02:   # spread only the Point Lisas estate cluster
                x, y = x + dx.get(r["sec"], (0, 0))[0], y + dx.get(r["sec"], (0, 0))[1]
            if r["live"]:
                axis.scatter(x, y, s=r["_a"], color=colours[r["sec"]], alpha=0.8, edgecolor="white", linewidth=0.7,
                             zorder=3, marker="o" if r["tonnage"] else "s")
            else:
                axis.scatter(x, y, s=r["_a"], facecolor="none", edgecolor=colours[r["sec"]], linewidth=1.2, zorder=4,
                             marker="o" if r["tonnage"] else "s")

    tt = g["country"] == "Trinidad and Tobago"
    draw(ax, g[~tt])
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_axis_off()
    ins = ax.inset_axes([0.60, 0.70, 0.38, 0.27])
    world.plot(ax=ins, color="#EDEDED", edgecolor="white", linewidth=0.6)
    off = {"Ammonia/urea": (-0.30, 0.20), "Methanol": (0.25, 0.22), "Steel DRI": (0.35, -0.08), "Cement": (0.12, -0.25)}
    draw(ins, g[tt], off)
    for lab, (x, y) in {"Point Lisas estate": (-61.20, 10.66), "Point Fortin (LNG)": (-61.45, 9.95),
                        "La Brea (CGCL)": (-62.35, 10.33)}.items():
        ins.text(x, y, lab, fontsize=7, color="#333333", ha="left")
    ins.set_xlim(-62.4, -60.6)
    ins.set_ylim(9.8, 10.95)
    ins.set_xticks([])
    ins.set_yticks([])
    for side in ins.spines.values():
        side.set_edgecolor("#888888")
        side.set_linewidth(0.6)
    ins.set_title("Trinidad & Tobago (markers spread by sector)", fontsize=8, color="#333333", pad=2)
    ax.indicate_inset((-62.4, 9.8, 1.8, 1.15), edgecolor="#888888", linewidth=0.6)
    present = [s for s in SECTOR_ORDER if s in set(m["sec"])]
    hs = [Line2D([], [], marker="o", linestyle="none", markersize=8, color=colours[s],
                 label="Other (petrochemical)" if s == "Other" else s) for s in present]
    hs += [Line2D([], [], marker="o", linestyle="none", markersize=8, markerfacecolor="none", markeredgecolor="#555555",
                  label="Idle, under construction or\nstatus unverified (hollow)"),
           Line2D([], [], marker="s", linestyle="none", markersize=fixed ** 0.5, color="#888888",
                  label="No tonnage capacity (ceramics in m2,\nor not published): fixed small square")]
    leg = ax.legend(handles=hs, loc="lower left", frameon=False, fontsize=8, bbox_to_anchor=(0.0, 0.14))
    ax.add_artist(leg)
    hs2 = [Line2D([], [], marker="o", linestyle="none", markersize=(v * k) ** 0.5, color="#BBBBBB",
                  label=f"{v:g} Mt/yr") for v in (0.5, 2, 5)]
    ax.legend(handles=hs2, loc="lower right", frameon=False, fontsize=8, labelspacing=2.4, borderpad=1.5,
              handletextpad=1.6, bbox_to_anchor=(1.0, 0.02), title="Nameplate capacity\n(Mt/yr of product; LNG MTPA)",
              title_fontsize=8)
    ax.set_title("Large industrial gas users: South & Central America and the Caribbean\n"
                 "(marker size = nameplate capacity; locations approximate)", loc="left", fontsize=12,
                 fontweight="bold")
    fig.text(0.02, 0.01, src + " Mexico out of scope.", fontsize=7, color="#666666")
    map_png = os.path.join(png_dir, f"{STEM}_map.png")
    fig.savefig(map_png, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    out.append(map_png)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default=CSV)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--png-dir", default=os.path.join(ROOT, "output", "PNG Charts"))
    ap.add_argument("--no-png", action="store_true")
    a = ap.parse_args()
    p = build_plants(load(a.csv))
    piv = pivot(p)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    xlsx_notes.write_workbook(a.out, {"Plants": p, "By country & sector": piv}, notes(p),
                              {"AS-OF", "UNITS", "BY COUNTRY & SECTOR",
                               "LARGEST OPERATING PLANTS BY CAPACITY (Mt/yr of product)", "REPORTED GAS USE",
                               "COMPLETENESS", "SOURCES"})
    print(f"Saved {a.out}: {len(p)} plants")
    print(piv.loc[[i for i in piv.index if i[1] == "All sectors"]].to_string())
    if not a.no_png:
        for f in render_pngs(p, a.png_dir):
            print(f"Saved {f}")


if __name__ == "__main__":
    main()
