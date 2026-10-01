"""
Register of large industrial natural-gas users in South America, Central America and the Caribbean:
fertiliser (ammonia/urea), methanol, steel DRI, alumina, cement (where gas-fired), glass, ceramics,
petrochemicals and LNG liquefaction. A reference list, not a time series.

The data live in a hand-curated CSV, south_america/industrial_gas_users.csv (one row per plant, every
capacity with its source URL). Edit the CSV, then rebuild:

    python3 south_america/INDUSTRIAL_GAS_USERS.py [--png-dir "output/PNG Charts"] [--no-png]

Writes output/Data and Chart Outputs/latin_america_industrial_gas_users.xlsx:
  Units               sources, method for the gas-demand estimates, completeness by sector, as-of date
  Plants              one row per plant, with the estimated gas demand
  By country & sector plant counts, capacity and estimated gas demand (mcm/d)
then add_charts.py adds the native chart (stacked bar of estimated gas demand by country and sector).
With --png-dir it also renders PNGs: the same bar chart and a map of the plants (geopandas 0.14
naturalearth_lowres, as south_america/COVERAGE_MAP.py).

Estimated gas demand is NOT measured: it is nameplate capacity x a standard gas intensity per sector
(INTENSITY below), i.e. demand at full rates, unless the row carries a company-stated gas volume
(gas_stated_mcm_d), which is used instead. Rows with no usable basis (ethane-fed crackers, plants whose
kiln fuel is unconfirmed, non-gas feedstock) get no estimate.
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

DAYS_PER_YEAR_TPD = 350        # t/d nameplate -> t/yr (allowing ~15 days of turnarounds)
GJ_PER_M3 = 38.0               # gross calorific value of pipeline gas, GJ per thousand m3 = MJ/m3

# Standard gas intensity per sector: (mcm of gas per unit of basis, basis unit, GJ/t-equivalent, note).
# mcm per kt == million m3 per thousand tonnes == m3 per tonne / 1000.
INTENSITY = {
    "Ammonia/urea": (1.00, "kt ammonia", "~38 GJ/t NH3 (feed + fuel, GCV); ~1.0 bcm per Mt ammonia. Urea "
                     "and UAN made from that ammonia are not counted again."),
    "Methanol": (1.00, "kt methanol", "~38 GJ/t (feed + fuel, GCV); ~1.0 Mt methanol per bcm."),
    "Steel DRI": (0.28, "kt DRI/HBI", "~10.5 GJ/t DRI for gas-based shaft furnaces (Midrex/HYL/Finmet)."),
    "LNG liquefaction": (1.45, "kt LNG", "~1,360 m3 of gas per tonne of LNG plus ~7% used as plant fuel."),
    "Alumina": (0.24, "kt alumina", "~9 GJ/t alumina (Bayer digestion steam + calcination), where gas is the fuel."),
    "Cement": (0.08, "kt cement", "~3.0 GJ/t cement (~3.5 GJ/t clinker x ~0.85 clinker factor), gas-fired kilns only."),
    "Glass": (0.17, "kt glass", "~6.5 GJ/t float glass (melting furnace)."),
    "Ceramics": (2.0, "million m2", "~0.076 GJ/m2 (~2 m3 of gas per m2 of tile); low confidence: dry-route "
                 "plants (Santa Gertrudes) use less, spray-dried wet route more."),
}
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
               "is listed without an estimate (fuel unconfirmed). Alumar, CBA, Alpart and Windalco are not listed "
               "(gas use unconfirmed).",
    "Cement": "Partial: only where gas firing is documented (Trinidad TCL, Bolivia) is gas estimated. Venezuela and "
              "Argentina plants are listed without an estimate (fuel mix unconfirmed). Brazil, Colombia, Chile, "
              "Peru, Ecuador, Central America and the Dominican Republic burn mainly coal/petcoke: not listed.",
    "Glass": "Partial: Brazil's four float-glass producers only. Container glass (O-I, Verallia...) and float glass "
             "elsewhere (Colombia, Argentina) are not listed.",
    "Ceramics": "Partial: Brazil's four largest tile groups (2024 production, ANFACER / Ceramic World Review), "
                "Colombia (Ceramica Italia, Corona Sopo new plant), Ecuador (Graiman, Italpisos). Brazil alone "
                "has ~60 companies / 71 plants (ANFACER); Peru, Argentina and Mexico-owned plants are not listed.",
    "Petrochemical": "Ethane-fed crackers only (Dow Bahia Blanca, Braskem Duque de Caxias); no gas estimate "
                     "because the feed is ethane (an NGL), not pipeline methane.",
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

COLS = ["country", "plant", "operator", "sector", "products", "capacity", "capacity_unit", "capacity_check", "status", "status_as_of",
        "city", "region", "lat", "lon", "coords", "basis_kt_or_Mm2_per_yr", "gas_mcm_d_estimate", "gas_estimate_basis",
        "counted_as", "notes", "sources"]


def load(path=CSV):
    d = pd.read_csv(path, dtype=str, keep_default_na=False)
    for c in ("capacity", "basis_value", "lat", "lon", "gas_stated_mcm_d"):
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


def estimate(d):
    """Annual basis (kt/yr or million m2/yr) and the gas estimate (mcm/d) per row."""
    basis = d["basis_value"].where(d["basis_unit"] != "t/d", d["basis_value"] * DAYS_PER_YEAR_TPD / 1000)
    gas, why = [], []
    for (_, r), b in zip(d.iterrows(), basis):
        if pd.notna(r["gas_stated_mcm_d"]):
            gas.append(round(float(r["gas_stated_mcm_d"]), 3))
            why.append(f"Stated: {r['gas_stated_basis']}")
        elif r["sector"] in INTENSITY and pd.notna(b):
            k, unit, _ = INTENSITY[r["sector"]]
            gas.append(round(b * k / 365, 3))
            src = f" ({r['basis_value']:g} t/d x {DAYS_PER_YEAR_TPD} d)" if r["basis_unit"] == "t/d" else ""
            why.append(f"Estimate: {b:,.0f} {unit}/yr{src} x {k:g} mcm per {unit} / 365, at nameplate")
        else:
            gas.append(None)
            why.append("No estimate (see notes)")
    return basis, gas, why


def build_plants(d):
    basis, gas, why = estimate(d)
    p = d.assign(**{"basis_kt_or_Mm2_per_yr": basis.round(1), "gas_mcm_d_estimate": gas, "gas_estimate_basis": why,
                    "counted_as": d["status"].map(counted_as)})
    p["sector"] = pd.Categorical(p["sector"], SECTOR_ORDER, ordered=True)
    p = p.sort_values(["country", "sector", "plant"]).reset_index(drop=True)
    p["sector"] = p["sector"].astype(str)
    p.index = pd.RangeIndex(1, len(p) + 1, name="id")
    return p[COLS]


def pivot(p):
    g = p.groupby(["country", "sector"], sort=False)
    out = pd.DataFrame({
        "plants": g.size(),
        "operating": g["counted_as"].apply(lambda s: (s == "Operating").sum()),
        "idle_closed": g["counted_as"].apply(lambda s: s.isin(["Idle", "Closed"]).sum()),
        "under_construction": g["counted_as"].apply(lambda s: (s == "Under construction").sum()),
        "capacity_basis_operating": p[p["counted_as"] == "Operating"].groupby(["country", "sector"])[
            "basis_kt_or_Mm2_per_yr"].sum(min_count=1),
        "gas_mcm_d_operating": p[p["counted_as"] == "Operating"].groupby(["country", "sector"])[
            "gas_mcm_d_estimate"].sum(min_count=1),
        "gas_mcm_d_idle": p[p["counted_as"] == "Idle"].groupby(["country", "sector"])["gas_mcm_d_estimate"].sum(min_count=1),
        "gas_mcm_d_under_construction": p[p["counted_as"] == "Under construction"].groupby(["country", "sector"])[
            "gas_mcm_d_estimate"].sum(min_count=1),
    })
    out["capacity_basis_unit"] = [INTENSITY.get(s, (None, "", ""))[1] + ("/yr" if s in INTENSITY else "")
                                  for _, s in out.index]
    out = out.reset_index()
    tot = out.groupby("country", sort=False)[["plants", "operating", "idle_closed", "under_construction",
                                              "gas_mcm_d_operating", "gas_mcm_d_idle",
                                              "gas_mcm_d_under_construction"]].sum(min_count=1).reset_index()
    tot["sector"] = "All sectors"
    out = pd.concat([out, tot], ignore_index=True)
    out["_o"] = out["sector"].map({s: i for i, s in enumerate(SECTOR_ORDER + ["All sectors"])})
    out = out.sort_values(["country", "_o"]).drop(columns="_o")
    grand = out[out["sector"] == "All sectors"].drop(columns=["country", "sector"]).sum(numeric_only=True, min_count=1)
    grand["country"], grand["sector"] = "ALL", "All sectors"
    out = pd.concat([out, grand.to_frame().T], ignore_index=True)
    cols = ["country", "sector", "plants", "operating", "idle_closed", "under_construction",
            "capacity_basis_operating", "capacity_basis_unit", "gas_mcm_d_operating", "gas_mcm_d_idle",
            "gas_mcm_d_under_construction"]
    out = out[cols]
    for c in cols[2:6]:
        out[c] = pd.to_numeric(out[c]).astype("Int64")
    for c in cols[6:]:
        if c != "capacity_basis_unit":
            out[c] = pd.to_numeric(out[c], errors="coerce").round(2)
    return out.set_index(["country", "sector"])


def trinidad_check(p):
    """Compare the Trinidad estimates with MEEI's measured gas use by sector (trinidad_gas.xlsx), when present."""
    path = os.path.join(ROOT, "output", "Data and Chart Outputs", "trinidad_gas.xlsx")
    lines = []
    if not os.path.exists(path):
        return ["Trinidad cross-check skipped: trinidad_gas.xlsx not found in this checkout."]
    u = pd.read_excel(path, sheet_name="Utilization by sector")
    u["date"] = pd.to_datetime(u["date"])
    last = u["date"].max()
    recent = u[u["date"] > last - pd.DateOffset(months=6)]
    measured = recent.groupby("sector")["mmscfd"].mean() / 35.3147   # MMscf/d -> mcm/d
    tt = p[(p["country"] == "Trinidad and Tobago") & (p["counted_as"] == "Operating")]
    est = tt.groupby("sector")["gas_mcm_d_estimate"].sum()
    pairs = [("Ammonia/urea", ["Ammonia Manufacture", "Ammonia Derivatives"]), ("Methanol", ["Methanol Manufacture"]),
             ("LNG liquefaction", ["LNG"]), ("Steel DRI", ["Iron & Steel Manufacture"]), ("Cement", ["Cement Manufacture"])]
    lines.append(f"Trinidad cross-check: register estimate at nameplate (operating plants) vs MEEI measured gas use, "
                 f"mean of the 6 months to {last:%b-%Y} (mcm/d):")
    for sec, meei in pairs:
        m = measured.reindex(meei).sum()
        lines.append(f"  {sec}: estimate {est.get(sec, 0):.1f} vs MEEI {m:.1f}")
    lines.append("  (Methanol: Titan was idled from 29-Jun-2026, so the 6-month MEEI mean still includes it; LNG runs "
                 "below nameplate for lack of feed gas.)")
    return lines


def notes(p, piv):
    asof = dt.date.today().isoformat()
    op = p[p["counted_as"] == "Operating"]
    top = op.sort_values("gas_mcm_d_estimate", ascending=False).head(10)
    lines = ["Latin America & Caribbean industrial gas users - plant register", "",
             "AS-OF", f"Register compiled {asof}; each row's status carries its own as-of date (status_as_of).",
             "A reference list of large gas-using industrial plants, not a time series. Mexico is out of scope.",
             "Edit south_america/industrial_gas_users.csv and rerun south_america/INDUSTRIAL_GAS_USERS.py to rebuild.", "",
             "UNITS",
             "capacity / capacity_unit: nameplate as published by the source (kt/yr, t/d, Mt/yr, million m2/yr).",
             "capacity_check: whether an automated fetch found the figure on a cited page (Oct-2026); 'UNVERIFIED' "
             "rows rest on figures whose pages block automated access - treat with care.",
             "basis_kt_or_Mm2_per_yr: the capacity the gas estimate is built on, annualised (t/d x 350 days).",
             "gas_mcm_d_estimate: ESTIMATED gas demand in million cubic metres per day (1 mcm/d = 35.3 MMscf/d) at "
             "nameplate - not measured.",
             "lat / lon: approximate (see 'coords': GEM wiki plant coordinates where available, otherwise town or "
             "industrial-estate level).", "",
             "METHOD",
             "Gas demand = annual basis capacity x standard intensity / 365, i.e. full-rate demand. Standard intensities:"]
    for s, (k, unit, why) in INTENSITY.items():
        lines.append(f"  {s}: {k:g} mcm per {unit} - {why}")
    lines += ["Where a company states its gas use or contract volume, that figure is used instead (gas_estimate_basis "
              "says which).",
              f"Energy conversions assume pipeline gas at {GJ_PER_M3:g} MJ/m3 (gross).",
              "Double counting: Ecuador's Bajo Alto LNG plant liquefies the gas that Cuenca's ceramics plants burn, so "
              "their estimates overlap (both small). LNG plants' gas is feed for export, not domestic end use.",
              "Pivot ('By country & sector'): plant counts by status; capacity and gas summed over plants counted as "
              "operating; idle and under-construction gas shown separately. Capacity sums mix units across sectors, "
              "so read them within a sector.", ""]
    lines += ["CROSS-CHECK"] + trinidad_check(p) + [""]
    lines += ["LARGEST ESTIMATED USERS (operating)"]
    for _, r in top.iterrows():
        lines.append(f"  {r['gas_mcm_d_estimate']:.2f} mcm/d - {r['plant']} ({r['country']}, {r['sector']})")
    lines += ["", "COMPLETENESS"]
    for s in SECTOR_ORDER:
        if s in COMPLETENESS:
            lines.append(f"  {s}: {COMPLETENESS[s]}")
    lines += ["Countries with no plant listed: Uruguay, Paraguay, Central America (Guatemala-Panama, Belize), "
              "Dominican Republic, Puerto Rico - gas there goes to power; cement kilns burn coal/petcoke.", "",
              "SOURCES"] + SOURCES + ["Every row of 'Plants' carries its own source URL(s); statuses marked "
                                      "'unverified' were not confirmed by a dated source."]
    return lines


def render_pngs(p, png_dir):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    sys.path.insert(0, ROOT)
    from xlsx_charts import PALETTE
    os.makedirs(png_dir, exist_ok=True)
    colours = {s: "#" + c for s, c in zip([s for s in SECTOR_ORDER if s in INTENSITY], PALETTE)}

    # 1) stacked bar, the same table as the native chart
    t = chart_table(p)
    fig, ax = plt.subplots(figsize=(11, 6), dpi=150)
    bottom = pd.Series(0.0, index=t.index)
    for s in t.columns:
        ax.bar(t.index, t[s], bottom=bottom, color=colours[s], label=s, width=0.65, edgecolor="white", linewidth=1)
        bottom += t[s]
    for x, v in zip(t.index, bottom):
        ax.text(x, v + bottom.max() * 0.01, f"{v:.1f}", ha="center", va="bottom", fontsize=8, color="#333333")
    ax.set_ylim(0, bottom.max() * 1.08)
    ax.set_ylabel("million m3/day (estimate at nameplate)")
    ax.set_title("Estimated gas demand of operating industrial plants, by country and sector", loc="left",
                 fontsize=12, fontweight="bold")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.grid(axis="y", color="#E5E5E5", linewidth=0.6)
    ax.set_axisbelow(True)
    plt.setp(ax.get_xticklabels(), rotation=0, fontsize=9)
    ax.legend(frameon=False, fontsize=8, ncol=4, loc="upper right")
    fig.text(0.01, 0.005, "Source: plant register (USGS Minerals Yearbook, MEEI Trinidad, GEM, company reports); "
             "gas = capacity x standard intensity (see Units tab).", fontsize=7, color="#666666")
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    bar_png = os.path.join(png_dir, f"{STEM}__Chart.png")
    fig.savefig(bar_png, facecolor="white")
    plt.close(fig)

    # 2) map
    out = [bar_png]
    try:
        import warnings
        import geopandas as gpd
        from matplotlib.lines import Line2D
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
    m = p[p["counted_as"].isin(["Operating", "Idle", "Under construction", "Status unverified"])].dropna(subset=["lat", "lon"])
    m = m.assign(live=m["counted_as"] == "Operating", gas=m["gas_mcm_d_estimate"].fillna(0))
    # one marker per site, sector and live/not-live (Point Lisas holds ~20 plants at the same approximate point)
    g = m.groupby(["lat", "lon", "sector", "live", "country"], as_index=False)["gas"].sum()
    allc = dict(colours, **{"Petrochemical": "#7A5C3E", "Other": "#9A9A9A"})
    area = lambda v: v.clip(lower=0.03) * 90 + 12   # noqa: E731  marker area (pt^2) ~ mcm/d

    def draw(axis, q, dx=None):
        q = q.sort_values("gas", ascending=False)        # big markers first, small ones stay visible on top
        for _, r in q.iterrows():
            x, y = r["lon"], r["lat"]
            if dx is not None and abs(y - 10.40) < 0.02:   # spread only the Point Lisas estate cluster
                x, y = x + dx.get(r["sector"], (0, 0))[0], y + dx.get(r["sector"], (0, 0))[1]
            if r["live"]:
                axis.scatter(x, y, s=area(pd.Series([r["gas"]]))[0], color=allc[r["sector"]], alpha=0.8,
                             edgecolor="white", linewidth=0.7, zorder=3)
            else:
                axis.scatter(x, y, s=area(pd.Series([r["gas"]]))[0], facecolor="none", edgecolor=allc[r["sector"]],
                             linewidth=1.2, zorder=4)

    tt = g["country"] == "Trinidad and Tobago"
    draw(ax, g[~tt])
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_axis_off()
    # Trinidad inset: the Point Lisas plants spread out by sector so each is visible
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
    present = [s for s in SECTOR_ORDER if s in set(m["sector"])]
    hs = [Line2D([], [], marker="o", linestyle="none", markersize=8, color=allc[s], label=s) for s in present]
    hs += [Line2D([], [], marker="o", linestyle="none", markersize=8, markerfacecolor="none", markeredgecolor="#555555",
                  label="Idle, under construction or\nstatus unverified (hollow)")]
    leg = ax.legend(handles=hs, loc="lower left", frameon=False, fontsize=8, bbox_to_anchor=(0.0, 0.16))
    ax.add_artist(leg)
    hs2 = [Line2D([], [], marker="o", linestyle="none", markersize=float(area(pd.Series([v]))[0]) ** 0.5,
                  color="#BBBBBB", label=f"{v:g} mcm/d") for v in (1, 5, 15)]
    ax.legend(handles=hs2, loc="lower right", frameon=False, fontsize=8, labelspacing=2.6, borderpad=1.5,
              handletextpad=1.6, bbox_to_anchor=(1.0, 0.02), title="Estimated gas demand", title_fontsize=8)
    ax.set_title("Large industrial gas users: South & Central America and the Caribbean\n"
                 "(size = estimated gas demand at nameplate; locations approximate)", loc="left", fontsize=12,
                 fontweight="bold")
    fig.text(0.02, 0.01, "Source: plant register (USGS, MEEI Trinidad, GEM, company reports). Mexico out of scope.",
             fontsize=7, color="#666666")
    map_png = os.path.join(png_dir, f"{STEM}_map.png")
    fig.savefig(map_png, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    out.append(map_png)
    return out


def chart_table(p):
    """Operating plants: estimated gas (mcm/d), countries (largest first) x sectors with an estimate."""
    op = p[(p["counted_as"] == "Operating") & p["gas_mcm_d_estimate"].notna()]
    t = op.pivot_table(index="country", columns="sector", values="gas_mcm_d_estimate", aggfunc="sum")
    t = t[[s for s in SECTOR_ORDER if s in t.columns]]
    t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    t.index = [c.replace("Trinidad and Tobago", "Trinidad & Tobago") for c in t.index]
    return t.fillna(0).round(3)


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
    xlsx_notes.write_workbook(a.out, {"Plants": p, "By country & sector": piv}, notes(p, piv),
                              {"AS-OF", "UNITS", "METHOD", "CROSS-CHECK", "LARGEST ESTIMATED USERS (operating)",
                               "COMPLETENESS", "SOURCES"})
    print(f"Saved {a.out}: {len(p)} plants")
    print(piv.loc[[i for i in piv.index if i[1] == "All sectors"]].to_string())
    if not a.no_png:
        for f in render_pngs(p, a.png_dir):
            print(f"Saved {f}")


if __name__ == "__main__":
    main()
