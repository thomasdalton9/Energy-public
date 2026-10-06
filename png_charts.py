"""
PNG copies of a workbook's charts, for output/PNG Charts/: the same chart specs add_charts.py draws natively in
Excel (add_charts.REGISTRY), rendered with matplotlib in the same palette - stacked bars / lines / areas with
mmm/yy dates (years for annual series), and the AGSI-style water-year chart (5-year min-max band, 5-year average,
previous and current water year, Oct-Sep). No chart borders.

    python3 png_charts.py "output/Data and Chart Outputs/thailand_hydro_reservoirs.xlsx" [...] [--out "output/PNG Charts"]
    python3 png_charts.py <master.xlsx> --sheet "SSEA generation total data" --title "..." --units "TWh per month"

Writes <workbook stem>__<chart name>.png per chart (with --sheet: <workbook stem>__<sheet>.png, a stacked bar of
that chart-data tab - date in the first column, one series per column).
"""
import argparse
import os
import re
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.ticker  # noqa: E402
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import add_charts  # noqa: E402
import water_year_chart  # noqa: E402
import xlsx_charts  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "PNG Charts")
COLOURS = ["#" + c for c in xlsx_charts.PALETTE]
GREY = "#" + xlsx_charts.OTHER_GREY


def _style(ax, title, units):
    ax.set_title(title, fontsize=11, loc="left")
    ax.set_ylabel(units, fontsize=9)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.tick_params(labelsize=8)
    ax.grid(axis="y", color="#e5e5e5", linewidth=0.6)
    ax.set_axisbelow(True)


def _colour(name, i):
    return GREY if str(name) == "Other" else COLOURS[i % len(COLOURS)]


def series_chart(spec, path):
    df = spec["df"].dropna(how="all").apply(pd.to_numeric, errors="coerce")
    df = xlsx_charts.prepare(df, spec.get("line_cols", ()) if spec["kind"] != "line" else ())[0]   # fold only the bars, not the lines
    if df.empty:
        return False
    fig, ax = plt.subplots(figsize=(10, 5))
    annual = spec.get("date_format") == "%Y"
    x = df.index
    ff = pd.Timestamp(spec["forecast_from"]) if spec.get("forecast_from") is not None else None
    sf = pd.Timestamp(spec["scenario_from"]) if spec.get("scenario_from") is not None else None
    line_cols = [c for c in spec.get("line_cols", ()) if c in df.columns] if spec["kind"] != "line" else []
    body = df.drop(columns=line_cols)
    if spec["kind"] == "stacked_bar":
        width = 300 if annual else (25 if len(df) < 400 else 1)
        pos, neg = pd.Series(0.0, index=x), pd.Series(0.0, index=x)
        for i, c in enumerate(body.columns):
            v = body[c].fillna(0)
            base = pos.where(v >= 0, neg)
            alphas = [(0.28 if (sf is not None and d >= sf) else 0.5) if (ff is not None and d >= ff) else 1.0 for d in x]
            bars = ax.bar(x, v, width=width, bottom=base, color=_colour(c, i), label=str(c), align="edge" if not annual else "center")
            for b, al in zip(bars, alphas):
                b.set_alpha(al)
            pos, neg = pos + v.clip(lower=0), neg + v.clip(upper=0)
    elif spec["kind"] == "stacked_area":
        ax.stackplot(x, *[body[c].fillna(0) for c in body.columns], labels=[str(c) for c in body.columns],
                     colors=[_colour(c, i) for i, c in enumerate(body.columns)])
    else:
        for i, c in enumerate(body.columns):
            ax.plot(x, body[c], color=_colour(c, i), label=str(c), linewidth=1.8,
                    linestyle="--" if "elayed" in str(c) else "-")
    for j, c in enumerate(line_cols):
        col, ls = (spec.get("line_styles") or {}).get(str(c), ("252525", ("solid", "dash", "sysDot")[j % 3]))
        ax.plot(x, df[c], color="#" + col, linewidth=1.6, linestyle={"solid": "-", "dash": "--", "sysDot": ":"}[ls], label=str(c))
    if ff is not None:
        ax.axvline(ff, color="#555555", linewidth=0.9, linestyle=":")
        ax.text(ff, ax.get_ylim()[1], " forecast", fontsize=8, color="#555555", va="top")
    if sf is not None:
        ax.axvline(sf, color="#555555", linewidth=0.9, linestyle="--")
        ax.text(sf, ax.get_ylim()[1], " scenario (not a forecast)", fontsize=8, color="#555555", va="top")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y" if annual else "%b/%y"))
    if spec["kind"] == "stacked_bar" and not annual and len(df) <= 6:      # a handful of snapshots: one tick per bar, not a weekly grid
        ax.set_xticks([d + pd.Timedelta(days=12.5) for d in x])
    _style(ax, spec["title"], spec["units"])
    if str(spec["units"]).startswith("Bcf/d") and "MISO" in spec["title"]:
        ax.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.2f"))
    ax.legend(fontsize=8, frameon=False, ncol=min(len(df.columns), 4 if max(len(str(c)) for c in df.columns) <= 30 else 3), loc="upper left", bbox_to_anchor=(0, -0.1))
    fig.tight_layout()
    fig.savefig(path, dpi=110)
    plt.close(fig)
    return True


def water_year(spec, path):
    table, meta = water_year_chart.water_year_table(spec["water_year"])
    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(table))
    lo, hi = table["5Y min"], table["5Y max"]
    ax.fill_between(x, lo, hi, color="#d9e4f2", label=f"5-year range ({meta['hist']})")
    ax.plot(x, table["5Y average"], color="#7f7f7f", linestyle="--", linewidth=1.2, label="5-year average")
    wy = [c for c in table.columns if str(c).startswith("WY ")]
    for c, col, lw in zip(wy, ("#EB6834", "#2A78D6"), (1.4, 2.2)):
        ax.plot(x, table[c], color=col, linewidth=lw, label=str(c))
    ticks = [i for i, d in enumerate(table["Day"]) if str(d).startswith("01-")]
    ax.set_xticks(ticks, [pd.Timestamp(f"2001-{str(table['Day'][i])[3:]}-01").strftime("%b") for i in ticks])
    _style(ax, spec["title"] + " (water year Oct-Sep)", spec["units"])
    ax.legend(fontsize=8, frameon=False, ncol=4, loc="upper left", bbox_to_anchor=(0, -0.08))
    fig.tight_layout()
    fig.savefig(path, dpi=110)
    plt.close(fig)
    return True


def haynesville_curve(xlsx, path):
    """Haynesville cost curve (americas/HAYNESVILLE_COST_CURVE.py): step chart of breakeven vs cumulative capacity for two horizons, from the script's
    'Curve values' tab, with the Henry Hub 12-month average."""
    c = pd.read_excel(xlsx, "Curve values").set_index("Rank")
    hh = float(pd.read_excel(xlsx, "Henry Hub", index_col=0).iloc[1, 0])
    fig, ax = plt.subplots(figsize=(11, 6))
    for (H, col, ls, lab) in ((5, COLOURS[0], "-", "by Dec 2030 (drilled over 5 years)"), (8, COLOURS[1], "-", "by Dec 2033 (drilled over 8 years)")):
        cum = pd.concat([pd.Series([0.0]), c[f"Cumulative capacity H={H} (Bcf/d)"].reset_index(drop=True)], ignore_index=True)
        be = c["Breakeven $/MMBtu HH"].reset_index(drop=True)
        xs, ys = [], []
        for k in range(len(be)):
            xs += [cum[k], cum[k + 1]]
            ys += [be[k], be[k]]
        ax.plot(xs, ys, color=col, linewidth=2.2, linestyle=ls, label=f"Supply added {lab}")
        if H == 8:
            for k in range(len(be)):
                ax.text((cum[k] + cum[k + 1]) / 2, be[k] + 0.12, c["Tier (sorted by breakeven)"].iloc[k].replace(" (BEG)", "").replace("Western Haynesville - ", "WH - "),
                        fontsize=7, ha="center", color="#444444", rotation=0 if cum[k + 1] - cum[k] > 4 else 90, va="bottom")
    ax.axhline(hh, color="#252525", linestyle="--", linewidth=1.2, label=f"Henry Hub, last 12 months average (${hh:.2f})")
    ax.set_xlim(left=0)
    ax.set_ylim(0, 11)
    ax.set_xlabel("Cumulative capacity, Bcf/d (cost only: NOT constrained by rigs, crews or takeaway)", fontsize=9)
    _style(ax, "Haynesville supply cost curve - APPROXIMATE, built from company disclosures (pre-tax, 10% discount)", "Breakeven, $/MMBtu Henry Hub")
    ax.legend(fontsize=8, frameon=False, ncol=3, loc="upper left", bbox_to_anchor=(0, -0.12))
    fig.text(0.01, 0.005, "Sources: Comstock Oct 2026 deck + Q2 2026 release (inventory, D&C $/ft, opex); Expand 3Q25 deck (productivity); BEG/OGJ Dec 2015 (tiers).\n"
             "Western Haynesville EUR ASSUMED = Tier 1. Legacy inventory is 2012-vintage (wells since not deducted). See the workbook's Sources tab.", fontsize=6.5, color="#555555")
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    fig.savefig(path, dpi=110)
    plt.close(fig)
    return True


def haynesville_marginal(xlsx, path):
    """Marginal breakeven at each Gulf 'Key question' date (STEO growth + extra supply), base and delayed, against the Henry Hub 12-month average."""
    m = pd.read_excel(xlsx, "Mapping values")
    hh = float(pd.read_excel(xlsx, "Henry Hub", index_col=0).iloc[1, 0])
    col = "Marginal breakeven - STEO growth + extra"
    yrs = list(dict.fromkeys(m["Year-end"]))
    fig, ax = plt.subplots(figsize=(10, 5.5))
    w = 0.36
    for i, (sc, lab) in enumerate((("Base", "Base case"), ("LNG and takeaway delayed 6 months", "LNG and takeaway delayed 6 months"))):
        v = [m[(m["Scenario"] == sc) & (m["Year-end"] == y)][col].iloc[0] for y in yrs]
        xs = [k + (i - 0.5) * w for k in range(len(yrs))]
        bars = ax.bar(xs, [0 if pd.isna(x) else x for x in v], width=w, color=COLOURS[i], label=lab)
        for x, y in zip(xs, v):
            ax.text(x, (0 if pd.isna(y) else y) + 0.05, "beyond inventory" if pd.isna(y) else f"{y:.2f}", ha="center", fontsize=8)
    ax.axhline(hh, color="#252525", linestyle="--", linewidth=1.2, label=f"Henry Hub, last 12 months average (${hh:.2f})")
    ax.set_xticks(range(len(yrs)), yrs)
    ax.set_ylim(0, max(6, hh + 1.5))
    _style(ax, "Implied marginal breakeven of Haynesville growth, Gulf 'Key question' (cost only, APPROXIMATE)", "$/MMBtu Henry Hub")
    ax.legend(fontsize=8, frameon=False, ncol=3, loc="upper left", bbox_to_anchor=(0, -0.08))
    fig.text(0.01, 0.005, "Cost curve only: no rig, crew or takeaway limit, no base-decline replacement (lower bound). Western Haynesville productivity ASSUMED; see 'Sensitivity values' ($2.8-5.1).", fontsize=6.5, color="#555555")
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    fig.savefig(path, dpi=110)
    plt.close(fig)
    return True


PNG_RENDERERS = {"haynesville_curve": haynesville_curve, "haynesville_marginal": haynesville_marginal}


def render(path, out_dir=OUT):
    stem = os.path.splitext(os.path.basename(path))[0]
    build = add_charts.REGISTRY.get(os.path.basename(path))
    if build is None:
        print(f"{stem}: not in add_charts.REGISTRY - skipped")
        return []
    os.makedirs(out_dir, exist_ok=True)
    done = []
    for spec in build(path):
        name = re.sub(r"[^A-Za-z0-9_-]+", "_", spec.get("sheet") or spec["name"]).strip("_")
        out = os.path.join(out_dir, f"{stem}__{name}.png")
        try:
            ok = (water_year(spec, out) if "water_year" in spec else PNG_RENDERERS[spec["png"]](path, out) if "png" in spec
                  else series_chart(spec, out))
        except Exception as e:  # noqa: BLE001
            print(f"  {stem} {name}: {type(e).__name__}: {e}")
            continue
        if ok:
            done.append(out)
    print(f"{stem}: {len(done)} PNG(s)")
    return done


def render_sheet(path, sheet, title, units, out_dir=OUT, kind="stacked_bar"):
    """A master's chart-data tab (first column dates, then one column per series) -> PNG."""
    d = pd.read_excel(path, sheet_name=sheet)
    d = d.set_index(pd.to_datetime(d.iloc[:, 0], errors="coerce")).iloc[:, 1:]
    d = d[d.index.notna()].apply(pd.to_numeric, errors="coerce").dropna(axis=1, how="all")
    d = d[[c for c in d.columns if not str(c).startswith("Unnamed")]]
    stem = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(out_dir, f"{stem}__{re.sub(r'[^A-Za-z0-9_-]+', '_', sheet).strip('_')}.png")
    os.makedirs(out_dir, exist_ok=True)
    annual = len(d) > 1 and d.index.to_series().diff().median().days > 300   # annual tabs: years on the axis
    if series_chart({"df": d, "title": title, "units": units, "kind": kind,
                     "date_format": "%Y" if annual else "%Y-%m"}, out):
        print(f"{stem}: {out}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbooks", nargs="+")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--sheet", help="render this chart-data tab instead of the add_charts specs")
    ap.add_argument("--title", default="")
    ap.add_argument("--units", default="")
    ap.add_argument("--kind", default="stacked_bar", help="with --sheet: stacked_bar (default) or line")
    args = ap.parse_args()
    for p in args.workbooks:
        if args.sheet:
            render_sheet(p, args.sheet, args.title or args.sheet, args.units, args.out, args.kind)
        else:
            render(p, args.out)


if __name__ == "__main__":
    main()
