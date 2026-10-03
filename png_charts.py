"""
PNG copies of a workbook's charts, for output/PNG Charts/: the same chart specs add_charts.py draws natively in
Excel (add_charts.REGISTRY), rendered with matplotlib in the same palette - stacked bars / lines / areas with
mmm/yy dates (years for annual series), and the AGSI-style water-year chart (5-year min-max band, 5-year average,
previous and current water year, Oct-Sep). No chart borders.

    python3 png_charts.py "output/Data and Chart Outputs/thailand_hydro_reservoirs.xlsx" [...] [--out "output/PNG Charts"]

Writes <workbook stem>__<chart name>.png per chart.
"""
import argparse
import os
import re
import sys

import matplotlib

matplotlib.use("Agg")
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
    df = xlsx_charts._fold_to_palette(spec["df"].dropna(how="all").apply(pd.to_numeric, errors="coerce"))
    if df.empty:
        return False
    fig, ax = plt.subplots(figsize=(10, 5))
    annual = spec.get("date_format") == "%Y"
    x = df.index
    if spec["kind"] == "stacked_bar":
        width = 300 if annual else (25 if len(df) < 400 else 1)
        pos, neg = pd.Series(0.0, index=x), pd.Series(0.0, index=x)
        for i, c in enumerate(df.columns):
            v = df[c].fillna(0)
            base = pos.where(v >= 0, neg)
            ax.bar(x, v, width=width, bottom=base, color=_colour(c, i), label=str(c), align="edge" if not annual else "center")
            pos, neg = pos + v.clip(lower=0), neg + v.clip(upper=0)
    elif spec["kind"] == "stacked_area":
        ax.stackplot(x, *[df[c].fillna(0) for c in df.columns], labels=[str(c) for c in df.columns],
                     colors=[_colour(c, i) for i, c in enumerate(df.columns)])
    else:
        for i, c in enumerate(df.columns):
            ax.plot(x, df[c], color=_colour(c, i), label=str(c), linewidth=1.4)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y" if annual else "%b/%y"))
    _style(ax, spec["title"], spec["units"])
    ax.legend(fontsize=8, frameon=False, ncol=min(len(df.columns), 4), loc="upper left", bbox_to_anchor=(0, -0.1))
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
            ok = water_year(spec, out) if "water_year" in spec else series_chart(spec, out)
        except Exception as e:  # noqa: BLE001
            print(f"  {stem} {name}: {type(e).__name__}: {e}")
            continue
        if ok:
            done.append(out)
    print(f"{stem}: {len(done)} PNG(s)")
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbooks", nargs="+")
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    for p in args.workbooks:
        render(p, args.out)


if __name__ == "__main__":
    main()
