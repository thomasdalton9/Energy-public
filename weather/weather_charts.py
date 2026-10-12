"""Chart specs for the '<region>_degree_days.xlsx' workbooks (weather/degree_days.py), used by add_charts.REGISTRY and
weather/WEATHER_MASTER.py. Which countries are charted is set per region in weather/regions.json ('charts').

  cumulative HDD since the season start (Oct, or the region's season_start_month) as an AGSI-style season chart: 5-year
  min-max band, 5-year average, previous and current season, plus the 1991-2020 normal line (water_year_chart);
  monthly CDD vs the 1991-2020 normal; monthly HDD minus the normal (anomaly), base 18 C.
"""
import json
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)
import add_charts  # noqa: E402
import degree_days  # noqa: E402

MONTH_NAMES = {1: "1 Jan", 2: "1 Feb", 3: "1 Mar", 4: "1 Apr", 5: "1 May", 6: "1 Jun", 7: "1 Jul", 8: "1 Aug", 9: "1 Sep",
               10: "1 Oct", 11: "1 Nov", 12: "1 Dec"}
CHART_FROM = "2019-01-01"


def region_of(path):
    return os.path.basename(path).replace("_degree_days.xlsx", "")


def _cum_since(x, start_month):
    """Daily degree days -> cumulative since the season start; a season's cumulative stops at its first missing day."""
    last = x.last_valid_index()
    if last is None:
        return x.iloc[0:0]
    x = x.loc[:last]
    full = x.reindex(pd.date_range(x.index.min(), last, freq="D"))
    wy = pd.Series([d.year if d.month >= start_month else d.year - 1 for d in full.index], index=full.index)
    out = full.groupby(wy).cumsum()
    gap = full.isna().groupby(wy).cummax()
    return out.where(~gap).dropna()


def labels(path):
    region = region_of(path)
    cfg = degree_days.load_regions()[region]
    cs = pd.read_csv(os.path.join(HERE, cfg["countries"]), keep_default_na=False)
    names = dict(zip(cs["country_code"], cs["country"]))
    names.update(cfg.get("aggregates", {}))
    return cfg, names


def specs(path):
    cfg, names = labels(path)
    region = region_of(path)
    sm = int(cfg.get("season_start_month", 10))
    ch = cfg["charts"]
    hdd = add_charts.by_date(add_charts.read(path, "Daily HDD"), "date")
    norm_d = add_charts.read(path, "Normals doy").set_index("day of year (leap-year numbering)")
    monthly = add_charts.by_date(add_charts.read(path, "Monthly"), "month")
    norm_m = add_charts.read(path, "Normals month").set_index("month")
    out = []
    title_region = cfg["title"]
    for k in ch.get("cumulative_hdd", []):
        col = f"{k} 18C"
        if col not in hdd or not hdd[col].notna().any():
            continue
        s = _cum_since(hdd[col], sm)
        s.attrs["season_start_month"] = sm
        s.attrs["normal"] = degree_days.season_normal_cumulative(norm_d[f"{k} HDD 18C"], sm)
        out.append({"name": f"HDD {k}", "water_year": s, "y_decimals": 0, "sheet": f"HDD season - {k}",
                    "title": f"Cumulative HDD since {MONTH_NAMES[sm]}: {names.get(k, k)} vs 1991-2020 normal",
                    "units": "degree days (base 18 C), cumulative"})
    for k in ch.get("monthly_cdd", []):
        col = f"{k} CDD 18C"
        if col not in monthly:
            continue
        m = monthly[[col]].dropna()
        m = m[m.index >= CHART_FROM]
        nm = pd.Series(m.index.month.map(norm_m[col]), index=m.index, dtype=float)
        df = pd.DataFrame({"CDD (base 18 C)": m[col], "1991-2020 normal": nm}).dropna(how="all")
        out.append(add_charts.spec(f"CDD {k}", df, f"Monthly CDD vs 1991-2020 normal: {names.get(k, k)}",
                                   "degree days per month (base 18 C)", "line", "%b/%y"))
    anom = add_charts.by_date(add_charts.read(path, "Anomaly"), "month")
    for k in ch.get("anomaly_hdd", []):
        col = f"{k} HDD 18C abs"
        if col not in anom:
            continue
        a = anom[[col]].dropna()
        a = a[a.index >= CHART_FROM].rename(columns={col: "HDD minus 1991-2020 normal"})
        out.append(add_charts.spec(f"HDD anomaly {k}", a, f"Monthly HDD vs 1991-2020 normal (actual minus normal): {names.get(k, k)}",
                                   "degree days per month (base 18 C)", "stacked_bar", "%b/%y"))
    return out
