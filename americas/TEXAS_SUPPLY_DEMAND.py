"""
Texas supply and demand to Dec 2028 (Bcf/d, monthly): the 'Supply and demand' tab of texas_production_forecast.xlsx.
Called by TEXAS_PRODUCTION_FORECAST.py (nothing is written here; read-only on texas_gas_monthly.xlsx).

  DEMAND (stacked)  five EIA consumption sectors + LNG feedgas + pipeline exports to Mexico. Actual months are EIA; forecast months
                    are the 'Forecast values' tab of texas_gas_monthly.xlsx (TEXAS_GAS_FORECAST.py: LNG feedgas base / delayed,
                    seasonal-trend sectors and Mexico). LNG feedgas = EIA exports x 1.09 (liquefaction fuel and shrink).
  PRODUCTION        DRY gas: EIA dry production (history; after Dec 2024 marketed less the last-12-month extraction-loss share, as
                    TEXAS_GAS.py does) and, for the forecast, the marketed forecast (STEO regions, Permian takeaway cap, base and
                    takeaway-delayed) x (1 - that share). Marketed gas includes the NGLs a plant removes (about 16% of it), which
                    never enters a pipeline, so the gap to demand is taken on dry gas.
  IMPLIED OUTFLOW   dry production - demand stack = net gas leaving Texas by pipeline to other states, plus lease/plant and pipeline
                    fuel EIA withholds, less any storage withdrawal (storage is not forecast, so it is left out of history too).
                    Base = base production with base LNG; delayed = takeaway-delayed production with LNG-delayed feedgas (both
                    delays are the 6-month cases of their workbooks). Sectors and Mexico are the same in both.
"""
import pandas as pd

SECT = ["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel"]
SHEET = "Supply and demand"
LNG_B, LNG_D = "LNG feedgas total - base", "LNG feedgas total - delayed"
MEX = "Pipeline exports to Mexico"
START = pd.Timestamp("2021-01-01")


def load_gas(texas_xlsx):
    fv = pd.read_excel(texas_xlsx, sheet_name="Forecast values")
    fv["Month"] = pd.to_datetime(fv["Month"])
    return fv.set_index("Month").sort_index()


def build(idx, hist_dry, mkt_base, mkt_delayed, loss_share, fv):
    """idx: month index; hist_dry: EIA dry production (NaN after history); mkt_base / mkt_delayed: marketed forecast (NaN in
    history); loss_share: extraction loss as a share of marketed; fv: load_gas() frame. Returns the 'Supply and demand' frame."""
    idx = idx[idx >= START]
    o = pd.DataFrame(index=idx)
    o["Type"] = fv["Actual / forecast"].reindex(idx).fillna("Forecast")
    for c in SECT:
        o[c] = fv[c].reindex(idx)
    o["LNG feedgas (base)"] = fv[LNG_B].reindex(idx)
    o["Pipeline exports to Mexico"] = fv[MEX].reindex(idx)
    o["Demand incl. LNG, base"] = o[SECT + ["LNG feedgas (base)", "Pipeline exports to Mexico"]].sum(axis=1, min_count=7)
    o["LNG feedgas (delayed)"] = fv[LNG_D].reindex(idx)
    o["Demand incl. LNG, LNG delayed"] = o["Demand incl. LNG, base"] - o["LNG feedgas (base)"] + o["LNG feedgas (delayed)"]
    k = 1 - loss_share
    hist = hist_dry.reindex(idx)
    o["Dry production, base"] = hist.fillna(mkt_base.reindex(idx) * k)
    o["Dry production, takeaway delayed"] = hist.fillna(mkt_delayed.reindex(idx) * k)
    o["Implied net outflow, base"] = o["Dry production, base"] - o["Demand incl. LNG, base"]
    o["Implied net outflow, delayed (production and LNG both delayed)"] = (
        o["Dry production, takeaway delayed"] - o["Demand incl. LNG, LNG delayed"])
    o.index.name = "Month"
    return o


def checks(o, last_hist):
    """Prints the join and plausibility checks; returns them as text lines (for the Units tab)."""
    b, d = "Implied net outflow, base", "Implied net outflow, delayed (production and LNG both delayed)"
    a, f = o.loc[:last_hist], o.loc[last_hist + pd.offsets.MonthBegin(1):]
    t12 = a[b].iloc[-12:]
    jump = f[b].iloc[0] - a[b].iloc[-1]
    pj = f["Dry production, base"].iloc[0] - a["Dry production, base"].iloc[-1]
    ann = o[[b, d]].groupby(o.index.year).mean()
    lines = [f"last 12 history months implied outflow: mean {t12.mean():.1f}, min {t12.min():.1f}, max {t12.max():.1f} Bcf/d",
             f"join {last_hist:%b/%y} -> {f.index[0]:%b/%y}: outflow {a[b].iloc[-1]:.1f} -> {f[b].iloc[0]:.1f} ({jump:+.1f}), dry production {pj:+.1f} Bcf/d",
             f"forecast outflow max base {f[b].max():.1f} ({f[b].idxmax():%b/%y}), delayed {f[d].max():.1f}; annual means " +
             "; ".join(f"{y}: {r[b]:.1f}/{r[d]:.1f}" for y, r in ann.iterrows() if y >= 2025)]
    for l in lines:
        print("SD check:", l)
    return lines
