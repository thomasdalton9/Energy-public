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
  STORAGE           Texas gross injections (DEMAND side) and withdrawals (SUPPLY side), Bcf/d, EIA Natural Gas Monthly (stor/sum, area STX;
                    EIA-191). History = EIA. Forecast = the calendar-month mean of the last 36 EIA months (same 3-year seasonal idea as the
                    LNG utilisation, not smoothed), with the net annual storage change held at the last-3-year mean unless the editable
                    'Net storage change override' on the Assumptions tab is filled (ASSUMPTION; no storage capacity additions assumed).
  IMPLIED OUTFLOW   dry production + storage withdrawal - (demand stack + storage injection) = net gas leaving Texas by pipeline to
                    other states, plus lease/plant and pipeline fuel EIA withholds. The earlier definition (dry production - demand,
                    no storage) is kept as the memo columns 'before storage' (it carried the storage seasonality: higher in summer,
                    lower in winter). Base = base production with base LNG; delayed = takeaway-delayed production with LNG-delayed
                    feedgas (both delays are the 6-month cases of their workbooks). Sectors and Mexico are the same in both.
"""
import pandas as pd

SECT = ["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel"]
SHEET = "Supply and demand"
LNG_B, LNG_D = "LNG feedgas total - base", "LNG feedgas total - delayed"
MEX = "Pipeline exports to Mexico"
START = pd.Timestamp("2021-01-01")
STOR_OVERRIDE = "Net storage change override (Bcf/d average, + = net injection; blank = last-3-year mean)"


def load_gas(texas_xlsx):
    fv = pd.read_excel(texas_xlsx, sheet_name="Forecast values")
    fv["Month"] = pd.to_datetime(fv["Month"])
    return fv.set_index("Month").sort_index()


def storage_hist(raw):
    """EIA stor/sum store (monthly MMcf, columns 'stor|...') -> Bcf/d injections 'inj', withdrawals 'wd' and month-end working gas
    'stock' (Bcf). Rows without both flows are dropped from the flow columns."""
    c = lambda n: raw["stor|" + n] if "stor|" + n in raw else pd.Series(index=raw.index, dtype=float)   # noqa: E731
    days = raw.index.days_in_month
    h = pd.DataFrame({"inj": c("Injections") / days / 1000.0, "wd": c("Withdrawals") / days / 1000.0, "stock": c("Working gas") / 1000.0})
    h.index = pd.to_datetime(h.index)
    return h.sort_index()


def storage_path(hist, idx, override=None, years=3):
    """Gross injection / withdrawal path over idx. History = EIA. Later months = calendar-month mean of the last `years` x 12 EIA months
    (injection shifted by a constant so the annual mean net injection equals `override` when given, else the window mean). Returns
    (frame inj, wd, net injection, basis, stock; dict of facts)."""
    h = hist.dropna(subset=["inj", "wd"])
    last = h.index.max()
    w = h.iloc[-12 * years:]
    pi, pw = w.groupby(w.index.month)["inj"].mean(), w.groupby(w.index.month)["wd"].mean()
    net_def = float((w["inj"] - w["wd"]).mean())
    ov = None if override is None or pd.isna(override) else float(override)
    target = net_def if ov is None else ov
    shift = target - float((pi - pw).mean())
    o = pd.DataFrame(index=idx)
    o["inj"], o["wd"] = h["inj"].reindex(idx), h["wd"].reindex(idx)
    fut = idx[idx > last]
    o.loc[fut, "inj"] = [pi[m.month] + shift for m in fut]
    o.loc[fut, "wd"] = [pw[m.month] for m in fut]
    o["net_inj"] = o["inj"] - o["wd"]
    o["basis"] = ["EIA" if m <= last else f"seasonal pattern of {w.index[0]:%b/%y}-{w.index[-1]:%b/%y}" for m in idx]
    stock = hist["stock"].reindex(idx)
    s0 = hist["stock"].dropna()
    s_last = s0.iloc[-1] if len(s0) else float("nan")
    d_last = s0.index.max() if len(s0) else None
    if d_last is not None:
        run = s_last
        for m in idx[idx > d_last]:
            run = run + o.at[m, "net_inj"] * m.days_in_month
            stock[m] = run
    o["stock"] = stock
    o["stock_basis"] = ["EIA" if (d_last is not None and m <= d_last) else "projected from net injection" for m in idx]
    return o, dict(last=last, net_def=net_def, target=target, override=ov, window=f"{w.index[0]:%b/%y}-{w.index[-1]:%b/%y}",
                   inj12=float(h["inj"].iloc[-12:].mean()), wd12=float(h["wd"].iloc[-12:].mean()))


def build(idx, hist_dry, mkt_base, mkt_delayed, loss_share, fv, stor=None):
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
    o["Implied net outflow before storage, base"] = o["Dry production, base"] - o["Demand incl. LNG, base"]
    o["Implied net outflow before storage, delayed"] = o["Dry production, takeaway delayed"] - o["Demand incl. LNG, LNG delayed"]
    if stor is not None:
        o["Storage injection (demand)"] = stor["inj"].reindex(idx)
        o["Storage withdrawal (supply)"] = stor["wd"].reindex(idx)
        o["Storage basis"] = stor["basis"].reindex(idx)
        o["Dry production + storage withdrawal, base"] = o["Dry production, base"] + o["Storage withdrawal (supply)"]
        o["Dry production + storage withdrawal, takeaway delayed"] = o["Dry production, takeaway delayed"] + o["Storage withdrawal (supply)"]
        o["Implied net outflow, base"] = o["Dry production + storage withdrawal, base"] - o["Demand incl. LNG, base"] - o["Storage injection (demand)"]
        o["Implied net outflow, delayed (production and LNG both delayed)"] = (
            o["Dry production + storage withdrawal, takeaway delayed"] - o["Demand incl. LNG, LNG delayed"] - o["Storage injection (demand)"])
    else:
        o["Implied net outflow, base"] = o["Implied net outflow before storage, base"]
        o["Implied net outflow, delayed (production and LNG both delayed)"] = o["Implied net outflow before storage, delayed"]
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
    nb = "Implied net outflow before storage, base"
    sw = (f"; before storage (memo) mean {a[nb].iloc[-12:].mean():.1f}, min {a[nb].iloc[-12:].min():.1f}, max {a[nb].iloc[-12:].max():.1f}"
          if nb in a else "")
    lines = [f"last 12 history months implied outflow: mean {t12.mean():.1f}, min {t12.min():.1f}, max {t12.max():.1f} Bcf/d" + sw,
             f"join {last_hist:%b/%y} -> {f.index[0]:%b/%y}: outflow {a[b].iloc[-1]:.1f} -> {f[b].iloc[0]:.1f} ({jump:+.1f}), dry production {pj:+.1f} Bcf/d",
             f"forecast outflow max base {f[b].max():.1f} ({f[b].idxmax():%b/%y}), delayed {f[d].max():.1f}; annual means " +
             "; ".join(f"{y}: {r[b]:.1f}/{r[d]:.1f}" for y, r in ann.iterrows() if y >= 2025)]
    for l in lines:
        print("SD check:", l)
    return lines
