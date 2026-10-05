"""
Texas gas demand to Dec 2033 (Bcf/d): LNG feedgas build 2026-28 ('overhang') followed by growing data-centre power demand.
Called by TEXAS_PRODUCTION_FORECAST.py (nothing runs on its own); adds three tabs to texas_production_forecast.xlsx:

  'Assump - Data centres'  yellow input cells (editable, read back from the committed workbook on every run): IEA anchor, Texas
                           shares, load factor, gas share of marginal supply, heat rate, unit constants, post-2028 extension rules,
                           ERCOT cross-check inputs. Every number from memory is labelled 'unverified'.
  'Data centres'           year-end energised Texas data-centre load (GW) for LOW / BASE / HIGH, 2024-2033, and its gas burn in
                           Bcf/d - LIVE Excel formulas on the assumption cells.
  'Demand to 2033'         the script's own evaluation, month by month to Dec 2033 (used for the charts, PNGs and dashboard):
                           stacked demand (EIA sectors + LNG feedgas + Mexico + data centres), LOW/BASE/HIGH totals, dry
                           production and the implied net outflow to other states.

LOAD   IEA 'Energy and AI' (Apr 2025; figures unverified, from memory): US data-centre electricity ~180 TWh in 2024, +~240 TWh by
       2030 (base). IEA gives scenario endpoints at 2030 and 2035 only: 2025-29 and 2031-34 are LINEARLY INTERPOLATED (labelled).
       Texas = a share of the 2024 US stock + an editable Texas share of US GROWTH. Energised GW = TWh / 8.76 / load factor.
       Cross-check: ERCOT's large-load interconnection queue / long-term load forecast (unverified; see the Assump tab).
GAS    Bcf/d = GW x 1000 MW x load factor x gas share of marginal supply x heat rate (MMBtu/MWh) x 24 h / (MMBtu per Mcf) / 1e6.
       Heat rate = mean of the last 12 calibrated months of ercot_gas_burn_daily.xlsx. Roughly 0.08-0.10 Bcf/d per GW.
       Only the burn ADDED after the last EIA actual month is stacked (EIA's electric-power history already contains today's
       data-centre load); the annual tab shows the total.
2033   LNG: the train table of texas_gas_monthly.xlsx ('Assump - LNG'; last train 2031) held at its steady utilisation by
       TEXAS_GAS_FORECAST's own logic run to Dec 2033. Other sectors: the same seasonal-trend base; electric power is held flat
       after 2028 (editable) because data centres are the explicit power-growth layer. Production: STEO to Dec 2027, 2028 as in
       the base workbook, then each region's growth DAMPED (editable factor per year, compounded month by month from the Dec 2028 level, no seasonal shape) - 'extension, not STEO' - capped by the
       takeaway table (capacity held after the last listed pipeline). Everything after Dec 2028 is a SCENARIO.
"""
import math
import os

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

import TEXAS_GAS_FORECAST as gf

END = pd.Timestamp("2033-12-01")
CHART_START = pd.Timestamp("2024-01-01")
AS = "Assump - Data centres"
DC = "Data centres"
VIEW = "Demand to 2033"
CASES = ("LOW", "BASE", "HIGH")
YEARS = list(range(2024, 2034))
FILL_IN = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
FILL_HEAD = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
FILL_LIGHT = PatternFill("solid", start_color="EDEDED", end_color="EDEDED")
BOLD = Font(bold=True)
IEA = "IEA, Energy and AI (Apr 2025)"
UNV = "unverified (from memory)"

# (key, label, low, base, high, unit, source / status); rows start at ROW0
ROW0 = 5
PARAMS = [
    ("us24", "US data-centre electricity, 2024", 180, 180, 180, "TWh", f"{IEA}: ~180 TWh in 2024 - {UNV}"),
    ("inc30", "US increase 2024 -> 2030", 120, 240, 360, "TWh", f"BASE = IEA's ~+240 TWh by 2030 ({UNV}); LOW / HIGH = own sensitivities (about half / 1.5x)"),
    ("inc35", "US increase 2024 -> 2035", 220, 380, 620, "TWh", f"IEA scenario endpoint 2035 - values {UNV}; own cases around it. 2031-34 are interpolated between the 2030 and 2035 endpoints"),
    ("sh24", "Texas share of the US 2024 data-centre stock", 0.10, 0.10, 0.10, "share", "own assumption (unverified)"),
    ("shg", "Texas share of US data-centre GROWTH", 0.15, 0.25, 0.40, "share", "own assumption; ERCOT's large-load queue is the cross-check (rows below), unverified"),
    ("lf", "Load factor (average load / energised capacity)", 0.85, 0.85, 0.85, "ratio", "own assumption"),
    ("gas", "Gas share of marginal supply (share of the extra load met by gas burn)", 0.40, 0.50, 0.60, "share", "own assumption"),
    ("hr", "Effective heat rate", None, None, None, "MMBtu/MWh", "mean of the last 12 calibrated months of ercot_gas_burn_daily.xlsx (set by the script; overwrite to override)"),
    ("mcf", "MMBtu per Mcf", 1.036, 1.036, 1.036, "MMBtu/Mcf", "pipeline-quality gas, as specified (ercot_gas_burn_daily.xlsx uses ~1.015)"),
]
R = {k: ROW0 + i for i, (k, *_r) in enumerate(PARAMS)}
R_PERGW = ROW0 + len(PARAMS)                    # Bcf/d per GW (formula)
SCAL0 = R_PERGW + 3                             # scalar block
SCALARS = [
    ("gelec", "Electric-power sector underlying growth after Dec 2028", 0.0, "per year",
     "0 = flat: data centres are the explicit power-demand growth layer, so EIA's trend growth is not extended (avoids double counting)"),
    ("damp", "Production extension after 2028: growth multiplier per year", 0.5, "x",
     "STEO regional growth of 2027/2026 is damped by this factor each year after 2028 (0.5 = halves every year). 'Extension, not STEO'. Own assumption"),
    ("queue", "ERCOT large-load interconnection queue (all large loads)", 226, "GW", f"ERCOT large-load queue, late 2025 - {UNV}; not fetched (probe: discovery_archive/ERCOT_IEA_DATACENTRE_PROBE.py - ERCOT pages reachable but no figures extracted, IEA 403)"),
    ("qdc", "Data-centre share of the queue", 0.70, "share", f"{UNV}"),
]
RS = {k: SCAL0 + i for i, (k, *_r) in enumerate(SCALARS)}
YR0 = 7                                          # first data row (2024) of the 'Data centres' tab


def ercot_heat_rate(path):
    """Mean heat rate (MMBtu/MWh) of the last 12 calibrated complete months of ercot_gas_burn_daily.xlsx; (value, text)."""
    try:
        m = pd.read_excel(path, sheet_name="Monthly")
        ok = m[(m["ERCOT_days"] >= 28) & ~m["Heat_rate_basis"].astype(str).str.startswith("estimated")].tail(12)
        return float(ok["Heat_rate_used_MMBtu_per_MWh"].mean()), (
            f"{ok['month'].iloc[0]:%b/%y}-{ok['month'].iloc[-1]:%b/%y} ({len(ok)} months)")
    except Exception as e:  # noqa: BLE001
        print(f"  ERCOT heat rate unreadable ({type(e).__name__}: {e}) - 8.5 MMBtu/MWh used", flush=True)
        return 8.5, "fallback 8.5 (ercot_gas_burn_daily.xlsx unreadable)"


def read_prior(path):
    """Editable cells of the committed workbook (None when absent). Called BEFORE the workbook is rewritten."""
    if not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, data_only=False)
        if AS not in wb.sheetnames:
            return None
        a = wb[AS]
        p = {k: [a.cell(R[k], c).value for c in (2, 3, 4)] for k, *_x in PARAMS}
        s = {k: a.cell(RS[k], 2).value for k, *_x in SCALARS}
        return {"p": p, "s": s}
    except Exception as e:  # noqa: BLE001
        print(f"  prior data-centre assumptions unreadable ({type(e).__name__}: {e}) - defaults used", flush=True)
        return None


def assumptions(prior, hr):
    par = {k: [lo, ba, hi] for k, _l, lo, ba, hi, _u, _n in PARAMS}
    par["hr"] = [hr, hr, hr]
    sc = {k: v for k, _l, v, _u, _n in SCALARS}
    if prior:
        for k, v in prior["p"].items():
            if all(isinstance(x, (int, float)) for x in v):
                par[k] = [float(x) for x in v]
        for k, v in prior["s"].items():
            if isinstance(v, (int, float)):
                sc[k] = float(v)
    return par, sc


# --------------------------------------------------------------------------- python model of the Excel formulas
def annual(par, ci):
    """Year-end energised load and gas burn for one case: {year: dict(us, tx_twh, gw, bcfd)}."""
    g = lambda k: par[k][ci]   # noqa: E731
    per_gw = 1000 * g("lf") * g("gas") * g("hr") * 24 / g("mcf") / 1e6
    out = {}
    for y in YEARS:
        us = g("us24") + (g("inc30") * (y - 2024) / 6 if y <= 2030 else g("inc30") + (g("inc35") - g("inc30")) * (y - 2030) / 5)
        tx = g("us24") * g("sh24") + g("shg") * (us - g("us24"))
        gw = tx / 8.76 / g("lf")
        out[y] = {"us": us, "tx_twh": tx, "gw": gw, "bcfd": gw * per_gw}
    return out, per_gw


def monthly_path(ann, key):
    """Year-end annual values -> monthly path (linear within the year; month m of year y = year-end y-1 + m/12 of the step)."""
    out = {}
    for y in YEARS[1:]:
        for m in range(1, 13):
            out[pd.Timestamp(year=y, month=m, day=1)] = ann[y - 1][key] + (ann[y][key] - ann[y - 1][key]) * m / 12
    return pd.Series(out)


# --------------------------------------------------------------------------- production extension
def takeaway_caps(tk, local, delay, idx):
    add = tk[~tk["Pipeline"].str.contains("Existing", case=False) & tk["Status"].astype(str).str.lower().ne("in service")]
    base_ex = float(tk[tk["Pipeline"].str.contains("Existing", case=False)]["Capacity Bcf/d"].sum())
    caps = {}
    for nm, d in (("base", 0), ("delayed", delay)):
        c = pd.Series(base_ex + local, index=idx, dtype=float)
        for _, r in add.iterrows():
            c[c.index >= pd.Timestamp(r["In-service month"] + "-01") + pd.DateOffset(months=d)] += float(r["Capacity Bcf/d"])
        caps[nm] = c
    return caps


def extend_production(df, tk, shares, other, local, delay, damp, loss_share):
    """Marketed and dry Texas production to Dec 2033: months to Dec 2028 as the base workbook, then damped STEO-region growth
    capped by the takeaway table (capacity held after the last pipeline)."""
    idx = pd.date_range(df.index.min(), END, freq="MS")
    sp, se, sh = shares
    reg = {"Permian": "STEO Permian marketed (Bcf/d)", "Eagle Ford": "STEO Eagle Ford marketed (Bcf/d)",
           "Haynesville": "STEO Haynesville marketed (Bcf/d)"}
    s = {k: df[c].reindex(idx) for k, c in reg.items()}
    for k, ser in s.items():
        g = ser["2027-01-01":"2027-12-01"].mean() / ser["2026-01-01":"2026-12-01"].mean() - 1
        for m in idx[idx > pd.Timestamp("2028-12-01")]:
            # month-on-month from the Dec 2028 level (no seasonal shape): a same-month-last-year rule would step down at the join
            ser[m] = ser[m - pd.DateOffset(months=1)] * (1 + g * damp ** (m.year - 2028)) ** (1 / 12)
    caps = takeaway_caps(tk, local, delay, idx)
    out = {}
    for sc in ("base", "delayed"):
        perm = pd.concat([s["Permian"], caps[sc]], axis=1).min(axis=1)
        tx = sp * perm + se * s["Eagle Ford"] + sh * s["Haynesville"] + other
        tx = tx.where(idx > pd.Timestamp("2028-12-01"), df[f"Texas production forecast, {sc}"].reindex(idx))
        out[sc] = tx
    return idx, out, caps, s


# --------------------------------------------------------------------------- build
def build(prod_df, sdf, tk, params, texas_xlsx, ercot_xlsx, prior, loss_share, other, delay, last_hist):
    """prod_df: 'Forecast' frame of TEXAS_PRODUCTION_FORECAST (Month index); sdf: 'Supply and demand' frame. Returns (view frame,
    context for the Excel tabs and notes)."""
    hr, hr_txt = ercot_heat_rate(ercot_xlsx)
    par, sc = assumptions(prior, hr)
    ann, per_gw = {}, {}
    for ci, c in enumerate(CASES):
        ann[c], per_gw[c] = annual(par, ci)
    # --- gas demand to 2033 from the existing LNG / sector model run to Dec 2033
    cons = pd.read_excel(texas_xlsx, sheet_name="Consumption by sector", index_col=0, parse_dates=True)
    exports = pd.read_excel(texas_xlsx, sheet_name="Exports", index_col=0, parse_dates=True)
    bal = pd.read_excel(texas_xlsx, sheet_name="Balance", index_col=0, parse_dates=True)
    old_last = gf.LAST
    gf.LAST = END
    try:
        f, ctx = gf.build(texas_xlsx, cons, exports, bal)
    finally:
        gf.LAST = old_last
    fv = gf.values_sheet(f)
    N = gf.NAME
    idx = pd.date_range(CHART_START, END, freq="MS")
    v = pd.DataFrame(index=idx)
    first_fc = ctx["fs"]
    last_act = first_fc - pd.offsets.MonthBegin(1)
    v["Type"] = ["Actual" if m < first_fc else ("Forecast" if m <= pd.Timestamp("2028-12-01") else "Scenario") for m in idx]
    for c in gf.SECT:
        v[c] = fv[c].reindex(idx).astype(float)
    ge = sc["gelec"]
    j = idx > pd.Timestamp("2028-12-01")
    v.loc[j, "Electric power"] = [v.at[m - pd.DateOffset(years=m.year - 2028), "Electric power"] * (1 + ge) ** (m.year - 2028) for m in idx[j]]
    v["LNG feedgas (base case)"] = fv[N["lng_b"]].reindex(idx).astype(float)
    v["Pipeline exports to Mexico"] = fv[N["mex"]].reindex(idx).astype(float)
    # join check against the shipped 2028 workbook (must be identical before 2029)
    old = sdf.reindex(idx)
    k = idx <= pd.Timestamp("2028-12-01")
    diff = float(np.nanmax(np.abs(v.loc[k, "LNG feedgas (base case)"] - old.loc[k, "LNG feedgas (base)"])))
    diff2 = float(np.nanmax(np.abs(v.loc[k, gf.SECT].sum(axis=1) - old.loc[k, gf.SECT].sum(axis=1))))
    # --- data centres: added burn since the last EIA actual month
    for ci, c in enumerate(CASES):
        g_ = monthly_path(ann[c], "bcfd")
        gw_ = monthly_path(ann[c], "gw")
        base_m = g_[last_act]
        v[f"Data centres GW (year-end path), {c}"] = gw_.reindex(idx)
        v[f"Data centres gas burn, total, {c}"] = g_.reindex(idx)
        v[f"Data centres added gas burn, {c}"] = (g_.reindex(idx) - base_m).where(idx > last_act, 0.0)
    stack = v[gf.SECT + ["LNG feedgas (base case)", "Pipeline exports to Mexico"]].sum(axis=1, min_count=7)
    for c in CASES:
        v[f"Total demand incl. data centres, {c}"] = stack + v[f"Data centres added gas burn, {c}"]
    # --- production
    shares = (params["Permian Texas share of STEO Permian"], params["Eagle Ford Texas share"],
              params["Haynesville Texas share of STEO Haynesville"])
    pidx, mkt, caps, steo = extend_production(prod_df, tk, shares, other, params["Local Permian demand (power, industrial, Mexico-bound not included)"],
                                              delay, sc["damp"], loss_share)
    k_ = 1 - loss_share
    hist_dry = sdf["Dry production, base"].where(sdf["Type"].eq("Actual")).reindex(idx)
    v["Dry production, base"] = hist_dry.fillna(mkt["base"].reindex(idx) * k_)
    v["Dry production, takeaway delayed"] = hist_dry.fillna(mkt["delayed"].reindex(idx) * k_)
    for c in CASES:
        v[f"Implied net outflow, {c}"] = v["Dry production, base"] - v[f"Total demand incl. data centres, {c}"]
    v["Permian takeaway + local demand, base (Bcf/d)"] = caps["base"].reindex(idx)
    v["Permian STEO / extension marketed (Bcf/d)"] = steo["Permian"].reindex(idx)
    v.index.name = "Month"
    ctx2 = {"par": par, "sc": sc, "ann": ann, "per_gw": per_gw, "hr": hr, "hr_txt": hr_txt, "first_fc": first_fc, "last_act": last_act,
            "join_lng": diff, "join_sec": diff2, "loss_share": loss_share}
    return v, ctx2


def checks(v, c2):
    """Sanity flags; returns text lines (printed and put on the Units tab)."""
    out = []
    jan = pd.Timestamp("2029-01-01")
    dec = pd.Timestamp("2028-12-01")
    out.append(f"model run to 2033 reproduces the 2028 workbook: max |LNG diff| {c2['join_lng']:.1e}, sectors {c2['join_sec']:.1e} Bcf/d")
    for c in CASES:
        g = v[f"Data centres gas burn, total, {c}"]
        out.append(f"data-centre gas burn {c}: Dec/28 {g[dec]:.2f}, Dec/30 {g[pd.Timestamp('2030-12-01')]:.2f}, Dec/33 {g[END]:.2f} Bcf/d"
                   + ("  ** EXCEEDS 10 Bcf/d **" if g.max() > 10 else " (below the 10 Bcf/d flag)"))
    for c in CASES:
        o = v[f"Implied net outflow, {c}"]
        o_all, o = o, o[o.index >= c2["first_fc"]]
        neg = o[o < 0]
        out.append(f"implied outflow {c}: Dec/28 {o[dec]:.1f}, Dec/30 {o[pd.Timestamp('2030-12-01')]:.1f}, Dec/33 {o[END]:.1f}, min {o.min():.1f} ({o.idxmin():%b/%y})"
                   + (f"  ** NEGATIVE from {neg.index[0]:%b/%y} ({len(neg)} months): production cannot meet Texas demand, Texas would have to draw gas in from other states **" if len(neg) else ""))
    tot = v["Total demand incl. data centres, BASE"]
    typical = tot.diff().abs()[(tot.index > pd.Timestamp("2027-01-01")) & (tot.index <= dec)].median()
    jt, jp = tot[jan] - tot[dec], v["Dry production, base"][jan] - v["Dry production, base"][dec]
    jo = v["Implied net outflow, BASE"][jan] - v["Implied net outflow, BASE"][dec]
    ob = v["Implied net outflow, BASE"]
    jo_prev = ob[pd.Timestamp("2028-01-01")] - ob[pd.Timestamp("2027-12-01")]      # same seasonal step a year earlier
    out.append(f"join Dec/28 -> Jan/29 (BASE): total demand {jt:+.2f} (median monthly step 2027-28 {typical:.2f}), dry production {jp:+.2f}, outflow {jo:+.2f} Bcf/d "
               f"(the same Dec -> Jan step a year earlier was {jo_prev:+.2f}: winter demand) " + ("** JUMP **" if abs(jo - jo_prev) > 1.0 else "- no jump beyond the seasonal step"))
    return out


# --------------------------------------------------------------------------- Excel
def _hdr(ws, row, labels, c0=1):
    for j, t in enumerate(labels, start=c0):
        c = ws.cell(row, j, t)
        c.font = BOLD
        c.fill = FILL_HEAD
        c.alignment = Alignment(wrap_text=True, vertical="top")


def write_sheets(path, v, c2, check_lines):
    par, sc = c2["par"], c2["sc"]
    wb = load_workbook(path)
    for n in (AS, DC, VIEW):
        if n in wb.sheetnames:
            del wb[n]
    a = wb.create_sheet(AS)
    a["A1"] = "Texas data-centre assumptions (yellow cells are editable; 'Data centres' recalculates in Excel, 'Demand to 2033' and the charts follow on the next run)"
    a["A1"].font = Font(bold=True, size=13)
    a["A2"] = (f"Anchored to {IEA}: the report gives scenario endpoints at 2030 and 2035, so 2025-29 and 2031-34 are LINEAR INTERPOLATIONS. "
               f"Every figure marked '{UNV}' was not fetched (one Actions probe, 5 Oct 2026: ERCOT's load-forecast and large-load pages answered 200 but yielded no readable queue figures, the IEA pages returned 403) - verify before relying on it.")
    a["A2"].alignment = Alignment(wrap_text=True)
    _hdr(a, 4, ["Assumption", "LOW", "BASE", "HIGH", "Unit", "Source / status"])
    for i, (k, lab, *_r, unit, note) in enumerate(PARAMS):
        r = ROW0 + i
        a.cell(r, 1, lab)
        for j in range(3):
            c = a.cell(r, 2 + j, par[k][j])
            c.fill = FILL_IN
            c.number_format = "0.00" if k in ("sh24", "shg", "lf", "gas") else ("0.000" if k == "mcf" else "0.0")
        a.cell(r, 5, unit)
        a.cell(r, 6, note + (f"; last 12 months {c2['hr_txt']}" if k == "hr" else ""))
    a.cell(R_PERGW, 1, "Gas burn per GW of energised load (formula)").font = BOLD
    for j, col in enumerate("BCD"):
        c = a.cell(R_PERGW, 2 + j, f"=1000*{col}{R['lf']}*{col}{R['gas']}*{col}{R['hr']}*24/{col}{R['mcf']}/1000000")
        c.number_format = "0.0000"
    a.cell(R_PERGW, 5, "Bcf/d per GW")
    a.cell(R_PERGW, 6, "= 1000 MW/GW x load factor x gas share x heat rate x 24 h / (MMBtu per Mcf) / 1e6")
    a.cell(SCAL0 - 1, 1, "Other inputs").font = BOLD
    for i, (k, lab, _v, unit, note) in enumerate(SCALARS):
        r = SCAL0 + i
        a.cell(r, 1, lab)
        c = a.cell(r, 2, sc[k])
        c.fill = FILL_IN
        c.number_format = "0.00" if k in ("gelec", "damp", "qdc") else "0"
        a.cell(r, 5, unit)
        a.cell(r, 6, note)
    r = SCAL0 + len(SCALARS) + 1
    a.cell(r, 1, "ERCOT cross-check: BASE Texas energised GW in 2030 / (queue GW x data-centre share)").font = BOLD
    a.cell(r, 2, f"='{DC}'!J{YR0 + 6}/(B{RS['queue']}*B{RS['qdc']})").number_format = "0.0%"
    a.cell(r, 6, "share of the queued data-centre load that is energised by 2030 in BASE; ERCOT queues are speculative, so a low double-digit share is plausible, "
                 "> 40% would look aggressive. Queue size unverified")
    for col, w in zip("ABCDEF", (66, 11, 11, 11, 12, 130)):
        a.column_dimensions[col].width = w
    a.row_dimensions[2].height = 45
    a.merge_cells("A2:F2")
    # ---- Data centres (live formulas)
    d = wb.create_sheet(DC)
    d["A1"] = "Texas data-centre load (year-end energised GW) and gas burn, LOW / BASE / HIGH - live formulas on 'Assump - Data centres'"
    d["A1"].font = Font(bold=True, size=13)
    d["A2"] = ("Texas TWh = US 2024 TWh x Texas stock share + Texas share of growth x (US TWh - US 2024 TWh); energised GW = TWh / 8.76 / load factor; "
               "gas Bcf/d = GW x gas burn per GW (assumption tab). 2025-29 and 2031-33 are interpolated between IEA scenario endpoints.")
    d["A3"] = ("EIA's electric-power history already contains the 2024-26 data-centre burn: the charts stack only the burn ADDED after the last EIA actual month "
               "(see 'Demand to 2033'); this tab shows the total.")
    for rr in (2, 3):
        d[f"A{rr}"].alignment = Alignment(wrap_text=True)
        d.merge_cells(f"A{rr}:Q{rr}")
        d.row_dimensions[rr].height = 32
    labels = ["Year (year-end)", "Basis"]
    for c in CASES:
        labels += [f"{c}: US TWh", f"{c}: Texas TWh", f"{c}: Texas energised GW", f"{c}: gas burn Bcf/d", f"{c}: added since Dec 2025, Bcf/d"]
    _hdr(d, YR0 - 1, labels)
    d.row_dimensions[YR0 - 1].height = 48
    for i, y in enumerate(YEARS):
        r = YR0 + i
        d.cell(r, 1, y)
        d.cell(r, 2, "IEA anchor (2024, unverified)" if y == 2024 else ("IEA scenario endpoint (2030, unverified)" if y == 2030 else
                                                                       ("IEA scenario endpoint (2035)" if y == 2035 else
                                                                        ("interpolated 2024-30" if y < 2030 else "interpolated 2030-35 (not an IEA data point)"))))
        if y > 2030:
            for cc in range(1, 18):
                d.cell(r, cc).fill = FILL_LIGHT
        for ci in range(3):
            col = "BCD"[ci]
            A = lambda k: f"'{AS}'!${col}${R[k]}"   # noqa: E731
            c0 = 3 + 5 * ci
            us, tx, gw, bc, inc = (L(c0 + n) for n in range(5))
            d[f"{us}{r}"] = (f"=IF($A{r}<=2030,{A('us24')}+{A('inc30')}*($A{r}-2024)/6,"
                             f"{A('us24')}+{A('inc30')}+({A('inc35')}-{A('inc30')})*($A{r}-2030)/5)")
            d[f"{tx}{r}"] = f"={A('us24')}*{A('sh24')}+{A('shg')}*({us}{r}-{A('us24')})"
            d[f"{gw}{r}"] = f"={tx}{r}/8.76/{A('lf')}"
            d[f"{bc}{r}"] = f"={gw}{r}*'{AS}'!${col}${R_PERGW}"
            d[f"{inc}{r}"] = f"={bc}{r}-{bc}${YR0 + 1}"
            for n, fmt in zip((us, tx, gw, bc, inc), ("0.0", "0.0", "0.00", "0.00", "0.00")):
                d[f"{n}{r}"].number_format = fmt
    d.column_dimensions["A"].width = 14
    d.column_dimensions["B"].width = 38
    for cc in range(3, 18):
        d.column_dimensions[L(cc)].width = 13
    d.freeze_panes = f"C{YR0}"
    # ---- values view
    w = wb.create_sheet(VIEW)
    cols = ["Month"] + list(v.columns)
    _hdr(w, 1, cols)
    w.row_dimensions[1].height = 62
    for i, (m, row) in enumerate(v.iterrows(), start=2):
        w.cell(i, 1, m.to_pydatetime()).number_format = "mmm/yy"
        for j, x in enumerate(row.values, start=2):
            if isinstance(x, str):
                w.cell(i, j, x)
            elif x is not None and not (isinstance(x, float) and math.isnan(x)):
                w.cell(i, j, float(x)).number_format = "0.00"
            if row["Type"] == "Scenario":
                w.cell(i, j).fill = FILL_LIGHT
    for j in range(1, len(cols) + 1):
        w.column_dimensions[L(j)].width = 13
    w.freeze_panes = "C2"
    base = [s for s in wb._sheets if s.title not in (AS, DC, VIEW)]
    # keep the new tabs together just before 'Supply and demand' (or at the end)
    names = [s.title for s in base]
    pos = names.index("Supply and demand") + 1 if "Supply and demand" in names else len(base)
    wb._sheets = base[:pos] + [wb[AS], wb[DC], wb[VIEW]] + base[pos:]
    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def notes_lines(c2, check_lines):
    par = c2["par"]
    return [
        "", "DATA CENTRES AND THE 2033 VIEW (americas/TEXAS_DATACENTRE.py; Bcf/d)",
        f"Tabs 'Assump - Data centres' (yellow cells editable, survive reruns), 'Data centres' (live formulas: year-end energised Texas data-centre GW and gas burn, LOW/BASE/HIGH, 2024-33) and 'Demand to 2033' (script values, month by month to Dec 2033, charted).",
        f"Anchor: {IEA}, US data centres ~180 TWh in 2024 and +~240 TWh by 2030 (BASE) - {UNV}; IEA scenario endpoints are 2030 and 2035, so 2025-29 and 2031-34 are linearly interpolated. LOW/HIGH, the Texas shares (stock 10%, growth 15/25/40%), load factor, gas shares (40/50/60%) are own assumptions; ERCOT's large-load queue (226 GW, 70% data centres) is a {UNV} cross-check, not fetched.",
        f"Gas = GW x 1000 x load factor {par['lf'][1]:.2f} x gas share x heat rate {c2['hr']:.2f} MMBtu/MWh (mean of the last 12 calibrated ERCOT months, {c2['hr_txt']}) x 24 / {par['mcf'][1]:.3f} MMBtu per Mcf / 1e6 = {c2['per_gw']['BASE']:.3f} Bcf/d per GW in BASE.",
        "Only burn ADDED after the last EIA actual month is stacked on the sector demand (EIA's electric-power history already includes the present data-centre load); the electric-power sector is held flat after 2028 so the trend and the data-centre layer do not double count.",
        "2029-33 is a SCENARIO (lighter bars): LNG = the existing train table (last train 2031) at steady utilisation; other sectors = the seasonal-trend base extended; production = STEO to Dec 2027, 2028 as in the base forecast, then region growth damped by the editable factor (extension, not STEO) and capped by the takeaway table with capacity held after the last listed pipeline.",
        "Implied net outflow = dry production (base) - demand incl. LNG feedgas, Mexico and data centres, for LOW / BASE / HIGH data-centre cases; a negative value means Texas production cannot meet Texas demand and gas must flow in from other states - it is shown, never clipped.",
        "Sanity checks of the last run: " + "; ".join(check_lines),
    ]
