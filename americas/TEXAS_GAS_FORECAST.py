"""
Texas LNG demand forecast to Dec 2028 (Bcf/d), added to texas_gas_monthly.xlsx by TEXAS_GAS.py on every run (forecast tabs are
recomputed each run; the EIA history in 'Raw MMcf' stays the incremental store; the editable assumption tab is read back from
the committed workbook so the owner's edits survive). Production / takeaway / outflow forecasts live in the separate
TEXAS_PRODUCTION_FORECAST.py workbook.

  LNG demand      Corpus Christi + Freeport = nameplate feedgas x the monthly utilisation of the last 36 EIA months (summer
                  heat derate, maintenance and Freeport's trips are in that profile). Golden Pass, Corpus Christi trains 8-9,
                  Rio Grande and Port Arthur = one row per train in 'Assump - LNG' (nameplate, first-LNG month, ramp months,
                  ramp-start and steady utilisation). Feedgas = EIA port exports x 1.09 (liquefaction fuel and shrink; the same
                  ratio lng_feedgas_daily.xlsx uses to calibrate its meters). Base and 'delayed N months' scenarios (N in
                  'Assump - LNG'!B4; trains not yet started slip, running ones don't).
  Other demand    Seasonal-trend base: each sector's last-3-year monthly profile x last-12-month level x damped trend.

Excel: 'Forecast' (+ 'LNG trains calc') carries live formulas for LNG, driven by 'Assump - LNG'. Consumption and Mexico
forecasts and the seasonal profiles are computed by the script (override columns, or rerun). 'Forecast values' is the script's
own evaluation of the same logic (used for PNGs and the master dashboards, which cannot read formula results).

FERC's LNG pages list no dates (discovery_archive/americas/TEXAS_STEO_FERC_PROBE.py), so first-LNG months are company statements
(NextDecade, Sempra, ExxonMobil, Cheniere, EIA) or assumptions - see the 'Date status' column.
"""
import math
import os

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FIRST = pd.Timestamp("2021-01-01")
LAST = pd.Timestamp("2028-12-01")
HIST_WINDOW = 36               # months of EIA history behind the utilisation / storage profiles
FIRST_ROW = 6                  # first data row of the Forecast sheets (Jan 2021)
SECT = ["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel"]
EXISTING = ("Corpus Christi", "Freeport")
PLANTS = ["Corpus Christi", "Freeport", "Golden Pass", "Rio Grande", "Port Arthur"]
AL = "Assump - LNG"
TRAIN_ROWS = (12, 31)          # rows of the train table in AL
SEAS_ROWS = (38, 49)           # twelve months in AL

# nameplate feedgas of the operating plants, Bcf/d (lng_feedgas_daily.xlsx 'Train start-ups': FERC feed-gas letters)
CC_STEPS = [("2018-08-16", 0.717), ("2019-03-11", 0.717), ("2020-09-24", 0.717), ("2024-12-23", 0.205),
            ("2025-06-05", 0.205), ("2025-08-28", 0.205), ("2025-10-20", 0.205), ("2026-02-09", 0.205),
            ("2026-04-24", 0.205), ("2026-07-28", 0.205)]
FP_STEPS = [("2019-07-19", 0.717), ("2019-11-07", 0.717), ("2020-03-09", 0.717)]

# (plant, train, nameplate feedgas Bcf/d, first-LNG month, ramp months, ramp-start util, steady util, date status, note)
TRAINS = [
    ("Golden Pass", "Train 1", 0.865, "2025-11", 24, 0.10, 0.90, "actual (first LNG Nov 2025)",
     "Started Nov 2025 (lng_feedgas_daily Train start-ups); slow ramp: 0.12-0.24 Bcf/d of exports Apr-Jul 2026, 0.33 Bcf/d metered Sep 2026"),
    ("Golden Pass", "Train 2", 0.865, "2026-12", 12, 0.15, 0.90, "conflicting statements",
     "FERC approved Train 2 commissioning Jun 2026 (Argus); EIA: start-up 2H 2026; ExxonMobil/McDermott: first LNG 1H 2027. Dec 2026 chosen; the 6-month delay case = Jun 2027"),
    ("Golden Pass", "Train 3", 0.865, "2027-07", 12, 0.15, 0.90, "unverified",
     "Mechanical completion targeted 2Q 2027 (earlier repo research, not re-verified); EIA has it exporting in 2027"),
    ("Rio Grande", "Train 1", 0.841, "2027-04", 8, 0.15, 0.95, "company guidance (month assumed)",
     "NextDecade Q2 2026 update (30 Jul 2026): first gas later in 2026, first LNG from Train 1 in 1H 2027; Trains 1-2 74% complete, ahead of the guaranteed EPC schedule"),
    ("Rio Grande", "Train 2", 0.841, "2027-10", 8, 0.15, 0.95, "unverified",
     "Month estimated: EIA has Trains 1-2 exporting in 2027; NextDecade gives no month for Train 2"),
    ("Rio Grande", "Train 3", 0.841, "2028-07", 8, 0.15, 0.95, "unverified",
     "Month estimated ~9 months after Train 2; Train 3 more than 50% complete (Jun 2026)"),
    ("Rio Grande", "Train 4", 0.841, "2030-07", 8, 0.15, 0.95, "unverified",
     "FID Sep 2025, 15.5% complete (Jun 2026); beyond the forecast horizon"),
    ("Rio Grande", "Train 5", 0.841, "2031-01", 8, 0.15, 0.95, "unverified",
     "FID Oct 2025, 9.4% complete (Jun 2026); beyond the forecast horizon"),
    ("Port Arthur", "Train 1", 0.932, "2027-03", 8, 0.15, 0.95, "unverified (month assumed)",
     "Sempra: Train 1 commercial operations 2027; Argus (2026): commissioning steps under way, ramp-up could begin in Q4 2026. First LNG month assumed"),
    ("Port Arthur", "Train 2", 0.932, "2028-02", 8, 0.15, 0.95, "unverified (month assumed)",
     "Sempra: Train 2 commercial operations 2028"),
    ("Port Arthur", "Train 3", 0.968, "2030-09", 8, 0.15, 0.95, "unverified",
     "Phase 2 FID Sep 2025, completion 2030-31 (Sempra); beyond the horizon"),
    ("Corpus Christi", "Midscale Train 8", 0.215, "2028-06", 4, 0.30, 1.00, "unverified (month assumed)",
     "Cheniere: Midscale Trains 8-9 substantial completion 2H 2028 (earlier repo research); steady 1.00 as the plant runs at nameplate"),
    ("Corpus Christi", "Midscale Train 9", 0.215, "2028-09", 4, 0.30, 1.00, "unverified (month assumed)",
     "As Train 8. Not included: Corpus Christi Stage 4 (pre-FID; FERC draft EIS 18 Sep 2026), Texas LNG Brownsville (no FID verified), Freeport Train 4 (shelved)"),
]

KEYS = ["month", "type",
        "lng_cc_b", "lng_fp_b", "lng_gp_b", "lng_rg_b", "lng_pa_b", "lng_ot_b", "lng_b",
        "lng_cc_d", "lng_fp_d", "lng_gp_d", "lng_rg_d", "lng_pa_d", "lng_ot_d", "lng_d",
        "exp_b", "exp_d",
        "elec", "ind", "res", "com", "veh", "cons", "mex",
        "dem_b", "dem_d",
        "x_cc", "x_fp", "x_gp"]
COL = {k: get_column_letter(i + 1) for i, k in enumerate(KEYS)}
NAME = {
    "month": "Month", "type": "Actual / forecast",
    "lng_cc_b": "LNG feedgas Corpus Christi - base", "lng_fp_b": "LNG feedgas Freeport - base",
    "lng_gp_b": "LNG feedgas Golden Pass - base", "lng_rg_b": "LNG feedgas Rio Grande - base",
    "lng_pa_b": "LNG feedgas Port Arthur - base", "lng_ot_b": "LNG feedgas other new plants - base",
    "lng_b": "LNG feedgas total - base",
    "lng_cc_d": "LNG feedgas Corpus Christi - delayed", "lng_fp_d": "LNG feedgas Freeport - delayed",
    "lng_gp_d": "LNG feedgas Golden Pass - delayed", "lng_rg_d": "LNG feedgas Rio Grande - delayed",
    "lng_pa_d": "LNG feedgas Port Arthur - delayed", "lng_ot_d": "LNG feedgas other new plants - delayed",
    "lng_d": "LNG feedgas total - delayed",
    "exp_b": "LNG exports (plain, = feedgas / ratio) - base", "exp_d": "LNG exports (plain) - delayed",
    "elec": "Electric power", "ind": "Industrial", "res": "Residential", "com": "Commercial", "veh": "Vehicle fuel",
    "cons": "Consumption, published sectors", "mex": "Pipeline exports to Mexico",
    "dem_b": "Total demand incl. LNG feedgas - base", "dem_d": "Total demand incl. LNG feedgas - delayed",
    "x_cc": "EIA LNG exports Corpus Christi", "x_fp": "EIA LNG exports Freeport", "x_gp": "EIA LNG exports Golden Pass",
}
FILL_IN = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
FILL_HEAD = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
BOLD = Font(bold=True)


# --------------------------------------------------------------------------- helpers
def month_floor(x):
    x = pd.Timestamp(x)
    return pd.Timestamp(year=x.year, month=x.month, day=1)


def _date(x):
    if x is None or x == "" or (isinstance(x, float) and math.isnan(x)):
        return pd.NaT
    return month_floor(x)


def capacity_by_month(steps, months):
    """Feedgas nameplate in operation in each month (a train counts for the share of the month after its start-up)."""
    out = []
    for m in months:
        end = m + pd.offsets.MonthBegin(1)
        tot = 0.0
        for d, cap in steps:
            d = pd.Timestamp(d)
            if d < m:
                tot += cap
            elif d < end:
                tot += cap * (end - d).days / (end - m).days
        out.append(tot)
    return pd.Series(out, index=months)


def seasonal_trend(s, months, g_lo=-0.03, g_hi=0.06):
    """Seasonal-trend forecast of a monthly series: the last 12 months' level x the average calendar-month profile of the
    last 3 twelve-month blocks x a damped trend (last 12 vs the 12 before, clipped), compounded from the block's midpoint."""
    s = s.dropna()
    last = s.index.max()
    blocks = []
    for k in range(3):
        b = s[(s.index > last - pd.DateOffset(months=12 * (k + 1))) & (s.index <= last - pd.DateOffset(months=12 * k))]
        if len(b) == 12:
            blocks.append(b)
    level = blocks[0].mean()
    g = float(np.clip(level / blocks[1].mean() - 1, g_lo, g_hi)) if len(blocks) > 1 else 0.0
    prof = pd.concat([pd.Series((b / b.mean()).values, index=b.index.month) for b in blocks], axis=1).mean(axis=1)
    mid = last - pd.DateOffset(months=5)
    out = {}
    for m in months:
        yrs = ((m.year - mid.year) * 12 + m.month - mid.month) / 12.0
        out[m] = level * (1 + g) ** yrs * prof.get(m.month, 1.0)
    return pd.Series(out), g


def read_prior(path):
    """Editable cells of the committed workbook (trains, parameters, overrides) - None when absent."""
    if not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path)
        if AL not in wb.sheetnames:
            return None
        a = wb[AL]
        trains = {}
        for r in range(TRAIN_ROWS[0], TRAIN_ROWS[1] + 1):
            if a.cell(r, 1).value and a.cell(r, 2).value:
                trains[(a.cell(r, 1).value, a.cell(r, 2).value)] = [a.cell(r, c).value for c in range(3, 10)]
        over = {c: [a.cell(r, c).value for r in range(SEAS_ROWS[0], SEAS_ROWS[1] + 1)] for c in (3, 6, 9)}
        return {"trains": trains, "over": over, "delay": a["B4"].value, "ratio": a["B5"].value,
                "cap_cc": a["B35"].value, "cap_fp": a["C35"].value}
    except Exception as e:  # noqa: BLE001
        print(f"  prior assumptions unreadable ({type(e).__name__}: {e}) - defaults used", flush=True)
        return None


def build_assumptions(prior, par):
    """Defaults, then any value the owner edited in the committed workbook wins; owner-added rows are kept."""
    trains = {(p, t): [n, _date(d), r, rs, st, ds, note] for p, t, n, d, r, rs, st, ds, note in TRAINS}
    par = dict(par)
    over = {}
    if prior:
        for k, v in prior["trains"].items():
            old = trains.get(k)
            # the owner's numbers/dates win; status and note stay the script's unless the row is the owner's own
            trains[k] = [v[0], _date(v[1]), v[2], v[3], v[4], v[5] if old is None else v[5], v[6]]
        for k in ("delay", "ratio", "cap_cc", "cap_fp"):
            if prior.get(k) not in (None, ""):
                par[k] = prior[k]
        over = prior["over"]
    return trains, par, over


# --------------------------------------------------------------------------- python model
def run_model(hist, trains, par, prof, over, fs):
    """Forecast values frame (same logic as the Excel formulas)."""
    months = pd.date_range(FIRST, LAST, freq="MS")
    f = pd.DataFrame(index=months, columns=KEYS[1:], dtype=object)
    ratio, delay = float(par["ratio"]), int(par["delay"])
    act = months[months < fs]
    fc = months[months >= fs]

    def used(col_over, comp):
        return [o if o not in (None, "") else c for o, c in zip(col_over, comp)]
    seas = {"cc": used(over.get(3, [None] * 12), prof["cc"]), "fp": used(over.get(6, [None] * 12), prof["fp"]),
            "nb": used(over.get(9, [None] * 12), prof["nb"])}

    def train_flow(tr, m, dly):
        nameplate, start, ramp, rs, st = tr[0], tr[1], tr[2], tr[3], tr[4]
        if pd.isna(start):
            return 0.0
        if dly and start >= fs:
            start = start + pd.DateOffset(months=dly)
        if m < start:
            return 0.0
        n = (m.year - start.year) * 12 + m.month - start.month
        return nameplate * (rs + (st - rs) * min(1.0, n / max(ramp, 1))) * seas["nb"][m.month - 1]

    plant_key = {"Corpus Christi": "cc", "Freeport": "fp", "Golden Pass": "gp", "Rio Grande": "rg", "Port Arthur": "pa"}
    for dly, sfx in ((0, "_b"), (delay, "_d")):
        for m in months:
            if m < fs:
                h = lambda c: 0.0 if (m not in hist.index or pd.isna(hist.at[m, c])) else float(hist.at[m, c])   # noqa: E731
                vals = {"cc": h("x_cc") * ratio, "fp": h("x_fp") * ratio, "gp": h("x_gp") * ratio, "rg": 0.0, "pa": 0.0, "ot": 0.0}
            else:
                vals = {"cc": par["cap_cc"] * seas["cc"][m.month - 1], "fp": par["cap_fp"] * seas["fp"][m.month - 1],
                        "gp": 0.0, "rg": 0.0, "pa": 0.0, "ot": 0.0}
                for k, tr in trains.items():
                    vals[plant_key.get(k[0], "ot")] += train_flow(tr, m, dly)
            for k, v in vals.items():
                f.at[m, f"lng_{k}{sfx}"] = v
            f.at[m, f"lng{sfx}"] = sum(vals.values())
            f.at[m, f"exp{sfx}"] = f.at[m, f"lng{sfx}"] / ratio
    for k in ("x_cc", "x_fp", "x_gp"):
        f[k] = hist[k].reindex(months)
    for key, col in (("elec", "Electric power"), ("ind", "Industrial"), ("res", "Residential"), ("com", "Commercial"),
                     ("veh", "Vehicle fuel"), ("mex", "mex")):
        f[key] = pd.concat([hist[col].reindex(act), prof["fc_" + key].reindex(fc)])
    f["cons"] = f[["elec", "ind", "res", "com", "veh"]].astype(float).sum(axis=1, min_count=5)
    f["dem_b"] = f["cons"] + f["lng_b"].astype(float)
    f["dem_d"] = f["cons"] + f["lng_d"].astype(float)
    f["type"] = ["Actual" if m < fs else "Forecast" for m in months]
    f.insert(0, "month", months)
    f.index.name = "Month"
    return f


# --------------------------------------------------------------------------- Excel sheets
def _hdr(ws, row, labels, widths=None):
    for j, t in enumerate(labels, start=1):
        c = ws.cell(row, j, t)
        c.font = BOLD
        c.fill = FILL_HEAD
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for j, w in enumerate(widths or [], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w


def write_sheets(path, trains, par, prof, over, f, fs):
    wb = load_workbook(path)
    for n in (AL, "Forecast", "LNG trains calc"):
        if n in wb.sheetnames:
            del wb[n]
    ratio_cell = f"'{AL}'!$B$5"
    a = wb.create_sheet(AL, 1)
    a["A1"] = "LNG assumptions (yellow cells are editable; the Forecast tab recalculates in Excel)"
    a["A1"].font = Font(bold=True, size=13)
    a["A2"] = ("Feedgas = LNG exports x ratio. First-LNG month is a real date (type e.g. 2027-04-01). "
               "Trains already started keep their date in the delayed scenario.")
    for r, (lab, v, note) in enumerate([
            ("Delayed scenario: months of delay", int(par["delay"]), "applied to trains that have not started yet"),
            ("Feedgas / export ratio", float(par["ratio"]), "liquefaction fuel and shrink (~9%), as lng_feedgas_daily.xlsx"),
            ("First forecast month", fs.to_pydatetime(), "set by the script: the month after the last EIA actual (do not edit)")], start=4):
        a.cell(r, 1, lab)
        c = a.cell(r, 2, v)
        if r < 6:
            c.fill = FILL_IN
        a.cell(r, 3, note)
    a["B6"].number_format = "mmm/yy"
    a["A10"] = "New and ramping trains (one row per train; spare rows below can be filled)"
    a["A10"].font = BOLD
    _hdr(a, 11, ["Plant", "Train", "Nameplate feedgas Bcf/d", "First-LNG month", "Ramp months", "Ramp-start utilisation",
                 "Steady utilisation", "Date status", "Source / status note"], [26, 22, 14, 14, 11, 13, 12, 30, 120])
    a.row_dimensions[11].height = 45
    order = {p: i for i, p in enumerate(PLANTS)}
    keys = sorted(trains, key=lambda k: (order.get(k[0], 9), trains[k][1] if not pd.isna(trains[k][1]) else pd.Timestamp("2100-01-01")))
    for i in range(TRAIN_ROWS[1] - TRAIN_ROWS[0] + 1):
        r = TRAIN_ROWS[0] + i
        if i < len(keys):
            k = keys[i]
            n, d, ramp, rs, st, ds, note = trains[k]
            vals = [k[0], k[1], n, d.to_pydatetime() if not pd.isna(d) else None, ramp, rs, st, ds, note]
        else:
            vals = [None] * 9
        for j, v in enumerate(vals, start=1):
            c = a.cell(r, j, v)
            if j <= 7:
                c.fill = FILL_IN
        a.cell(r, 4).number_format = "mmm/yy"
    a["A33"] = "Existing plants: nameplate feedgas (Corpus Christi trains 1-3 + midscale 1-7; Freeport trains 1-3) and monthly utilisation"
    a["A33"].font = BOLD
    a["B34"], a["C34"] = "Corpus Christi", "Freeport"
    a["A35"] = "Nameplate feedgas Bcf/d"
    a["B35"], a["C35"] = float(par["cap_cc"]), float(par["cap_fp"])
    a["B35"].fill = a["C35"].fill = FILL_IN
    a["D35"] = "FERC feed-gas letters (lng_feedgas_daily.xlsx Train start-ups); EIA exports x ratio can exceed it (plants run above nameplate)"
    a["A36"] = (f"Profiles: mean of the same calendar month over the last {HIST_WINDOW} EIA months (utilisation = feedgas / nameplate in operation, smoothed 1-2-1 across neighbouring months; "
                "outages, summer derate and maintenance are inside it). Type a number in an override column to replace one month.")
    _hdr(a, 37, ["Month", "CC utilisation (computed)", "CC override", "CC used", "Freeport utilisation (computed)", "Freeport override",
                 "Freeport used", "New-build seasonal index (computed, mean 1)", "New-build override", "New-build used"])
    a.row_dimensions[37].height = 60
    for i in range(12):
        r = SEAS_ROWS[0] + i
        a.cell(r, 1, i + 1)
        for col, key in ((2, "cc"), (5, "fp"), (8, "nb")):
            a.cell(r, col, float(prof[key][i])).number_format = "0.000"
            ov = a.cell(r, col + 1, over.get(col + 1, [None] * 12)[i])
            ov.fill = FILL_IN
            u = a.cell(r, col + 2, f"=IF({get_column_letter(col + 1)}{r}=\"\",{get_column_letter(col)}{r},{get_column_letter(col + 1)}{r})")
            u.number_format = "0.000"
    # ---- LNG trains calc
    n_tr = TRAIN_ROWS[1] - TRAIN_ROWS[0] + 1
    cs = wb.create_sheet("LNG trains calc", 2)
    cs["A1"] = "Feedgas by train, Bcf/d (formulas on 'Assump - LNG'; base block, then delayed block). Rows match the Forecast tab."
    cs["A1"].font = BOLD
    cs["A3"], cs["A4"], cs["A5"] = "Plant", "Train", "Month"
    for blk, off in (("b", 0), ("d", n_tr)):
        for i in range(n_tr):
            col, r = 2 + off + i, TRAIN_ROWS[0] + i
            cs.cell(3, col, f"=IF('{AL}'!$A{r}=\"\",\"\",'{AL}'!$A{r})")
            cs.cell(4, col, f"=IF('{AL}'!$B{r}=\"\",\"\",'{AL}'!$B{r})")
            cs.cell(5, col, "base" if blk == "b" else "delayed").font = BOLD
    fs_row = FIRST_ROW + (fs.year - FIRST.year) * 12 + fs.month - FIRST.month
    L = get_column_letter
    for rr in range(fs_row, FIRST_ROW + len(f)):
        cs.cell(rr, 1, f.index[rr - FIRST_ROW].to_pydatetime()).number_format = "mmm/yy"
        for blk, off in (("b", 0), ("d", n_tr)):
            for i in range(n_tr):
                r = TRAIN_ROWS[0] + i
                S = f"'{AL}'!$D{r}"
                if blk == "d":
                    S = f"IF({S}>='{AL}'!$B$6,EDATE({S},'{AL}'!$B$4),{S})"
                A_ = f"$A{rr}"
                formula = (f"=IF('{AL}'!$D{r}=\"\",0,IF({A_}<{S},0,'{AL}'!$C{r}*('{AL}'!$F{r}+('{AL}'!$G{r}-'{AL}'!$F{r})"
                           f"*MIN(1,((YEAR({A_})-YEAR({S}))*12+MONTH({A_})-MONTH({S}))/MAX('{AL}'!$E{r},1)))"
                           f"*INDEX('{AL}'!$J${SEAS_ROWS[0]}:$J${SEAS_ROWS[1]},MONTH({A_}))))")
                cs.cell(rr, 2 + off + i, formula).number_format = "0.000"
    cs.freeze_panes = "B6"
    cs.column_dimensions["A"].width = 10
    # ---- Forecast
    w = wb.create_sheet("Forecast", 2)
    w["A1"] = "Texas gas demand incl. LNG, to Dec 2028, Bcf/d - LNG columns are live formulas driven by 'Assump - LNG'"
    w["A1"].font = Font(bold=True, size=13)
    w["A2"] = ("Actual months are EIA (values); forecast months: LNG = formulas, consumption sectors and Mexico = script values (seasonal-trend base: "
               "last 3 years' monthly profile x last-12-month level x damped trend).")
    for j, k in enumerate(KEYS, start=1):
        c = w.cell(5, j, NAME[k])
        c.font = BOLD
        c.fill = FILL_HEAD
        c.alignment = Alignment(wrap_text=True, vertical="top")
        w.column_dimensions[L(j)].width = 13
    w.row_dimensions[5].height = 62
    b_rng = f"'LNG trains calc'!$B$3:${L(1 + n_tr)}$3"
    d_rng = f"'LNG trains calc'!${L(2 + n_tr)}$3:${L(1 + 2 * n_tr)}$3"
    for i, m in enumerate(f.index):
        r = FIRST_ROW + i
        actual = m < fs
        row = f.loc[m]

        def put(key, value):
            if value is None or (isinstance(value, float) and math.isnan(value)):
                return
            c = w[f"{COL[key]}{r}"]
            c.value = value
            c.number_format = "mmm/yy" if key == "month" else ("General" if key == "type" else "0.00")
        put("month", m.to_pydatetime())
        put("type", row["type"])
        for key in ("x_cc", "x_fp", "x_gp"):
            put(key, None if pd.isna(row[key]) else float(row[key]))
        seas = lambda col: f"INDEX('{AL}'!${col}${SEAS_ROWS[0]}:${col}${SEAS_ROWS[1]},MONTH($A{r}))"   # noqa: E731
        if actual:
            for key, x in (("lng_cc_b", "x_cc"), ("lng_fp_b", "x_fp"), ("lng_gp_b", "x_gp")):
                w[f"{COL[key]}{r}"] = f"=N({COL[x]}{r})*{ratio_cell}"
            for key in ("lng_rg_b", "lng_pa_b", "lng_ot_b"):
                w[f"{COL[key]}{r}"] = 0
            for p in ("cc", "fp", "gp", "rg", "pa", "ot"):
                w[f"{COL['lng_' + p + '_d']}{r}"] = f"={COL['lng_' + p + '_b']}{r}"
        else:
            for sfx, rng, lo, hi in (("b", b_rng, 2, 1 + n_tr), ("d", d_rng, 2 + n_tr, 1 + 2 * n_tr)):
                tr = f"'LNG trains calc'!${L(lo)}{r}:${L(hi)}{r}"
                w[f"{COL['lng_cc_' + sfx]}{r}"] = f"='{AL}'!$B$35*{seas('D')}+SUMIF({rng},\"Corpus Christi\",{tr})"
                w[f"{COL['lng_fp_' + sfx]}{r}"] = f"='{AL}'!$C$35*{seas('G')}+SUMIF({rng},\"Freeport\",{tr})"
                for p, nm in (("gp", "Golden Pass"), ("rg", "Rio Grande"), ("pa", "Port Arthur")):
                    w[f"{COL['lng_' + p + '_' + sfx]}{r}"] = f"=SUMIF({rng},\"{nm}\",{tr})"
                w[f"{COL['lng_ot_' + sfx]}{r}"] = (f"=SUM({tr})-SUMIF({rng},\"Corpus Christi\",{tr})-SUMIF({rng},\"Freeport\",{tr})"
                                                   f"-{COL['lng_gp_' + sfx]}{r}-{COL['lng_rg_' + sfx]}{r}-{COL['lng_pa_' + sfx]}{r}")
        for sfx in ("b", "d"):
            w[f"{COL['lng_' + sfx]}{r}"] = f"=SUM({COL['lng_cc_' + sfx]}{r}:{COL['lng_ot_' + sfx]}{r})"
            w[f"{COL['exp_' + sfx]}{r}"] = f"={COL['lng_' + sfx]}{r}/{ratio_cell}"
        for key in ("elec", "ind", "res", "com", "veh", "mex"):
            put(key, None if pd.isna(row[key]) else float(row[key]))
        if not pd.isna(row["cons"]):
            w[f"{COL['cons']}{r}"] = f"=SUM({COL['elec']}{r}:{COL['veh']}{r})"
            w[f"{COL['dem_b']}{r}"] = f"={COL['cons']}{r}+{COL['lng_b']}{r}"
            w[f"{COL['dem_d']}{r}"] = f"={COL['cons']}{r}+{COL['lng_d']}{r}"
        for k in KEYS:
            c = w[f"{COL[k]}{r}"]
            if c.value is not None and k not in ("month", "type"):
                c.number_format = "0.00"
    w.freeze_panes = "C6"
    order_ = ["Units", AL, "Forecast", "LNG trains calc"]
    rest = [s for s in wb._sheets if s.title not in order_]
    wb._sheets = [wb[n] for n in order_ if n in wb.sheetnames] + rest
    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# --------------------------------------------------------------------------- driver
def build(path, cons, exports, bal):
    """Returns (Forecast values frame, context). Writes nothing; write_sheets() adds the Excel tabs."""
    last = bal.index.max()
    fs = last + pd.offsets.MonthBegin(1)
    hist = pd.DataFrame(index=bal.index)
    for c in SECT:
        hist[c] = cons[c].reindex(hist.index)
    hist["mex"] = bal["Pipeline exports to Mexico"]
    for key, col in (("x_cc", "LNG - Corpus Christi"), ("x_fp", "LNG - Freeport"), ("x_gp", "LNG - Golden Pass")):
        hist[key] = exports[col].reindex(hist.index) if col in exports else np.nan
    fut = pd.date_range(fs, LAST, freq="MS")
    ratio = 1.09
    win = hist.index[-HIST_WINDOW:]
    prof, util = {}, {}
    for key, steps, col in (("cc", CC_STEPS, "x_cc"), ("fp", FP_STEPS, "x_fp")):
        u = hist.loc[win, col] * ratio / capacity_by_month(steps, win)
        util[key] = u
        raw_ = [float(u[u.index.month == m].mean()) for m in range(1, 13)]
        # three observations per calendar month are noisy (one trip moves a month): circular 1-2-1 smoothing
        prof[key] = [(raw_[i - 1] + 2 * raw_[i] + raw_[(i + 1) % 12]) / 4 for i in range(12)]
    cap_cc = float(capacity_by_month(CC_STEPS, [fs]).iloc[0])
    cap_fp = float(capacity_by_month(FP_STEPS, [fs]).iloc[0])
    nb = [(prof["cc"][i] * cap_cc + prof["fp"][i] * cap_fp) / (cap_cc + cap_fp) for i in range(12)]
    prof["nb"] = [v / float(np.mean(nb)) for v in nb]
    gtxt = []
    for key, col in (("elec", "Electric power"), ("ind", "Industrial"), ("res", "Residential"), ("com", "Commercial"),
                     ("veh", "Vehicle fuel"), ("mex", "mex")):
        s, g = seasonal_trend(hist[col], fut)
        prof["fc_" + key] = s
        gtxt.append(f"{col if col != 'mex' else 'Mexico exports'} {g:+.1%}/yr")
    prior = read_prior(path)
    trains, par, over = build_assumptions(prior, {"delay": 6, "ratio": ratio, "cap_cc": round(cap_cc, 3), "cap_fp": round(cap_fp, 3)})
    f = run_model(hist, trains, par, prof, over, fs)
    ctx = {"trains": trains, "par": par, "prof": prof, "over": over, "fs": fs, "util": util, "trend_txt": "; ".join(gtxt)}
    return f, ctx


def values_sheet(f):
    """Forecast values frame with readable headers (used by add_charts / png_charts / the masters)."""
    out = f.drop(columns=["month"]).copy()
    out.columns = [NAME[c] for c in out.columns]
    out.index.name = "Month"
    return out


def notes_lines(ctx):
    par = ctx["par"]
    L_ = [
        "",
        "LNG FORECAST (to Dec 2028, Bcf/d)",
        "Tabs: 'Assump - LNG' (yellow cells are editable and survive reruns), 'Forecast' (live Excel formulas for the LNG columns and the demand totals; consumption and Mexico columns are script",
        "values), 'LNG trains calc' (per-train feedgas) and 'Forecast values' (the script's own evaluation of the same logic, used for the charts, PNGs and master dashboards). Forecast tabs are",
        "recomputed on every run; edits to the assumptions update 'Forecast' in Excel at once, the 'Forecast values' tab, charts' PNGs and dashboards follow on the next run.",
        f"LNG feedgas = EIA port exports x {par['ratio']} (liquefaction fuel and shrink, the ratio lng_feedgas_daily.xlsx uses to calibrate its meters); the plain export figure is on the Forecast tab. Base and delayed ({int(par['delay'])} months) scenarios.",
        "Existing Corpus Christi and Freeport = nameplate x the monthly utilisation of the last 36 EIA months, smoothed 1-2-1 across neighbouring months (heat derate, maintenance and Freeport trips included); new trains = the table (nameplate, first-LNG month, ramp, utilisation).",
        "First-LNG months are company statements (NextDecade, Sempra, ExxonMobil, Cheniere, EIA) or assumptions - FERC's LNG pages list no dates (probe). The 'Date status' column marks guidance, assumed and unverified months.",
        f"Other consumption and Mexico exports: seasonal-trend base (last 3 years' monthly profile x last-12-month level x damped trend: {ctx['trend_txt']}). Not weather-adjusted.",
        "Demand incl. LNG = EIA consumption by sector (electric power, industrial, residential, commercial, vehicle fuel) + LNG feedgas; Mexico pipeline exports are separate.",
    ]
    return L_, {"LNG FORECAST (to Dec 2028, Bcf/d)"}
