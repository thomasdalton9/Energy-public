"""
Gas burned for power in South America, Central America and the Caribbean, monthly from 2021: the published
power-sector gas figure where a country publishes one, and an ESTIMATE from the grid operator's gas-fired
generation for the months it has not published yet (and for countries that publish nothing). Built only from this
repo's workbooks - no downloads.

Why: the grid operators' generation feeds run to the last few days, while ministry / regulator gas statistics lag
2-6 months (Brazil's MME split stopped at Jun 2025). Gas-fired MWh x a heat rate gives the power-sector gas burn
months before the statistics, for the whole region on one basis.

Method (the same as the Peru 'Daily power burn (est.)' sheet and americas/MISO_GAS_BURN.py):
    burn (MMBtu/d) = gas-fired generation (MWh/d, monthly mean) x heat rate (MMBtu/MWh)
  * Calibrated countries publish power-sector gas monthly. For each month both exist, the EFFECTIVE heat rate =
    published burn / gas-fired MWh (it absorbs the plants' efficiency and any difference in scope between the gas
    statistic and the operator's 'gas' category). Months whose effective rate is outside 5-15 MMBtu/MWh (a scope or
    data break, e.g. a plant switched to diesel but still tagged gas) are not used for calibration and are flagged.
  * Estimated months use the ratio of sums over the latest 12 calibrated months BEFORE that month (at least 6);
    before the first calibrated year, the first 12 calibrated months.
  * Countries without a published power-sector series use a flat heat rate (ASSUMPTION, stated per country).
  * Burn = published where published, else estimated ('Status' columns say which). Back-test: every published month
    is also estimated as if it were not published yet (trailing rate only) and compared - 'Back-test' sheet.
  Volumes: MMBtu -> million m3 at each source's own heat content (stated per country), -> Bcf/d x 0.0353147.

Usage: python3 south_america/SA_GAS_BURN_POWER.py [--data-dir DIR] [--out PATH]
"""
import argparse
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import add_charts  # noqa: E402
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(DATA_DIR, "south_america_gas_burn_power.xlsx")
START = "2021-01-01"
BCF_PER_MCM = 0.0353147
STD_HC = 36374.0                       # MMBtu per million m3 at 1,030 Btu/cf (repo standard)
HR_MIN, HR_MAX = 5.0, 15.0             # plausible effective heat rates, MMBtu/MWh
TRAIL, TRAIL_MIN = 12, 6               # months in the trailing calibration window / minimum needed
COVER_MIN = 0.95                       # the regional total needs countries with >= 95% of the recent burn
MIN_DAYS_SHARE = 0.8                   # a generation month needs >= 80% of its days (add_charts.complete_months)


def _monthly(path, sheet, date_col):
    d = add_charts.by_date(add_charts.read(path, sheet), date_col)
    d.index = d.index.to_period("M").to_timestamp()
    return d


def _per_day(s):
    return s / pd.Series(s.index.days_in_month, index=s.index)


# Published power-sector gas, MMBtu per day (monthly average). Each returns a Series indexed by month start.
def rep_argentina(dd):
    d = _monthly(os.path.join(dd, "argentina_gas_monthly.xlsx"), "National", "date")
    return _per_day(pd.to_numeric(d["centrales_electricas"], errors="coerce")) * HC.get("AR", STD_HC)


def rep_brazil(dd):
    d = _monthly(os.path.join(dd, "brazil_gas_monthly.xlsx"), "Demand by segment", "date")
    return pd.to_numeric(d["Power_Generation"], errors="coerce") * HC.get("BR", STD_HC)


def rep_bolivia(dd):
    d = _monthly(os.path.join(dd, "bolivia_gas_demand_by_sector.xlsx"), "Demand by sector", "Month")
    return pd.to_numeric(d["Power_mcm_per_day"], errors="coerce") * HC.get("BO", STD_HC)


def rep_colombia(dd):
    d = _monthly(os.path.join(dd, "colombia_gas_demand_by_sector.xlsx"), "Demand by sector", "Month")
    return pd.to_numeric(d["Power"], errors="coerce") * 1000.0          # GBTUD -> MMBtu/d


def rep_ecuador(dd):
    d = _monthly(os.path.join(dd, "ecuador_gas.xlsx"), "Total demand", "Month")
    return pd.to_numeric(d["Total_power_MMBtu_per_day"], errors="coerce")


def rep_peru(dd):
    d = _monthly(os.path.join(dd, "peru_gas_demand_by_sector.xlsx"), "Demand by sector", "Month")
    return pd.to_numeric(d["Power_MMPCD"], errors="coerce") * 1065.0       # MMcf/d x 1,065 Btu/scf (MINEM basis)


def rep_dominican(dd):
    d = _monthly(os.path.join(dd, "dominican_republic_gas.xlsx"), "Gas use", "Month")
    return pd.to_numeric(d["Power_MMBtu_per_day"], errors="coerce")


def rep_puerto_rico(dd):
    d = _monthly(os.path.join(dd, "puerto_rico_gas.xlsx"), "Gas use", "Month")
    cols = [c for c in d.columns if str(c).startswith("Power_") and str(c).endswith("_mcm_per_day")]
    return d[cols].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1) * HC.get("PR", STD_HC)


# Ecuador's daily CENACE feed only starts in Sep 2026; its earlier gas-fired output is the monthly CENACE column the
# Ecuador gas workbook already carries (cross-check sheet).
def gen_ecuador_history(dd):
    d = _monthly(os.path.join(dd, "ecuador_gas.xlsx"), "Total demand", "Month")
    return pd.to_numeric(d["CENACE_gas_generation_MWh_per_day"], errors="coerce").dropna()


# Heat content of each source's volumes, MMBtu per million m3
HC = {"AR": 36905.0,    # 9,300 kcal/m3 (ENARGAS standard)
      "PE": 37610.0,    # 1,065 Btu/scf (MINEM, as in the Peru workbook)
      "DO": 36621.0, "PR": 36621.0}   # 1,037 Btu/cf (EIA average, as in the DR / PR workbooks)

# (code, country, generation workbook, published-burn function, published-series label, flat heat rate if none)
COUNTRIES = [
    ("AR", "Argentina", "argentina_power_generation_daily.xlsx", rep_argentina,
     "Secretaría de Energía / ENARGAS 'centrales eléctricas' (gas delivered to power plants)", None),
    ("BR", "Brazil", "brazil_power_generation_daily.xlsx", rep_brazil,
     "MME gas bulletin 'Geração elétrica' (national, to Jun 2025 - nothing newer published)", None),
    ("BO", "Bolivia", "bolivia_power_generation_daily.xlsx", rep_bolivia, "INE 'generadoras eléctricas'", None),
    ("CL", "Chile", "chile_power_generation_daily.xlsx", None, None, 7.5),
    ("CO", "Colombia", "colombia_power_generation_daily.xlsx", rep_colombia, "Gestor del Mercado (BMC) thermal demand", None),
    ("EC", "Ecuador", "ecuador_power_generation_daily.xlsx", rep_ecuador,
     "Petroecuador gas to Termogas Machala + customs imports for power", None),
    ("PE", "Peru", "peru_power_generation_daily.xlsx", rep_peru, "MINEM (DGH) power-sector gas", None),
    ("PA", "Panama", "panama_power_generation_daily.xlsx", None, None, 7.0),
    ("SV", "El Salvador", "el_salvador_power_generation_daily.xlsx", None, None, 8.2),
    ("DO", "Dominican Republic", "dominican_republic_power_generation_daily.xlsx", rep_dominican,
     "SIE gas burned for power (SENI)", None),
    ("PR", "Puerto Rico", "puerto_rico_power_generation_daily.xlsx", rep_puerto_rico,
     "EIA-923 gas burned at the power plants", None),
    # Checked for gas-fired output and included automatically if any appears: Costa Rica, Guatemala, Honduras,
    # Nicaragua, Belize (none in 2021-26). Not included: Uruguay (its feed has no gas column; Punta del Tigre runs
    # mostly on diesel), Trinidad & Tobago (no generation feed; MEEI publishes power-sector gas itself), Jamaica
    # (annual data only), Paraguay (no gas plants), Mexico (North America master).
    ("CR", "Costa Rica", "costa_rica_power_generation_daily.xlsx", None, None, 7.5),
    ("GT", "Guatemala", "guatemala_power_generation_daily.xlsx", None, None, 7.5),
    ("HN", "Honduras", "honduras_power_generation_daily.xlsx", None, None, 7.5),
    ("NI", "Nicaragua", "nicaragua_power_generation_daily.xlsx", None, None, 7.5),
    ("BZ", "Belize", "belize_power_generation_daily.xlsx", None, None, 7.5),
]
FLAT_NOTE = {"CL": "ASSUMPTION: 7.5 MMBtu/MWh (combined cycles at part load plus open-cycle units; no Chilean "
                   "power-sector gas statistic is published monthly - CNE's import split is by region, not by use)",
             "PA": "ASSUMPTION: 7.0 MMBtu/MWh, the same as panama_gas.xlsx (AES Colón combined cycle)",
             "SV": "ASSUMPTION: 8.2 MMBtu/MWh, the same as el_salvador_gas.xlsx (Energía del Pacífico engines)"}


def gas_generation(dd, code, fname):
    """Gas-fired generation, MWh per day (monthly mean of the days present), months with >= 80% of their days."""
    path = os.path.join(dd, fname)
    out = pd.Series(dtype=float)
    if os.path.exists(path):
        d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
        if "Gas_MWh" in d:
            g = pd.to_numeric(d["Gas_MWh"], errors="coerce").dropna()
            g = g[g.index >= START]
            if len(g):
                n = g.resample("MS").size()
                m = g.resample("MS").mean()
                ok = n[n >= MIN_DAYS_SHARE * n.index.days_in_month].index
                monthly_rows = len(g) > 2 and g.index.to_series().diff().median().days > 20
                if monthly_rows:   # a feed with one row per month holds the month's total: -> per day
                    m = g.resample("MS").sum() / pd.Series(g.resample("MS").sum().index.days_in_month,
                                                           index=g.resample("MS").sum().index)
                    ok = g.index.to_period("M").to_timestamp().unique()
                out = m[m.index.isin(ok)]
    if code == "EC":
        hist = gen_ecuador_history(dd)
        out = pd.concat([hist[~hist.index.isin(out.index)], out]).sort_index()
    if out.empty:
        return pd.Series(dtype=float, index=pd.DatetimeIndex([]))
    return out[out.index >= START]


def trailing_rate(hr_ok, burn, gen, month):
    """Ratio of sums over the latest TRAIL calibrated months before `month` (else the first TRAIL calibrated)."""
    prior = hr_ok[hr_ok.index < month].index[-TRAIL:]
    if len(prior) < TRAIL_MIN:
        prior = hr_ok.index[:TRAIL]
        if len(prior) < TRAIL_MIN:
            return None, None
        basis = f"first {len(prior)} calibrated months ({prior[0]:%b/%y}-{prior[-1]:%b/%y})"
    else:
        basis = f"trailing {len(prior)} calibrated months ({prior[0]:%b/%y}-{prior[-1]:%b/%y})"
    return float(burn[prior].sum() / gen[prior].sum()), basis


def build(dd):
    detail, notes, log = [], [], []
    for code, country, fname, rep_fn, rep_label, flat in COUNTRIES:
        try:
            gen = gas_generation(dd, code, fname)
        except Exception as e:  # noqa: BLE001
            log.append(f"{country}: generation not read ({type(e).__name__}: {e})")
            continue
        hc = HC.get(code, STD_HC)
        rep = pd.Series(dtype=float)
        if rep_fn is not None:
            try:
                rep = rep_fn(dd).dropna()
                rep = rep[(rep.index >= START) & (rep > 0)]
            except Exception as e:  # noqa: BLE001
                log.append(f"{country}: published series not read ({type(e).__name__}: {e}) - flat rate used")
        if gen.empty or gen.max() <= 0:
            if code in ("CR", "GT", "HN", "NI", "BZ") and rep.empty:
                continue   # no gas-fired output in the feed: nothing to show
            log.append(f"{country}: no gas-fired generation months in {fname} - published months only, nothing "
                       f"estimated")
            gen = pd.Series(dtype=float, index=pd.DatetimeIndex([]))
            if rep.empty:
                continue
        months = gen.index.union(rep.index)
        both = gen.index.intersection(rep.index)
        hr_all = (rep[both] / gen[both].where(gen[both] > 0)).dropna()
        hr_ok = hr_all[(hr_all >= HR_MIN) & (hr_all <= HR_MAX)]
        bad = hr_all.index.difference(hr_ok.index)
        calibrated = len(hr_ok) >= TRAIL_MIN
        if rep_fn is not None and not calibrated and len(gen):
            log.append(f"{country}: only {len(hr_ok)} usable calibration months - flat 7.5 MMBtu/MWh used")
        for m in months:
            g = gen.get(m)
            r = rep.get(m)
            if calibrated:
                rate, basis = trailing_rate(hr_ok, rep, gen, m)
            else:
                rate, basis = (flat or 7.5), FLAT_NOTE.get(code, f"ASSUMPTION: flat {flat or 7.5} MMBtu/MWh")
            est = g * rate if (g is not None and pd.notna(g) and rate) else None
            published = r is not None and pd.notna(r)
            burn = r if published else est
            if burn is None:
                continue
            status = "published" if published else "estimated"
            detail.append({"Country": country, "Code": code, "Month": m,
                           "Gas_generation_MWh_per_day": g, "Published_MMBtu_per_day": r,
                           "Effective_heat_rate": hr_all.get(m), "Calibration_month": m in hr_ok.index,
                           "Heat_rate_flag": "outside 5-15, not used" if m in bad else "",
                           "Estimation_heat_rate": rate, "Estimation_basis": basis,
                           "Estimate_MMBtu_per_day": est, "Burn_MMBtu_per_day": burn, "Status": status,
                           "Heat_content_MMBtu_per_mcm": hc,
                           "Burn_mcm_per_day": burn / hc, "Burn_Bcf_per_day": burn / hc * BCF_PER_MCM})
        last_pub = rep.index.max() if len(rep) else None
        notes.append({"Country": country, "Published power-sector series": rep_label or "none (flat heat rate)",
                      "Heat rate basis": (f"calibrated: {len(hr_ok)} months" + (f", {len(bad)} months outside "
                                          f"5-15 left out" if len(bad) else "")) if calibrated
                      else "published months only (no gas-fired generation history yet)" if not len(gen)
                      else FLAT_NOTE.get(code, f"ASSUMPTION: flat {flat or 7.5} MMBtu/MWh"),
                      "Latest effective heat rate (12m)": (round(float(rep[hr_ok.index[-TRAIL:]].sum() /
                                                            gen[hr_ok.index[-TRAIL:]].sum()), 2) if calibrated else None),
                      "Heat content MMBtu per million m3": hc,
                      "Generation source": fname,
                      "First month": months.min().strftime("%b/%y"),
                      "Last published month": last_pub.strftime("%b/%y") if last_pub is not None else "",
                      "Last generation month": gen.index.max().strftime("%b/%y") if len(gen) else "",
                      "Months estimated ahead of publication": (
                          (gen.index.max().to_period("M") - last_pub.to_period("M")).n
                          if last_pub is not None and len(gen) and gen.index.max() > last_pub else 0)})
    return pd.DataFrame(detail), pd.DataFrame(notes), log


def backtest(detail):
    """Each published month re-estimated with the trailing rate only (what the estimate would have said then)."""
    rows, series = [], {}
    for country, d in detail.groupby("Country", sort=False):
        d = d.set_index("Month").sort_index()
        t = d[d["Status"].eq("published") & d["Estimate_MMBtu_per_day"].notna()
              & d["Estimation_basis"].astype(str).str.startswith("trailing")]
        if t.empty:
            continue
        err = (t["Estimate_MMBtu_per_day"] / t["Published_MMBtu_per_day"] - 1) * 100
        last12 = err[err.index > err.index.max() - pd.DateOffset(months=12)]
        rows.append({"Country": country, "Months tested": len(t), "From": f"{t.index.min():%b/%y}",
                     "To": f"{t.index.max():%b/%y}", "Mean error %": round(err.mean(), 1),
                     "Mean absolute error %": round(err.abs().mean(), 1),
                     "Mean absolute error %, latest 12": round(last12.abs().mean(), 1),
                     "Worst month": f"{err.abs().idxmax():%b/%y} ({err[err.abs().idxmax()]:+.0f}%)"})
        hc = d["Heat_content_MMBtu_per_mcm"].iloc[0]
        series[country] = pd.DataFrame({
            "Published": d["Published_MMBtu_per_day"] / hc * BCF_PER_MCM,
            "Estimate from gas-fired output": d["Estimate_MMBtu_per_day"] / hc * BCF_PER_MCM})
    return pd.DataFrame(rows), series


def wide(detail, col):
    w = detail.pivot_table(index="Month", columns="Country", values=col, aggfunc="first")
    return w[[c for c in detail["Country"].unique() if c in w.columns]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    detail, countries, log = build(args.data_dir)
    if detail.empty:
        sys.exit("No country could be built: " + "; ".join(log))
    bt, bt_series = backtest(detail)
    bcfd = wide(detail, "Burn_Bcf_per_day")
    status = wide(detail, "Status")
    est = bcfd.where(status.eq("estimated"))
    # The regional total runs to the last month whose countries made up at least COVER_MIN of the burn over the
    # previous 12 months (so a small, late-publishing country does not hold back the whole region); a country
    # missing from a month of the total is named on 'Total coverage', never filled.
    end, cover_rows = None, []
    for m in bcfd.index:
        prev = bcfd[(bcfd.index < m) & (bcfd.index >= m - pd.DateOffset(months=12))].mean()
        if prev.dropna().empty:
            continue
        have = bcfd.loc[m].notna()
        share = prev[have.reindex(prev.index, fill_value=False)].sum() / prev.sum()
        missing = [c for c in prev.index if not have.get(c, False) and prev[c] > 0]
        cover_rows.append({"date": m, "Share of the region's last-12-month burn present (%)": round(100 * share, 1),
                           "Countries missing": ", ".join(missing)})
        if share >= COVER_MIN:
            end = m
    coverage = pd.DataFrame(cover_rows).set_index("date")
    coverage = coverage[coverage.index <= end]
    monthly = bcfd[bcfd.index <= end].copy()
    monthly["Total"] = monthly.sum(axis=1, min_count=1)
    monthly["Of which estimated"] = est[est.index <= end].sum(axis=1, min_count=1).fillna(0)
    monthly.index.name = "date"
    mcmd = wide(detail, "Burn_mcm_per_day")
    mcmd.index.name = "date"
    hr = wide(detail, "Effective_heat_rate")
    hr.index.name = "date"
    status.index.name = "date"
    bt_wide = pd.concat({c: s for c, s in bt_series.items()}, axis=1)
    bt_wide.columns = [f"{c} | {k}" for c, k in bt_wide.columns]
    bt_wide.index.name = "date"

    lines = [
        "GAS BURN FOR POWER - SOUTH AMERICA, CENTRAL AMERICA AND THE CARIBBEAN",
        "Built by south_america/SA_GAS_BURN_POWER.py from this repo's workbooks (no downloads), 1st and 15th, before "
        "the South & Central America master.",
        "",
        "WHAT IT IS",
        "Power-sector gas burn per country, monthly average, from Jan 2021: the PUBLISHED figure where the country "
        "publishes one, and an ESTIMATE = gas-fired generation (grid operator, monthly mean MWh/day) x heat rate for "
        "months not published yet and for countries that publish nothing. 'Status' says which, month by month. "
        "Estimates are not measured data.",
        "",
        "HEAT RATE",
        "Calibrated countries: effective heat rate = published burn / gas-fired MWh for each month both exist (it "
        "absorbs plant efficiency AND any scope difference between the gas statistic and the operator's 'gas' "
        "category). Months outside 5-15 MMBtu/MWh are not used (flagged on 'Detail'). An estimated month uses the "
        "ratio of sums over the latest 12 calibrated months before it (minimum 6).",
        "Flat-rate countries (no published power-sector gas): Chile 7.5, Panama 7.0, El Salvador 8.2 MMBtu/MWh - "
        "ASSUMPTIONS, error roughly +/-10-15% from the heat rate alone.",
        "Brazil: calibrated on MME's national power-generation gas (complete scope, incl. LNG-to-power and Parnaíba), "
        "which stops at Jun 2025, so every later month is estimated on the Jul 2024 - Jun 2025 rate. ANP's grid "
        "series is NOT used: it leaves out plants fed off the transport grid (about a third of power gas).",
        "Argentina: 'centrales eléctricas' is gas delivered through the distribution/transport system; plants fed at "
        "the wellhead (ENARGAS 'Off System', Supply net sheet) are not in it, so the estimate follows the published "
        "scope.",
        "Ecuador: Termogas Machala only (the one gas plant); generation history from the CENACE column of "
        "ecuador_gas.xlsx before the daily CENACE feed starts (Sep 2026).",
        "",
        "UNITS",
        "Bcf/d (billion cubic feet per day, monthly average) = million m3/day x 0.0353147. MMBtu are converted to m3 at "
        "each source's own heat content: Argentina 9,300 kcal/m3 (36,905 MMBtu per million m3), Peru 1,065 Btu/scf "
        "(37,610), Dominican Republic and Puerto Rico 1,037 Btu/cf (36,621), others 1,030 Btu/cf (36,374).",
        "",
        "SHEETS",
        "Monthly (Bcf/d): burn by country, the regional Total and the part of it that is estimated. "
        "The Total runs to the last month whose countries made up at least 95% of the region's burn over the "
        "previous 12 months; a country missing from a month of it (a late publisher such as El Salvador's SIGET "
        "or EIA-923 for Puerto Rico) is named on 'Total coverage' - nothing is filled. Monthly mcm per day: the same in million m3/day. Status: published / "
        "estimated per country and month. Heat rates: effective heat rate per calibration month. Detail: one row "
        "per country and month with every input. Countries: sources, heat-rate basis, and how many months the "
        "estimate runs ahead of publication. Back-test: each published month re-estimated with the trailing rate "
        "only (what the estimate would have said before publication) - mean and mean absolute error. Back-test "
        "series: the two lines per country.",
        "",
        "NOT INCLUDED",
        "Uruguay (generation feed has no gas column), Trinidad & Tobago (no generation feed; MEEI publishes power gas "
        "itself - see trinidad_gas.xlsx), Jamaica (annual data only), Mexico (North America master; MISO/ERCOT burn "
        "there). Costa Rica, Guatemala, Honduras, Nicaragua and Belize are checked and appear only if their feeds "
        "show gas-fired output.",
    ]
    if log:
        lines += ["", "THIS RUN"] + log
    sheets = {"Monthly (Bcf per day)": monthly, "Total coverage": coverage,
              "Monthly mcm per day": mcmd, "Status": status, "Heat rates": hr,
              "Detail": detail.set_index("Country"), "Countries": countries.set_index("Country"),
              "Back-test": bt.set_index("Country") if not bt.empty else bt, "Back-test series": bt_wide}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, lines,
                              ["WHAT IT IS", "HEAT RATE", "UNITS", "SHEETS", "NOT INCLUDED", "THIS RUN"])
    print(f"Saved {args.out}: {len(bcfd.columns)} countries, {monthly.index.min():%b/%y}-{monthly.index.max():%b/%y}")
    print(countries[["Country", "Heat rate basis", "Last published month", "Last generation month",
                     "Months estimated ahead of publication"]].to_string(index=False))
    if not bt.empty:
        print(bt.to_string(index=False))
    for x in log:
        print(x)


if __name__ == "__main__":
    main()
