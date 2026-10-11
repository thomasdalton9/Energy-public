"""
Canada natural gas demand by sector and PROVINCE, plus gas burned for electricity, from Statistics Canada (no key):

  Table 25-10-0086-01  Natural gas supply and disposition, monthly      (residential / commercial / industrial by
                       province, from 2019; the national pull in CANADA_STATCAN.py keeps only the Canada total)
  Table 25-10-0029-01  Energy supply and demand, annual, terajoules     (REPORTED natural gas 'transformed to
                       electricity' by utilities and by industry, by province - the only reported gas-for-power series)
  Table 25-10-0015-01  Electric power generation, monthly               (non-renewable combustible generation by
                       province, the monthly profile used to spread the annual gas burn)
      -> canada_gas_by_province.xlsx

Monthly gas for power is DERIVED, not reported: each province's annual reported TJ (utilities + industry) is spread over
its twelve months in proportion to that month's combustible-fuel generation (the calibration approach used for
Singapore; no heat-rate assumption). Months after the last annual year carry forward that year's TJ per MWh and are
marked 'estimated'. Gas for power is NOT subtracted from the reported Industrial series: whether utility gas sits inside
StatCan's 'Industrial' is tested in the 'Annual check' sheet (annual power gas vs annual industrial, by province).

Incremental: raw sheets (gas by province, annual power gas, combustible generation) are the saved history; each table is
downloaded only when StatCan's WDS metadata shows a newer period, and the derived sheets are rebuilt from the raw
sheets on every run.

Usage: python3 CANADA_GAS_BY_PROVINCE.py [--out ...] [--force]
"""
import argparse
import os
import re
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import CANADA_STATCAN as cs  # noqa: E402  (shared StatCan download / incremental helpers)

GAS_TABLE, ENERGY_TABLE, POWER_TABLE = 25100086, 25100029, 25100015
TJ_PER_MCM = 38.3          # 1 million m3 of natural gas = 38.3 TJ (38.3 MJ/m3 gross); one constant for the workbook
AGGREGATES = {"Canada", "Atlantic provinces", "Yukon, Northwest Territories and Nunavut"}
GAS_ITEMS = [  # StatCan item (regex on the start of the label) -> short sector name
    (r"^residential consumption", "Residential"),
    (r"^commercial consumption", "Commercial"),
    (r"^industrial consumption", "Industrial"),
    (r"^industrial deliveries from transmission", "Industrial (transmission)"),
    (r"^industrial deliveries from distribution", "Industrial (distribution)"),
    (r"^deliveries to natural gas processing plants", "Processing plants"),
    (r"^pipeline fuel", "Pipeline fuel"),
]
UTIL, INDUS = "Transformed to electricity by utilities", "Transformed to electricity by industry"
SECTORS = ("Residential", "Commercial", "Industrial")


def col_of(dims, *words):
    return next(c for c in dims if all(w in c.lower() for w in words))


def province_gas(d, dims):
    """Monthly gas by province and sector, million m3/day: columns 'Province|Sector'."""
    item = next(c for c in dims if "unit" not in c.lower())
    d = d[~d["GEO"].isin(AGGREGATES)]
    d = d[d["UOM"].str.contains("metre", case=False, na=False)].copy()
    d["sector"] = d[item].map(lambda t: next((s for pat, s in GAS_ITEMS if re.search(pat, str(t), re.I)), None))
    d = d.dropna(subset=["sector", "value"])
    w = d.pivot_table(index="date", columns=["GEO", "sector"], values="value", aggfunc="sum")
    out = w.div(w.index.days_in_month, axis=0) / 1e6
    out.columns = [f"{g}|{s}" for g, s in out.columns]
    out = out.dropna(axis=1, how="all").round(3)
    out.index.name = "Month"
    return out


def annual_power_gas(d, dims):
    """Reported annual natural gas transformed to electricity, TJ: columns 'Province|Utilities_TJ' / 'Province|Industry_TJ'."""
    fuel, char = col_of(dims, "fuel"), col_of(dims, "characteristics")
    d = d[d[fuel].str.strip().str.lower().eq("natural gas") & d[char].isin([UTIL, INDUS])]
    d = d[~d["GEO"].isin(AGGREGATES) & d["UOM"].str.contains("terajoule", case=False, na=False)].dropna(subset=["value"])
    d = d.assign(part=d[char].map({UTIL: "Utilities_TJ", INDUS: "Industry_TJ"}))
    w = d.pivot_table(index="date", columns=["GEO", "part"], values="value", aggfunc="sum")
    w.columns = [f"{g}|{p}" for g, p in w.columns]
    w.index.name = "Year"
    return w.round(0)


def combustible_gen(d, dims):
    """Monthly non-renewable combustible generation (coal + gas + oil), MWh, by province (all classes of producer).
    StatCan carries the 'non-renewable' total only from 2020; before that combustible minus biomass."""
    cls, typ = col_of(dims, "class"), col_of(dims, "type")
    d = d[d[cls].str.contains("total all classes", case=False, na=False) & ~d["GEO"].isin(AGGREGATES)]
    d = d[d["UOM"].str.contains("megawatt", case=False, na=False)]
    piv = lambda pat: d[d[typ].str.contains(pat, case=False, na=False, regex=True)].pivot_table(  # noqa: E731
        index="date", columns="GEO", values="value", aggfunc="sum")
    nonren, comb, bio = piv(r"non-renewable combustible"), piv(r"from combustible fuels"), piv(r"from biomass")
    out = nonren.combine_first(comb.sub(bio.reindex_like(comb).fillna(0)))
    out = out.dropna(axis=1, how="all").round(0)
    out.index.name = "Month"
    return out


def spread_annual(annual, gen):
    """Monthly gas for power, million m3/day, by province, plus a 'basis' frame ('derived' / 'estimated')."""
    burn, basis = {}, {}
    for prov in sorted({c.split("|")[0] for c in annual.columns}):
        if prov not in gen.columns:
            continue
        tj = annual[[c for c in annual.columns if c.startswith(prov + "|")]].sum(axis=1, min_count=1).dropna().sort_index()
        g = gen[prov].dropna()
        months, how = {}, {}
        last_ratio = None
        for ts in tj.index:
            gy = g[g.index.year == ts.year]
            if len(gy) == 12 and gy.sum() > 0:
                for m, v in gy.items():
                    months[m], how[m] = tj[ts] * v / gy.sum(), "derived"
                last_ratio = tj[ts] / gy.sum()          # TJ per MWh in the latest full annual year
        if last_ratio is not None:
            for m, v in g[g.index > max(months)].items():
                months[m], how[m] = last_ratio * v, "estimated"
        if months:
            s = pd.Series(months).sort_index()
            burn[prov] = s / TJ_PER_MCM / s.index.days_in_month
            basis[prov] = pd.Series(how).reindex(s.index)
    out, bas = pd.DataFrame(burn).round(3), pd.DataFrame(basis)
    out.index.name = bas.index.name = "Month"
    return out, bas


def annual_check(gas, annual):
    """Annual means (million m3/day) of reported Industrial gas and reported gas for power, by province, plus the ratio."""
    rows = {}
    for prov in sorted({c.split("|")[0] for c in annual.columns}):
        ind = gas.get(f"{prov}|Industrial")
        if ind is None:
            continue
        full = ind.dropna().groupby(ind.dropna().index.year).filter(lambda s: len(s) == 12)
        yearly = full.groupby(full.index.year).mean()
        for ts, r in annual.iterrows():
            if ts.year not in yearly.index:
                continue
            util, indus = r.get(f"{prov}|Utilities_TJ"), r.get(f"{prov}|Industry_TJ")
            days = 366 if pd.Timestamp(ts.year, 12, 31).dayofyear == 366 else 365
            to_mcm = lambda x: None if pd.isna(x) else x / TJ_PER_MCM / days  # noqa: E731
            tot = pd.Series([util, indus]).sum(min_count=1)
            rows[(prov, ts.year)] = {"Province": prov, "Year": ts.year, "Industrial_gas_mcm_d": round(yearly[ts.year], 3),
                                     "Power_utilities_mcm_d": to_mcm(util), "Power_industry_mcm_d": to_mcm(indus),
                                     "Power_total_mcm_d": to_mcm(tot),
                                     "Power_total_vs_Industrial": None if not yearly[ts.year] or pd.isna(tot)
                                     else round(tot / TJ_PER_MCM / days / yearly[ts.year], 3)}
    return pd.DataFrame(list(rows.values())).set_index("Province") if rows else pd.DataFrame()


def long_table(gas, burn, basis):
    parts = []
    for c in gas.columns:
        prov, sector = c.split("|")
        parts.append(pd.DataFrame({"province": prov, "sector": sector, "mcm_d": gas[c], "basis": "reported"}))
    for prov in burn.columns:
        parts.append(pd.DataFrame({"province": prov, "sector": "Power (utilities + industry)", "mcm_d": burn[prov],
                                   "basis": basis[prov].map({"derived": "derived: annual reported, monthly profile",
                                                             "estimated": "estimated: latest TJ/MWh carried forward"})}))
    out = pd.concat(parts).dropna(subset=["mcm_d"]).reset_index().rename(columns={"index": "date", "Month": "date"})
    return out.sort_values(["province", "sector", "date"]).set_index("date")


def refresh(pid, saved, force, build, label):
    """Download table `pid` only if newer than `saved`; returns the merged raw sheet."""
    print(label, flush=True)
    if not cs.needs_update(pid, saved, force):
        return saved
    new = build(*cs.download(pid))
    new.index = pd.to_datetime(new.index)
    return cs.merge(new, saved)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(cs.DEFAULT_DIR, "canada_gas_by_province.xlsx"))
    ap.add_argument("--force", action="store_true", help="download even if the saved data look current")
    args = ap.parse_args()
    s = lambda name: cs.load(args.out, name)  # noqa: E731
    gas = refresh(GAS_TABLE, s("Raw gas by province"), args.force, province_gas, "Gas by province (25-10-0086-01)")
    annual = refresh(ENERGY_TABLE, s("Raw annual power gas"), args.force, annual_power_gas,
                     "Reported annual gas for electricity (25-10-0029-01)")
    gen = refresh(POWER_TABLE, s("Raw combustible generation"), args.force, combustible_gen,
                  "Combustible generation by province (25-10-0015-01)")
    if gas.empty or annual.empty or gen.empty:
        raise SystemExit("A StatCan table returned nothing; nothing written")
    burn, basis = spread_annual(annual, gen)
    check = annual_check(gas, annual)
    print(burn.tail(3).to_string(), flush=True)
    print(check.tail(12).to_string(), flush=True)
    notes = [
        "UNITS",
        "Gas by province: million cubic metres per day (mcm/d), monthly volume / days in month, as StatCan reports it "
        "(Residential, Commercial, Industrial and its transmission / distribution deliveries, plant deliveries, pipeline "
        "fuel). Newfoundland, PEI, Yukon and Nunavut have no values. Annual power gas: terajoules (TJ). Generation: MWh.",
        f"1 million m3 = {TJ_PER_MCM} TJ (38.3 MJ/m3 gross) - one assumption, used for every TJ -> mcm conversion here.",
        "",
        "METHOD",
        "'Power burn (derived)': each province's REPORTED annual natural gas transformed to electricity (table "
        "25-10-0029-01, utilities + industry, TJ) is spread over the year's twelve months in proportion to the month's "
        "non-renewable combustible generation (table 25-10-0015-01: coal + gas + oil; 'combustible minus biomass' before "
        "2020). Months after the latest annual year carry that year's TJ per MWh forward ('estimated'). Where a province "
        "still burns coal or oil the monthly split is approximate; the annual totals are exact.",
        "Gas for power is NOT removed from 'Industrial': 'Annual check' compares annual power gas with annual Industrial "
        "gas by province (a ratio above 1 means utility power gas cannot sit inside 'Industrial').",
        "'Long' holds every series in one table: date, province, sector, mcm_d, basis (reported / derived / estimated).",
        "",
        "COVERAGE",
        f"Gas by province {gas.index.min():%b %Y} to {gas.index.max():%b %Y}; annual power gas {annual.index.min():%Y} to "
        f"{annual.index.max():%Y} (StatCan publishes about 18 months after the year); generation to "
        f"{gen.index.max():%b %Y}.",
        "",
        "SOURCE",
        "Statistics Canada, Table 25-10-0086-01 Natural gas supply and disposition (monthly): "
        "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2510008601 ; Table 25-10-0029-01 Supply and demand of primary "
        "and secondary energy in terajoules (annual): https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2510002901 ; "
        "Table 25-10-0015-01 Electric power generation, monthly generation by type of electricity: "
        "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2510001501",
    ]
    sheets = {"Power burn (derived)": burn, "Annual check": check, "Long": long_table(gas, burn, basis),
              "Raw gas by province": gas, "Raw annual power gas": annual, "Raw combustible generation": gen}
    cs.save(args.out, sheets, notes)


if __name__ == "__main__":
    main()
