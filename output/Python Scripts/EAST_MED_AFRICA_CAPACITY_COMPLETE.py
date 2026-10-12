"""
East Mediterranean and Africa installed capacity, ONE gap-free monthly dataset per country and fuel (Jan 2021 - latest month, MW),
every filled cell flagged.  Two workbooks in output/Data and Chart Outputs/:
    east_med_capacity_monthly_complete.xlsx   Israel, Egypt, Cyprus, Jordan, Lebanon     (East Med master countries + Jordan, Lebanon)
    africa_capacity_monthly_complete.xlsx     Egypt, Ghana, Nigeria, South Africa, Morocco, Cameroon and the West Africa
                                              power-pool / Ember-only countries of the Africa master
Egypt is built once and written to both.

Same approach and sheet layout as europe/capacity/EUROPE_CAPACITY_COMPLETE.py and the generic rules in the repo-root
`capacity_fill.py` (imported, not copied; see its docstring for the flags). No downloads: this reads
  * official ANNUAL capacity workbooks where a pull exists (see OFFICIAL below: each entry says source, fuels and why), and
  * Ember yearly installed capacity by fuel from long_term_energy.xlsx (LONG_TERM_ENERGY.py, one whole-file download; columns
    Cap_<fuel>_GW), used as FALLBACK_EMBER wherever no official series holds a fuel.
Ember (and the national annual tables) give year-END stocks: a value for year Y is dated 1 January of Y+1 (capacity_fill convention).
ONE basis per country and fuel for the whole span (an official source for a fuel replaces Ember for that fuel). Fill rules: wind, solar,
bioenergy linear between annual stocks, every other fuel held (HELD_A); flat after the last stock (HELD_LATEST); Ember cells are all
FALLBACK_EMBER. Nothing is guessed: no source = NONE. Which official sources were tried and not usable is listed on the 'Tried' sheet.

Usage: python3 EAST_MED_AFRICA_CAPACITY_COMPLETE.py [--data-dir DIR] [--end YYYY-MM] [--no-charts]
"""
import argparse
import datetime
import os
import sys

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
import capacity_fill as cf  # noqa: E402
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT_EM = "east_med_capacity_monthly_complete.xlsx"
OUT_AF = "africa_capacity_monthly_complete.xlsx"
START = "2021-01-01"
FUELS = cf.FUEL_ORDER
LT = "long_term_energy.xlsx"
EMBER_SRC = "Ember yearly electricity data: installed capacity by fuel (year-end, FALLBACK)"
EMBER_URL = "https://ember-energy.org/data/yearly-electricity-data/"
# Ember capacity variable per fuel column (LONG_TERM_ENERGY.parse_ember): Oil = Ember 'Other Fossil', Bioenergy = 'Bioenergy',
# Other = 'Other Renewables' (geothermal, marine ...)
EMBER_VAR = {"Hydro": "Cap_Hydro_GW", "Gas": "Cap_Gas_GW", "Coal": "Cap_Coal_GW", "Nuclear": "Cap_Nuclear_GW", "Wind": "Cap_Wind_GW",
             "Solar": "Cap_Solar_GW", "Oil": "Cap_Oil_GW", "Bioenergy": "Cap_Bio_GW", "Other": "Cap_OtherRen_GW"}

# name, ISO3, regions
COUNTRIES = [
    ("Israel", "ISR", ("EM",)), ("Egypt", "EGY", ("EM", "AF")), ("Cyprus", "CYP", ("EM",)), ("Jordan", "JOR", ("EM",)),
    ("Lebanon", "LBN", ("EM",)),
    ("Ghana", "GHA", ("AF",)), ("Nigeria", "NGA", ("AF",)), ("South Africa", "ZAF", ("AF",)), ("Morocco", "MAR", ("AF",)),
    ("Cameroon", "CMR", ("AF",)),
    ("Senegal", "SEN", ("AF",)), ("Cote d'Ivoire", "CIV", ("AF",)), ("Mauritania", "MRT", ("AF",)), ("Mali", "MLI", ("AF",)),
    ("Burkina Faso", "BFA", ("AF",)), ("Guinea", "GIN", ("AF",)), ("Sierra Leone", "SLE", ("AF",)), ("Liberia", "LBR", ("AF",)),
    ("Benin", "BEN", ("AF",)), ("Togo", "TGO", ("AF",)), ("Niger", "NER", ("AF",)), ("Gambia", "GMB", ("AF",)),
    ("Guinea-Bissau", "GNB", ("AF",)), ("Cape Verde", "CPV", ("AF",)), ("Equatorial Guinea", "GNQ", ("AF",)),
]
REGION_NAME = {"EM": "East Med", "AF": "Africa"}

# Official sources tried and not usable (Tried sheet): country, source, result
TRIED = [
    ("Nigeria", "NERC quarterly report (2026 Q2), nerc.gov.ng", "reachable; gives AVERAGE AVAILABLE capacity of the 28 grid plants and an installed share in "
                                                                "percent only, not installed MW by fuel; no installed table"),
    ("Egypt", "EEHC / MOEE (eehc.gov.eg, moee.gov.eg)", "eehc.gov.eg answers a 600-byte stub page with no document links; moee.gov.eg does not "
                                                      "resolve from Actions; CAPMAS / MOEE ST_* tables used by EGYPT_GAS.py hold gas and energy balance "
                                                      "rows, no installed capacity by fuel"),
    ("Israel", "Electricity Authority 'State of the Electricity Sector' report 2024/25 (gov.il BlobFolder, curl_cffi)", "downloads (124 pages); "
              "installed capacity by fuel is a chart, one year; earlier-year reports at the same URL pattern return 404"),
    ("Cyprus", "CERA annual report 2024", "cera.org.cy/.../ektheseis/2024_en.pdf returns 404 (2022 and 2023 exist)"),
]

_cache = {}


def sheet(data_dir, file, name):
    key = (data_dir, file, name)
    if key not in _cache:
        p = os.path.join(data_dir, file)
        d = None
        if os.path.exists(p):
            try:
                d = pd.read_excel(p, sheet_name=name)
            except Exception as e:  # noqa: BLE001
                print(f"  {file}/{name}: {type(e).__name__}: {e}")
        _cache[key] = d
    return _cache[key]


def stocks(years, vals, scale=1.0):
    """Year-end stocks (calendar year Y) -> pandas Series dated 1 January of Y+1, MW."""
    idx = pd.to_datetime([f"{int(y) + 1}-01-01" for y in years])
    return pd.Series(np.asarray(vals, dtype=float) * scale, index=idx)


def ember(data_dir, iso, fuel):
    d = sheet(data_dir, LT, "Power")
    var = EMBER_VAR.get(fuel)
    if d is None or var is None:
        return None
    x = d[(d["iso3"] == iso) & (d["variable"] == var)]
    if x.empty:
        return None
    return cf.Series(stocks(x["year"], x["value"], 1000.0), "EMBER", None, EMBER_SRC, EMBER_URL,
                     "year-end stock dated 1 January of the next year")


# ------------------------------------------------------------------ official annual layers
JORDAN_SRC = ("NEPCO annual reports: 'Available Capacity of Generating Plants' on the transmission network (MW, year-end; "
              "jordan_power_annual.xlsx)")
JORDAN_URL = "https://www.nepco.com.jo/en/AnnualReports.aspx"


def jordan(data_dir):
    d = sheet(data_dir, "jordan_power_annual.xlsx", "Annual")
    if d is None or "Cap_Steam_MW" not in d:
        return {}
    d = d.dropna(subset=["Cap_Steam_MW"])
    if d.empty:
        return {}
    y = d["Year"]
    col = lambda c: pd.to_numeric(d[c], errors="coerce")  # noqa: E731
    mk = lambda v, note="": cf.Series(stocks(y, v), "A", None, JORDAN_SRC, JORDAN_URL, note)  # noqa: E731
    return {
        "Gas": mk(col("Cap_GasTurbine_MW") + col("Cap_CombinedCycle_MW"), "combined cycle + gas turbines"),
        "Oil": mk(col("Cap_Diesel_MW"), "NEPCO 'Diesel*': engines running on heavy fuel, diesel and natural gas"),
        "Hydro": mk(col("Cap_Hydro_MW")), "Wind": mk(col("Cap_Wind_MW")),
        "Solar": mk(col("Cap_Solar_MW"), "transmission-connected projects only"),
        "Other": mk(col("Cap_Steam_MW"), "steam units: Aqaba thermal (dual-fuel) plus, from 2022, Attarat oil shale; not split by fuel in the table"),
    }


def annual_workbook(data_dir, file, src, url, cols, notes=None):
    """Official annual capacity workbook written by africa/AFRICA_CAPACITY_ANNUAL.py: sheet 'Monthly' (date = 1 January of the year
    whose END the stock is, <Fuel>_MW). cols = {fuel: column}. Returns {fuel: Series} (empty when the workbook is missing)."""
    d = sheet(data_dir, file, "Monthly")
    if d is None:
        return {}
    out = {}
    for fuel, c in cols.items():
        if c in d and d[c].notna().any():
            x = d.dropna(subset=[c])
            out[fuel] = cf.Series(stocks(x["Year"], x[c]), "A", None, src, url, (notes or {}).get(fuel, ""))
    return out


def ghana(data_dir):
    return annual_workbook(
        data_dir, "ghana_capacity_annual.xlsx",
        "Energy Commission of Ghana, National Energy Statistics, Table 3.1 grid installed capacity (MW, year-end; ghana_capacity_annual.xlsx)",
        "https://www.energycom.gov.gh/index.php/planning/energy-statistics",
        {"Hydro": "Hydro_MW", "Gas": "Gas_MW", "Solar": "Solar_MW"},
        {"Gas": "the Commission's 'Thermal' (gas, light crude oil, dual-fuel plants; not split by fuel)",
         "Solar": "the Commission's 'Other renewables' (solar and waste); grid-connected only, off-grid and distributed PV not included"})


def morocco(data_dir):
    return annual_workbook(
        data_dir, "morocco_capacity_annual.xlsx",
        "ONEE installed capacity by source, HCP Annuaire statistique du Maroc table 5-4 (MW, year-end; morocco_capacity_annual.xlsx)",
        "https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html",
        {"Hydro": "Hydro_MW", "Coal": "Coal_MW", "Oil": "Oil_MW", "Gas": "Gas_MW", "Other": "Other_MW", "Wind": "Wind_MW", "Solar": "Solar_MW"},
        {"Oil": "fuel-oil steam + diesel + gasoil units", "Gas": "combined cycle (Tahaddart, Ain Beni Mathar)",
         "Other": "gas turbines (mostly gasoil-fired, not split by fuel)"})


# Eskom Integrated Report 2025 (FY to 31 March 2025), 'Power station capacities at 31 March 2025' (pp. 145-146): installed capacity of
# the Eskom-owned stations. Koeberg (2 x ~967 MW, Jul 1984 - Nov 1985) and the three pumped-storage schemes (Drakensberg 1,000, Ingula
# 1,332 (last unit Feb 2017), Palmiet 400) are unchanged since their last unit was commissioned, so the table's commissioning dates
# give the same stock from then: two points per fuel (first full month after the last commissioning, and 31 March 2025). Nothing else
# from the report is used (the other fuels need a multi-year series; Ember is used for them).
ESKOM_SRC = "Eskom Integrated Report 2025, 'Power station capacities at 31 March 2025' (installed MW; commissioning dates in the same table)"
ESKOM_URL = "https://www.eskom.co.za/wp-content/uploads/2025/10/Eskom-integrated-report-2025.pdf"


def south_africa(data_dir):
    mk = lambda mw, first, note: cf.Series(  # noqa: E731
        pd.Series([mw, mw], index=pd.to_datetime([first, "2025-04-01"])), "A", None, ESKOM_SRC, ESKOM_URL, note)
    return {"Nuclear": mk(1934.0, "1985-12-01", "Koeberg 1 + 2 (1 x 970 + 1 x 964 MW); unchanged since the last unit, Nov 1985"),
            "PumpedStorage": mk(2732.0, "2017-03-01", "Drakensberg 1,000 + Ingula 1,332 + Palmiet 400 MW; unchanged since Ingula's last unit, Feb 2017")}


OFFICIAL = {"Jordan": jordan, "Ghana": ghana, "Morocco": morocco, "South Africa": south_africa}
OFFICIAL_NOTE = {}


def build_country(data_dir, name, iso):
    off = OFFICIAL[name](data_dir) if name in OFFICIAL else {}
    cfg = {}
    for fuel in FUELS:
        em = ember(data_dir, iso, fuel)
        prim = []
        if off.get(fuel) is not None:
            prim.append(off[fuel])
        if em is not None:
            prim.append(em)
        cfg[fuel] = cf.FuelCfg(primary=prim, ref=em if (off.get(fuel) is not None and em is not None) else None,
                               note=OFFICIAL_NOTE.get((name, fuel), ""))
    return cfg


def compare_breaks(name, cfg):
    """Informational rows: official stock against Ember's at the latest 1 January both hold."""
    rows = []
    for fuel, c in cfg.items():
        if len(c.primary) < 2 or c.primary[0].res == "EMBER":
            continue
        o, e = c.primary[0].clean(), c.primary[1].clean()
        common = [d for d in o.index.intersection(e.index) if e[d] > 0]
        if not common:
            continue
        d = common[-1]
        rows.append(dict(country=name, fuel=fuel, month=f"{d:%Y-%m}", kind="BASIS (official vs Ember, information)",
                         ratio=round(o[d] / e[d], 3),
                         detail=f"official {o[d]:,.0f} MW vs Ember {e[d]:,.0f} MW at {d:%Y-%m-%d}; the official series is used for this fuel, "
                                f"Ember is not mixed in"))
    return rows


def build_all(data_dir, end=None):
    end = pd.Timestamp(end) if end else pd.Timestamp(datetime.date.today().replace(day=1))
    results, extra = {}, []
    for name, iso, regions in COUNTRIES:
        cfg = build_country(data_dir, name, iso)
        if not any(s is not None and not s.clean().empty for c in cfg.values() for s in c.primary):
            print(f"  {name}: no source at all - left out")
            continue
        r = cf.fill_country(name, FUELS, START, end, cfg)
        results[name] = (iso, regions, r)
        extra += compare_breaks(name, cfg)
    return results, end, extra


def country_sheet(r):
    df = r.values.copy()
    df.columns = [f"{c}_MW" for c in df.columns]
    df.insert(0, "Total_MW", cf.total(r.values))
    df.index.name = "date"
    return df.round(1)


COMMON_NOTE = {"Storage": "batteries: not in any source here; excluded from Total", "PumpedStorage": "excluded from Total"}


def write_region(data_dir, region, results, end, extra):
    out = os.path.join(data_dir, OUT_EM if region == "EM" else OUT_AF)
    members = [n for n, (_i, regs, _r) in results.items() if region in regs]
    sheets, flag_sheets, basis, breaks, cov_rows = {}, {}, [], [], []
    for n in members:
        iso, regs, r = results[n]
        sheets[n[:31]] = country_sheet(r)
        fl = r.flags.copy()
        fl.index.name = "date"
        flag_sheets[f"{n} flags"[:31]] = fl
        basis.extend(r.basis)
        breaks.extend(r.breaks)
        c = cf.coverage(r.flags)
        c.name = n
        cov_rows.append(c)
    coverage = pd.DataFrame(cov_rows)
    main = [f for f in FUELS if f not in ("PumpedStorage", "Storage")]
    coverage["NONE main fuels"] = [float((results[n][2].flags[main] == "NONE").values.mean() * 100) for n in members]
    coverage = coverage.round(1)
    coverage.index.name = "country"
    tot = sum(results[n][2].values.fillna(0) for n in members)
    anyv = sum(results[n][2].values.notna().astype(int) for n in members)
    tot = tot.where(anyv > 0)
    reg = tot.copy()
    reg.columns = [f"{c}_MW" for c in reg.columns]
    reg.insert(0, "Total_MW", cf.total(tot))
    reg.index.name = "date"
    reg = reg.round(1)
    allflags = pd.concat([results[n][2].flags for n in members], axis=1)

    def share(group):
        return allflags.isin(group).sum(axis=1) / allflags.shape[1] * 100

    obs = share(cf.OBSERVED)
    none = share(["NONE"])
    held = share(["HELD_LATEST"])
    ember_sh = share(["FALLBACK_EMBER"])
    cbm = pd.DataFrame({"Observed official (ANCHOR_A / RAW)": obs, "Official interpolated / held between points": 100 - obs - none - held - ember_sh,
                        "Official held after last point": held, "Ember fallback (FALLBACK_EMBER)": ember_sh,
                        "No source (NONE)": none}).round(1)
    cbm.index.name = "date"
    basis_df = pd.DataFrame(basis)
    basis_df["caveat"] = [(b["caveat"] + (" " if b["caveat"] else "") + COMMON_NOTE.get(b["fuel"], "")).strip() for b in basis]
    br = pd.DataFrame(breaks) if breaks else pd.DataFrame(columns=["country", "fuel", "month", "kind", "ratio", "detail"])
    br = pd.concat([br, pd.DataFrame([e for e in extra if e["country"] in members])], ignore_index=True)
    tried = pd.DataFrame([t for t in TRIED if t[0] in members], columns=["country", "source", "result"]).set_index("country")
    off_c = [n for n in members if (basis_df[(basis_df.country == n)]["rule"] == "ANNUAL").any()]
    notes = [
        f"{REGION_NAME[region]} installed generating capacity - one gap-free monthly dataset per country and fuel (MW)", "",
        "What this is",
        "Built by middle_east/capacity/EAST_MED_AFRICA_CAPACITY_COMPLETE.py with the generic rules in capacity_fill.py. No downloads: official annual "
        "capacity workbooks where a pull exists, otherwise Ember yearly installed capacity (long_term_energy.xlsx), flagged FALLBACK_EMBER.",
        f"Span: {START[:7]} to {end:%Y-%m}, monthly, MW. A row dated the 1st of a month is the stock at the END of the previous month. Year-end "
        "annual stocks are dated 1 January of the next year. Total_MW = all fuels except PumpedStorage and Storage.",
        "Countries with an official annual layer: " + (", ".join(off_c) if off_c else "none") + ". All other country-fuels are Ember (fallback).",
        "", "Fill rules (per country and fuel, ONE source basis for the whole span)",
        "ANCHOR_A: an official annual stock at its own date. INTERP_A: wind, solar, batteries, bioenergy linear between annual stocks. HELD_A: every "
        "other fuel held at the last annual stock until the next one. HELD_LATEST: after the last official point the stock is held flat to the latest "
        "month, no trend. FALLBACK_EMBER: Ember yearly capacity (every month of the series is flagged this way, interpolated and held months included). "
        "NONE: no source holds the fuel, left blank. No capacity is guessed.",
        "", "Sheets",
        f"Basis (source, rule per country and fuel), Breaks (official vs Ember at the latest common year, other level breaks), Tried (official sources "
        f"checked and not usable), Coverage, Coverage by month, {REGION_NAME[region]} total, one sheet per country and '<country> flags'.",
        "", "Caveats",
        "Ember's yearly capacity comes from IRENA / national statistics and is rounded; 'Oil' is Ember's 'Other Fossil', 'Other' is 'Other Renewables'. "
        "Ember's Solar can include distributed PV where an official transmission-network series does not, and Ember lists no pumped storage or batteries (blank, NONE). A fuel Ember has no row for (e.g. coal or nuclear in a country without any) is NONE, not zero. "
        "A held or interpolated value is not an observation.",
    ]
    if region == "AF":
        notes.append("Egypt is built once and appears in both the East Med and the Africa workbook.")
    titles = ["What this is", "Fill rules (per country and fuel, ONE source basis for the whole span)", "Sheets", "Caveats"]
    all_sheets = {"Basis": basis_df.set_index("country"), "Breaks": br.set_index("country"), "Tried": tried, "Coverage": coverage,
                  "Coverage by month": cbm, f"{REGION_NAME[region]} total": reg}
    all_sheets.update(sheets)
    all_sheets.update(flag_sheets)
    xlsx_notes.write_workbook(out, all_sheets, notes, titles)
    print(f"{out}: {len(sheets)} countries, {START[:7]} - {end:%Y-%m}")
    print(coverage[["OBSERVED", "FILLED", "FALLBACK_EMBER", "NONE", "NONE main fuels"]].to_string())
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--end", default=None)
    ap.add_argument("--no-charts", action="store_true")
    args = ap.parse_args()
    results, end, extra = build_all(args.data_dir, args.end)
    outs = [write_region(args.data_dir, reg, results, end, extra) for reg in ("EM", "AF")]
    if not args.no_charts:
        import add_charts
        for o in outs:
            add_charts.add_charts(o)


# ---------------------------------------------------------------- master wiring (East Med and Africa masters)
OFFICIAL_URL = {"Jordan": JORDAN_URL, "Ghana": "https://www.energycom.gov.gh/index.php/planning/energy-statistics",
                "Morocco": "https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html", "South Africa": ESKOM_URL}
OFFICIAL_SHORT = {"Jordan": "NEPCO annual reports", "Ghana": "Energy Commission National Energy Statistics (Table 3.1)",
                  "Morocco": "ONEE via HCP Annuaire statistique du Maroc (table 5-4)",
                  "South Africa": "Eskom Integrated Report 2025 (Koeberg, pumped storage)"}


def source_line(path, country, workbook):
    """(text, url) for a country's capacity chart: which fuels are official annual and which Ember fallback, from the Basis sheet."""
    try:
        b = pd.read_excel(path, sheet_name="Basis")
        b = b[b["country"] == country]
        off = [f for f, r in zip(b["fuel"], b["rule"]) if str(r).startswith("ANNUAL")]
        emb = [f for f, r in zip(b["fuel"], b["rule"]) if str(r).startswith("FALLBACK_EMBER")]
    except Exception:  # noqa: BLE001
        off, emb = [], []
    parts = []
    if off:
        parts.append(f"official annual - {OFFICIAL_SHORT.get(country, 'national source')} for {', '.join(off)}")
    if emb:
        parts.append(f"Ember yearly installed capacity (FALLBACK, year-end) for {', '.join(emb)}")
    txt = (f"{country} installed capacity, monthly series filled from annual stocks: " + "; ".join(parts) if parts else
           f"{country} installed capacity: no source") + (
        f". Wind, solar, bioenergy linear between annual stocks, other fuels held, flat after the last stock; every filled cell is flagged in "
        f"{workbook}")
    return txt, OFFICIAL_URL.get(country, EMBER_URL) if off else EMBER_URL


def cap_spec(sheet_name, title):
    def f(p):
        import add_charts
        d = add_charts.by_date(add_charts.read(p, sheet_name), "date")
        d = d[d.index >= START]
        g = add_charts.capacity_groups(pd.DataFrame({fu: d.get(f"{fu}_MW") for fu in add_charts.CAPACITY_FUELS}, index=d.index)) / 1000.0
        return [add_charts.spec("Capacity", g, title, "GW installed", "stacked_bar", "%Y-%m")]
    return f


def _full_months(daily, min_share=0.9):
    """Daily frame (index = date) -> (monthly mean, share of days present) for months with >= min_share of their days."""
    m = daily.resample("MS")
    mean, n = m.mean(), m.count().max(axis=1)
    ok = n >= min_share * mean.index.days_in_month
    last = daily.index.max()
    ok &= (mean.index + pd.offsets.MonthEnd(0)) <= last
    return mean[ok]


def generation_frames(data_dir, region):
    """Monthly GWh by capacity_factors fuel group for the countries of the region with a monthly generation feed:
    East Med: Cyprus (TSOC daily MW); Africa: Ghana (Energy Commission WEM, daily GWh), Cameroon (ARSEL plant injections, 2025),
    Morocco (monthly gas-fired ESTIMATE from Tarifa gas). Months must be complete; nothing is filled."""
    out = {}
    if region == "EM":
        d = sheet(data_dir, "cyprus_generation_mix_daily.xlsx", "Cyprus")
        if d is not None:
            d = d.set_index(pd.to_datetime(d["date"])).drop(columns="date").apply(pd.to_numeric, errors="coerce")
            mw = _full_months(d)
            hrs = pd.Series(mw.index.days_in_month * 24.0, index=mw.index)
            gwh = mw.mul(hrs, axis=0) / 1000.0
            out["Cyprus"] = pd.DataFrame({"Wind": gwh["wind"], "Solar": gwh["solar"], "Other": gwh["oil"] + gwh["biomass"]})
    else:
        d = sheet(data_dir, "ghana_wem_weekly_generation_daily.xlsx", "Daily")
        if d is not None:
            d = d.set_index(pd.to_datetime(d["date"]))[["Hydro_GWh", "Thermal_GWh", "Solar_GWh"]].apply(pd.to_numeric, errors="coerce")
            m = d.resample("MS").sum(min_count=1)
            n = d.resample("MS").count().max(axis=1)
            ok = (n >= 0.9 * m.index.days_in_month) & ((m.index + pd.offsets.MonthEnd(0)) <= d.index.max())
            m = m[ok]
            out["Ghana"] = pd.DataFrame({"Hydro": m["Hydro_GWh"], "Gas": m["Thermal_GWh"], "Solar": m["Solar_GWh"]})
        d = sheet(data_dir, "cameroon_arsel_energy_balance.xlsx", "Plant injections")
        if d is not None:
            d = d.set_index(pd.to_datetime(d["month"])).drop(columns="month").apply(pd.to_numeric, errors="coerce") / 1000.0
            hyd = [c for c in ("Songloulou", "Edea 3", "Lagdo", "Nachtigal", "Memve'ele", "Lom Pangar") if c in d]
            oth = [c for c in ("Limbe HFO", "Oyomabang", "Dibamba") if c in d]
            out["Cameroon"] = pd.DataFrame({"Hydro": d[hyd].sum(axis=1, min_count=1), "Gas": d["Kribi (gas)"] if "Kribi (gas)" in d else np.nan,
                                           "Other": d[oth].sum(axis=1, min_count=1)}).dropna(how="all")
        d = sheet(data_dir, "morocco_gas_daily.xlsx", "Estimate")
        if d is not None and "Gas_fired_GWh_ESTIMATE" in d:
            d = d.set_index(pd.to_datetime(d["month"]))["Gas_fired_GWh_ESTIMATE"].apply(pd.to_numeric, errors="coerce").dropna()
            out["Morocco"] = pd.DataFrame({"Gas": d})
    return {k: v[v.index >= START] for k, v in out.items() if not v.empty}


# ---------------------------------------------------------------- charts (registered in add_charts.REGISTRY)
CHART_GROUPS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]


def chart_specs(path):
    import add_charts
    names = pd.ExcelFile(path).sheet_names
    tot = next(n for n in names if n.endswith(" total"))
    specs = []
    eu = add_charts.by_date(add_charts.read(path, tot), "date")
    g = pd.DataFrame({f: eu.get(f"{f}_MW") for f in CHART_GROUPS}, index=eu.index) / 1000.0
    specs.append(add_charts.spec(tot, g, f"{tot[:-6]} installed capacity, complete monthly series (official annual where reachable, Ember fallback elsewhere)",
                                 "GW installed", "stacked_bar"))
    cbm = add_charts.by_date(add_charts.read(path, "Coverage by month"), "date")
    specs.append(add_charts.spec("Coverage", cbm, "Share of country-fuel-months observed, interpolated, Ember fallback or without source",
                                 "% of country-fuel cells", "stacked_bar"))
    for n in names:
        if n in ("Units", "Basis", "Breaks", "Tried", "Coverage", "Coverage by month", tot) or n.endswith(" flags") or n.startswith("Chart"):
            continue
        d = add_charts.by_date(add_charts.read(path, n), "date")
        gg = pd.DataFrame({f: d.get(f"{f}_MW") for f in CHART_GROUPS}, index=d.index) / 1000.0
        gg = gg.dropna(how="all", axis=1)
        if gg.empty:
            continue
        specs.append(add_charts.spec(n, gg, f"{n} installed capacity, complete monthly series", "GW installed", "stacked_bar"))
    return specs


if __name__ == "__main__":
    main()
